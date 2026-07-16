from __future__ import annotations

import html
import json
import re
import subprocess
import sys
import zipfile
from datetime import date
from pathlib import Path
from urllib.parse import urljoin
from urllib.request import Request, urlopen

ROOT = Path(__file__).resolve().parents[1]
RAW_DIR = ROOT / "data" / "raw"
ZIP_PATH = ROOT / "dashboard-cloudflare-pages.zip"
PUBLIC_DIR = ROOT / "dashboard-cloudflare-pages"
BASE_URL = "https://stat.gov.kz"
TRADE_SOURCE_PAGE = (
    "https://stat.gov.kz/ru/industries/economy/foreign-market/spreadsheets/"
    "?year={year}&name=40113&type=spreadsheets"
)
TRADE_CURRENT_DIR = RAW_DIR / "foreign_trade" / "current"
INDUSTRIAL_SOURCE_PAGE = (
    "https://stat.gov.kz/ru/industries/business-statistics/"
    "stat-industrial-production/spreadsheets/"
)
INDUSTRIAL_DIR = RAW_DIR / "industrial_production"
USER_AGENT = "Mozilla/5.0 (compatible; AgroPulse-Dashboard/1.0; +https://stat.gov.kz/)"

sys.path.insert(0, str(ROOT))
try:
    from bns_agent.downloader import sync_price_archive, sync_updates  # noqa: E402
except Exception as exc:  # requests может отсутствовать во встроенном Python
    sync_updates = None
    DOWNLOAD_IMPORT_ERROR = exc
else:
    DOWNLOAD_IMPORT_ERROR = None
from bns_agent.parser import discover_source_files  # noqa: E402


def try_download_from_bns() -> None:
    if sync_updates is None:
        print(f"Автоскачивание с сайта БНС пропущено: {DOWNLOAD_IMPORT_ERROR}")
        return
    try:
        sources = discover_source_files(RAW_DIR)
        downloaded = sync_updates(RAW_DIR, sources)
        if downloaded:
            print("Скачано с сайта БНС:", ", ".join(item.path.name for item in downloaded))
        else:
            print("На сайте БНС новых недельных файлов не найдено.")
    except Exception as exc:
        print(f"Сайт БНС недоступен или файл не скачан: {exc}")


def sync_bns_price_archive() -> None:
    if sync_updates is None:
        return
    try:
        files = sync_price_archive(RAW_DIR, {2024, 2025, 2026})
        print(f"Архив цен БНС: найдено/скачано {len(files)} файлов за 2024-2026.")
    except Exception as exc:
        print(f"Архив цен БНС не обновлен: {exc}")


def fetch_url(url: str) -> bytes:
    request = Request(url, headers={"User-Agent": USER_AGENT})
    with urlopen(request, timeout=90) as response:
        return response.read()


def clean_html(value: str) -> str:
    text = re.sub(r"<[^>]+>", " ", value)
    return " ".join(html.unescape(text).split())


def extract_trade_period(title: str, fallback_year: int) -> str:
    match = re.search(r"\(([^)]*?\d{4}\s*г?\.?)\)", title)
    if match:
        return " ".join(match.group(1).split())
    return str(fallback_year)


def sync_current_foreign_trade() -> None:
    year = date.today().year
    page_url = TRADE_SOURCE_PAGE.format(year=year)
    try:
        page = fetch_url(page_url).decode("utf-8", errors="ignore")
        candidates: list[dict[str, str]] = []
        link_pattern = re.compile(
            r'<a[^>]+href=["\'](?P<href>[^"\']*/api/iblock/element/(?P<element>\d+)/file/ru/?)["\'][^>]*>'
            r"(?P<title>.*?)</a>",
            re.I | re.S,
        )
        for match in link_pattern.finditer(page):
            title = clean_html(match.group("title"))
            normalized = title.casefold()
            if not normalized.startswith("экспорт и импорт товаров рк по 4,6,10 знакам тн вэд еаэс"):
                continue
            href = match.group("href")
            candidates.append(
                {
                    "element": match.group("element"),
                    "title": title,
                    "url": urljoin(BASE_URL, href),
                    "period": extract_trade_period(title, year),
                }
            )
        if not candidates:
            print(f"Текущий файл внешней торговли БНС за {year} не найден на странице {page_url}.")
            return
        latest = candidates[0]
        TRADE_CURRENT_DIR.mkdir(parents=True, exist_ok=True)
        path = TRADE_CURRENT_DIR / f"foreign_trade_hs_{year}_{latest['element']}.xlsx"
        if not path.exists() or path.stat().st_size == 0:
            path.write_bytes(fetch_url(latest["url"]))
            print(f"Скачан текущий экс-имп БНС: {latest['title']}")
        else:
            print(f"Текущий экс-имп БНС уже есть: {path.name}")
        meta = {
            "year": year,
            "file": path.name,
            "title": latest["title"],
            "period": latest["period"],
            "url": latest["url"],
            "pageUrl": page_url,
            "element": latest["element"],
            "updatedAt": date.today().isoformat(),
        }
        (TRADE_CURRENT_DIR / f"latest_{year}.json").write_text(
            json.dumps(meta, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except Exception as exc:
        print(f"Текущий экс-имп БНС не обновлен: {exc}")


def sync_industrial_production_sources() -> None:
    try:
        page = fetch_url(INDUSTRIAL_SOURCE_PAGE).decode("utf-8", errors="ignore")
        link_pattern = re.compile(
            r'<a[^>]+href=["\'](?P<href>[^"\']*/api/iblock/element/(?P<element>\d+)/file/ru/?)["\'][^>]*>'
            r"(?P<title>.*?)</a>",
            re.I | re.S,
        )
        current_year = date.today().year
        monthly: list[dict[str, str]] = []
        annual: list[dict[str, str]] = []
        for match in link_pattern.finditer(page):
            title = clean_html(match.group("title"))
            normalized = title.casefold()
            if not normalized.startswith("основные показатели работы промышленности республики казахстан"):
                continue
            item = {
                "element": match.group("element"),
                "title": title,
                "url": urljoin(BASE_URL, match.group("href")),
            }
            if f"{current_year}г" in normalized and "январ" in normalized:
                monthly.append(item)
            elif f"{current_year - 1}г" in normalized and "январ" not in normalized and "декабрь" not in normalized:
                annual.append(item)
        INDUSTRIAL_DIR.mkdir(parents=True, exist_ok=True)

        def download_latest(kind: str, items: list[dict[str, str]], prefix: str) -> None:
            if not items:
                print(f"Промышленность БНС: файл {kind} не найден на странице.")
                return
            latest = items[0]
            path = INDUSTRIAL_DIR / f"{prefix}_{latest['element']}.xlsx"
            if not path.exists() or path.stat().st_size == 0:
                path.write_bytes(fetch_url(latest["url"]))
                print(f"Скачана промышленность БНС ({kind}): {latest['title']}")
            else:
                print(f"Промышленность БНС ({kind}) уже есть: {path.name}")
            meta = {
                "kind": kind,
                "file": path.name,
                "title": latest["title"],
                "url": latest["url"],
                "pageUrl": INDUSTRIAL_SOURCE_PAGE,
                "element": latest["element"],
                "updatedAt": date.today().isoformat(),
            }
            (INDUSTRIAL_DIR / f"latest_{kind}.json").write_text(
                json.dumps(meta, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )

        download_latest("monthly", monthly, "T-04-01-М_рус")
        download_latest("annual", annual, "Т-04-05-Г_рус")
    except Exception as exc:
        print(f"Промышленность БНС не обновлена: {exc}")


def convert_downloaded_xls() -> None:
    xls_dir = RAW_DIR / "xls"
    converted_dir = RAW_DIR / "converted"
    if not xls_dir.exists():
        return
    converted_dir.mkdir(parents=True, exist_ok=True)
    command = (
        "$ErrorActionPreference = 'Stop'; "
        f"$xlsDir = '{str(xls_dir)}'; "
        f"$convertedDir = '{str(converted_dir)}'; "
        "$excel = New-Object -ComObject Excel.Application; "
        "$excel.Visible = $false; $excel.DisplayAlerts = $false; "
        "try { "
        "Get-ChildItem -LiteralPath $xlsDir -Filter '*.xls' | ForEach-Object { "
        "$dest = Join-Path $convertedDir ($_.BaseName + '.xlsx'); "
        "if ((Test-Path -LiteralPath $dest) -and ((Get-Item -LiteralPath $dest).LastWriteTime -ge $_.LastWriteTime)) { return }; "
        "$wb = $excel.Workbooks.Open($_.FullName); "
        "$wb.SaveAs($dest, 51); "
        "$wb.Close($false); "
        "Write-Output $dest "
        "} "
        "} finally { $excel.Quit() | Out-Null; [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }"
    )
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", command],
            cwd=ROOT,
            text=True,
            capture_output=True,
            check=True,
        )
        converted = [line for line in result.stdout.splitlines() if line.strip()]
        if converted:
            print(f"Конвертировано .xls в .xlsx: {len(converted)}")
    except Exception as exc:
        print(f"Конвертация .xls пропущена: {exc}")


def build_dashboard() -> None:
    subprocess.run([sys.executable, str(ROOT / "scripts" / "build_dashboard_data.py")], cwd=ROOT, check=True)


def build_zip() -> None:
    if not PUBLIC_DIR.exists():
        raise FileNotFoundError(f"Нет папки для публикации: {PUBLIC_DIR}")
    if ZIP_PATH.exists():
        ZIP_PATH.unlink()
    with zipfile.ZipFile(ZIP_PATH, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for path in PUBLIC_DIR.rglob("*"):
            if path.is_file():
                zf.write(path, path.relative_to(PUBLIC_DIR))
    print(f"ZIP готов: {ZIP_PATH}")


def main() -> None:
    print("Обновление АгроПульс: только официальные данные БНС/Талдау, ZIP для Cloudflare")
    try_download_from_bns()
    sync_bns_price_archive()
    sync_current_foreign_trade()
    sync_industrial_production_sources()
    convert_downloaded_xls()
    build_dashboard()
    build_zip()
    print("Готово. Если Cloudflare без API-токена, загрузите ZIP вручную через New deployment.")


if __name__ == "__main__":
    main()
