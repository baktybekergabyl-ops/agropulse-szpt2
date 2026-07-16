import html
import re
from pathlib import Path
from urllib.parse import urljoin
from urllib.request import Request, urlopen

from .config import BASE_URL, REQUEST_TIMEOUT, SOURCE_PAGE, USER_AGENT
from .models import Release, SourceFile
from .parser import extract_date


def _fetch(url: str) -> bytes:
    request = Request(url, headers={"User-Agent": USER_AGENT})
    with urlopen(request, timeout=REQUEST_TIMEOUT) as response:
        return response.read()


def _clean_html(value: str) -> str:
    text = re.sub(r"<[^>]+>", " ", value)
    return " ".join(html.unescape(text).split())


def _release_extension(text: str) -> str:
    return ".xlsx" if re.search(r"\bxlsx\b", text, re.I) else ".xls"


def discover_releases(session=None) -> list[Release]:
    # session is kept for backward compatibility with older callers.
    page = _fetch(SOURCE_PAGE).decode("utf-8", errors="ignore")
    found: dict[str, Release] = {}
    link_pattern = re.compile(
        r'<a[^>]+href=["\'](?P<href>[^"\']*/api/iblock/element/[^"\']*/file/[^"\']*)["\'][^>]*>'
        r"(?P<title>.*?)</a>",
        re.I | re.S,
    )
    for match in link_pattern.finditer(page):
        href = match.group("href")
        url = urljoin(BASE_URL, href)
        title = _clean_html(match.group("title"))
        context = _clean_html(page[max(0, match.start() - 1200): match.end() + 1200])
        observation_date = extract_date(title) or extract_date(context)
        if not observation_date:
            continue
        if "социально-значимые" not in context.casefold() and "социально значимые" not in context.casefold():
            continue
        extension = _release_extension(context or title)
        candidate = Release(title or context, observation_date, url, extension)
        existing = found.get(url)
        if existing is None or (
            candidate.extension == ".xlsx" and existing.extension != ".xlsx"
        ) or len(candidate.title) > len(existing.title):
            found[url] = candidate
    releases = sorted(found.values(), key=lambda item: item.observation_date, reverse=True)
    if not releases:
        raise RuntimeError("На странице БНС не найдены еженедельные релизы.")
    return releases


def download_release(
    release: Release,
    raw_dir: Path,
    session=None,
) -> SourceFile:
    if release.extension != ".xlsx":
        raise RuntimeError(
            f"Релиз за {release.observation_date:%d.%m.%Y} опубликован "
            "в формате .xls; требуется конвертация в .xlsx."
        )
    raw_dir.mkdir(parents=True, exist_ok=True)
    path = raw_dir / f"bns_prices_{release.observation_date.isoformat()}.xlsx"
    if not path.exists() or path.stat().st_size == 0:
        path.write_bytes(_fetch(release.download_url))
    return SourceFile(release.observation_date, path)


def download_release_file(
    release: Release,
    raw_dir: Path,
    session=None,
) -> Path:
    raw_dir.mkdir(parents=True, exist_ok=True)
    if release.extension == ".xlsx":
        path = raw_dir / f"bns_prices_{release.observation_date.isoformat()}.xlsx"
    else:
        xls_dir = raw_dir / "xls"
        xls_dir.mkdir(parents=True, exist_ok=True)
        path = xls_dir / f"bns_prices_{release.observation_date.isoformat()}.xls"
    if path.exists() and path.stat().st_size > 0:
        return path
    path.write_bytes(_fetch(release.download_url))
    return path


def sync_price_archive(
    raw_dir: Path,
    years: set[int] | None = None,
    session=None,
) -> list[Path]:
    years = years or {2024, 2025, 2026}
    releases = [
        release for release in discover_releases()
        if release.observation_date.year in years
    ]
    downloaded: list[Path] = []
    for release in sorted(releases, key=lambda item: item.observation_date):
        downloaded.append(download_release_file(release, raw_dir))
    return downloaded


def sync_updates(
    raw_dir: Path,
    local_sources: list[SourceFile],
    session=None,
) -> list[SourceFile]:
    releases = discover_releases()
    local_dates = {item.observation_date for item in local_sources}
    if local_sources:
        newest_local = max(local_dates)
        pending = [
            release for release in releases
            if release.observation_date > newest_local
            and release.observation_date not in local_dates
        ]
    else:
        pending = [
            release for release in releases[:2]
            if release.observation_date not in local_dates
        ]
    return [
        download_release(release, raw_dir)
        for release in sorted(pending, key=lambda item: item.observation_date)
        if release.extension == ".xlsx"
    ]
