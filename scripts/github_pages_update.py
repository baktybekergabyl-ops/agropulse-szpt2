from __future__ import annotations

import argparse
import json
import shutil
import sys
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo


ROOT = Path(__file__).resolve().parents[1]
LOCAL_TZ = ZoneInfo("Asia/Qyzylorda")

sys.path.insert(0, str(ROOT))

from scripts import weekly_update  # noqa: E402


def should_run(mode: str, now: datetime) -> bool:
    if mode == "all":
        return True
    if mode == "prices":
        if now.weekday() != 4 or now.hour < 12:
            print("Skip: weekly prices update is intended for Friday after 12:00 Asia/Qyzylorda.")
            return False
        return True
    if mode == "monthly":
        if now.day < 15 or now.hour < 15:
            print("Skip: monthly sources check is intended after the 15th day at 15:00 Asia/Qyzylorda.")
            return False
        return True
    raise ValueError(f"Unknown mode: {mode}")


def sync_root_public_files() -> None:
    dashboard_dir = ROOT / "dashboard"
    data_dir = ROOT / "data"
    public_data_dir = ROOT / "data"

    for name in ["index.html", "app.js", "styles.css"]:
        source = dashboard_dir / name
        if source.exists():
            shutil.copy2(source, ROOT / name)

    generated_json = dashboard_dir / "data" / "bns.json"
    if not generated_json.exists():
        raise FileNotFoundError(f"Generated JSON not found: {generated_json}")
    public_data_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy2(generated_json, data_dir / "bns.json")


def read_existing_public_json() -> dict | None:
    path = ROOT / "data" / "bns.json"
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        print(f"Existing public JSON not reused: {exc}")
        return None


def preserve_operational_stocks(previous: dict | None) -> None:
    if not previous:
        return
    previous_stocks = previous.get("stocks")
    if not isinstance(previous_stocks, dict) or not previous_stocks.get("enabled"):
        return

    generated_path = ROOT / "dashboard" / "data" / "bns.json"
    if not generated_path.exists():
        return
    current = json.loads(generated_path.read_text(encoding="utf-8"))
    current_stocks = current.get("stocks")
    if isinstance(current_stocks, dict) and current_stocks.get("enabled"):
        return

    preserved_stocks = dict(previous_stocks)
    preserved_stocks["cacheStatus"] = "preserved_from_previous_public_json"
    current["stocks"] = preserved_stocks

    previous_products = {
        item.get("id"): item.get("operationalStocksMio")
        for item in previous.get("products", [])
        if isinstance(item, dict) and item.get("id") and item.get("operationalStocksMio")
    }
    for item in current.get("products", []):
        if not isinstance(item, dict):
            continue
        preserved = previous_products.get(item.get("id"))
        if preserved:
            item["operationalStocksMio"] = preserved

    meta = current.setdefault("meta", {})
    previous_meta = previous.get("meta", {}) if isinstance(previous.get("meta"), dict) else {}
    meta["mioStocksUpdated"] = previous_meta.get("mioStocksUpdated") or preserved_stocks.get("updated")
    meta["mioStocksCacheStatus"] = "preserved_from_previous_public_json"

    generated_path.write_text(json.dumps(current, ensure_ascii=False, indent=2), encoding="utf-8")


def run_update(mode: str, skip_download: bool) -> None:
    previous_public_json = read_existing_public_json()
    if not skip_download:
        if mode in {"all", "prices"}:
            weekly_update.try_download_from_bns()
        if mode in {"all", "monthly"}:
            weekly_update.sync_current_foreign_trade()
            weekly_update.sync_industrial_production_sources()
        weekly_update.convert_downloaded_xls()
    else:
        print("Download step skipped; rebuilding from repository seed files.")

    weekly_update.build_dashboard()
    preserve_operational_stocks(previous_public_json)
    sync_root_public_files()


def main() -> None:
    parser = argparse.ArgumentParser(description="Update AgroPulse GitHub Pages data.")
    parser.add_argument("--mode", choices=("prices", "monthly", "all"), default="all")
    parser.add_argument("--skip-download", action="store_true")
    args = parser.parse_args()

    now = datetime.now(LOCAL_TZ)
    print(f"AgroPulse GitHub Pages update: mode={args.mode}; local_time={now:%Y-%m-%d %H:%M:%S %Z}")
    if not should_run(args.mode, now):
        return
    run_update(args.mode, args.skip_download)


if __name__ == "__main__":
    main()
