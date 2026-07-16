from __future__ import annotations

import argparse
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


def run_update(mode: str, skip_download: bool) -> None:
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
