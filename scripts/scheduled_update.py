from __future__ import annotations

import argparse
import subprocess
import sys
from datetime import datetime
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def run_weekly_update() -> None:
    subprocess.run(
        [sys.executable, str(ROOT / "scripts" / "weekly_update.py")],
        cwd=ROOT,
        check=True,
    )


def main() -> None:
    parser = argparse.ArgumentParser(description="Scheduled AgroPulse dashboard updater.")
    parser.add_argument(
        "--mode",
        choices=("prices", "monthly", "all"),
        default="all",
        help="prices: Friday price check; monthly: after 15th monthly source check; all: run now.",
    )
    args = parser.parse_args()
    now = datetime.now()
    print(f"AgroPulse scheduled update: mode={args.mode}; time={now:%Y-%m-%d %H:%M:%S}")

    if args.mode == "monthly" and now.day < 15:
        print("Monthly sources check skipped: today is before the 15th day of the month.")
        return

    if args.mode == "prices" and (now.weekday() != 4 or now.hour < 12):
        print("Weekly price check skipped: expected Friday after 12:00.")
        return

    run_weekly_update()


if __name__ == "__main__":
    main()
