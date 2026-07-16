import argparse
import logging
import sys

import requests

from .config import (
    BRIEF_TEMPLATE_PATH,
    COMPARISON_TEMPLATE_PATH,
    OUTPUT_DIR,
    RAW_DIR,
    TEMPLATE_PATH,
)
from .downloader import sync_updates
from .package import create_brief, create_comparison_workbook
from .parser import discover_source_files, parse_snapshot
from .report import create_report


def build_parser() -> argparse.ArgumentParser:
    return argparse.ArgumentParser(
        description="Локальный анализ файлов БНС РК из data/raw."
    )


def main() -> None:
    build_parser().parse_args()
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    try:
        sources = discover_source_files(RAW_DIR)
        logging.info("Проверяю обновления на сайте БНС...")
        with requests.Session() as session:
            downloaded = sync_updates(RAW_DIR, sources, session)
        if downloaded:
            logging.info(
                "Скачано новых релизов: %s",
                ", ".join(item.observation_date.isoformat() for item in downloaded),
            )
            sources = discover_source_files(RAW_DIR)
        else:
            logging.info("Новых релизов нет.")
        if not sources:
            raise FileNotFoundError(
                "В data/raw нет исходных файлов БНС в формате .xlsx."
            )
        current = parse_snapshot(sources[0])
        previous = parse_snapshot(sources[1]) if len(sources) >= 2 else None
        if previous is None:
            logging.warning(
                "Найден один период: отчет будет создан без предыдущей цены."
            )
        report = create_report(
            TEMPLATE_PATH, current, previous, OUTPUT_DIR
        )
        comparison = create_comparison_workbook(
            COMPARISON_TEMPLATE_PATH, current, OUTPUT_DIR
        )
        brief = create_brief(
            BRIEF_TEMPLATE_PATH, current, previous, OUTPUT_DIR
        )
        logging.info("Текущий источник: %s", current.source.path.name)
        if previous:
            logging.info("Предыдущий источник: %s", previous.source.path.name)
        logging.info("Готово: %s", report)
        logging.info("Готово: %s", comparison)
        logging.info("Готово: %s", brief)
    except Exception as error:
        logging.error("%s", error)
        sys.exit(1)
