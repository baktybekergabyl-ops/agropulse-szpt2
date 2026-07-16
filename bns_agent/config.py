from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = PROJECT_ROOT / "data"
RAW_DIR = DATA_DIR / "raw"
OUTPUT_DIR = PROJECT_ROOT / "output"
TEMPLATE_PATH = PROJECT_ROOT / "templates" / "bns_template.xlsx"
COMPARISON_TEMPLATE_PATH = (
    PROJECT_ROOT / "templates" / "Сравнение цен в разбивке по мес на 24.06.xlsx"
)
BRIEF_TEMPLATE_PATH = (
    PROJECT_ROOT / "templates" / "Справка по росту цен СЗПТ 24.06.docx"
)

SOURCE_PAGE = (
    "https://stat.gov.kz/ru/industries/economy/prices/"
    "spreadsheets/?name=19060"
)
BASE_URL = "https://stat.gov.kz"
REQUEST_TIMEOUT = 60
USER_AGENT = (
    "Mozilla/5.0 (compatible; BNS-Weekly-Price-Agent/2.0; "
    "+https://stat.gov.kz/)"
)
