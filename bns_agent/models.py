from dataclasses import dataclass, field
from datetime import date
from pathlib import Path


@dataclass(frozen=True)
class SourceFile:
    observation_date: date
    path: Path


@dataclass(frozen=True)
class Release:
    title: str
    observation_date: date
    download_url: str
    extension: str


@dataclass(frozen=True)
class PriceRecord:
    product: str
    region: str
    price: float


@dataclass
class Snapshot:
    source: SourceFile
    prices: dict[str, dict[str, float]] = field(default_factory=dict)
    annual_change: dict[str, float] = field(default_factory=dict)
    year_change: dict[str, float] = field(default_factory=dict)
    week_change: dict[str, float] = field(default_factory=dict)
