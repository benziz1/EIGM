from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_WORKBOOK = BASE_DIR / "requirements_db.xlsx"
DEFAULT_SHEET = "V00"
DATA_START_ROW = 7
HEADER_ROWS = (5, 6)
NUMBER_PREFIX = "REQ"
NUMBER_SUFFIX_DEFAULT = 1
TYPE_FREE_OPTION = "<Saisie libre>"
MILESTONE_KEYS = ["sor", "pdr", "cdr", "tqr", "trr", "fai"]
MILESTONE_MARK = "X"


@dataclass(frozen=True)
class ColumnConfig:
    letter: str
    key: str
    label: str
    required: bool = False
    dropdown: bool = False


COLUMN_DEFINITIONS: List[ColumnConfig] = [
    ColumnConfig("A", "index", "#"),
    ColumnConfig("B", "number", "NUMBER", required=True),
    ColumnConfig("C", "french_name", "FRENCH NAME", required=True),
    ColumnConfig("D", "english_name", "ENGLISH NAME", required=True),
    ColumnConfig("E", "type", "TYPE", required=True, dropdown=True),
    ColumnConfig("F", "iadt", "IADT", dropdown=True),
    ColumnConfig("G", "chapter", "CHAPTER", dropdown=True),
    ColumnConfig("H", "category", "CATEGORY", dropdown=True),
    ColumnConfig("I", "sub_category", "SUB CATEGORY", dropdown=True),
    ColumnConfig("J", "config", "CONFIG", dropdown=True),
    ColumnConfig("K", "french_resume", "FRENCH RESUME"),
    ColumnConfig("L", "english_resume", "ENGLISH RESUME"),
    ColumnConfig("M", "applicability", "APPLICABILITY", required=True, dropdown=True),
    ColumnConfig("N", "remarks_n", "REMARKS"),
    ColumnConfig("O", "coverage", "COVERAGE", required=True, dropdown=True),
    ColumnConfig("P", "remarks_p", "REMARKS"),
    ColumnConfig("Q", "decision", "DECISION"),
    ColumnConfig("R", "remarks_r", "REMARKS"),
    ColumnConfig("S", "sor", "SOR"),
    ColumnConfig("T", "pdr", "PDR"),
    ColumnConfig("U", "cdr", "CDR"),
    ColumnConfig("V", "tqr", "TQR"),
    ColumnConfig("W", "trr", "TRR"),
    ColumnConfig("X", "fai", "FAI"),
]

COLUMNS_BY_KEY: Dict[str, ColumnConfig] = {cfg.key: cfg for cfg in COLUMN_DEFINITIONS}

FIELD_KEYS = [
    "number",
    "french_name",
    "english_name",
    "type",
    "iadt",
    "chapter",
    "category",
    "sub_category",
    "config",
    "french_resume",
    "english_resume",
    "applicability",
    "remarks_n",
    "coverage",
    "remarks_p",
    "remarks_r",
    "sor",
    "pdr",
    "cdr",
    "tqr",
    "trr",
    "fai",
]

DEFAULT_SELECT_OPTIONS = {
    "type": ["GEN", "AUT"],
    "iadt": ["DEMONSTRATION"],
    "applicability": ["APPLICABLE", "NON APPLICABLE"],
    "coverage": ["COMPLIANT", "NOT COMPLIANT", "PARTIAL"],
}

ARCHIVE_MARKER = "[ARCHIVED]"
LOG_FILE = BASE_DIR / "history_log.csv"

SECTION_LAYOUT = {
    "Bloc 1 — Identification de l'exigence": ["number", "french_name", "english_name", "type"],
    "Bloc 2 — Méthode / contexte": ["iadt", "chapter", "category", "sub_category", "config"],
    "Bloc 3 — Description": ["french_resume", "english_resume"],
    "Bloc 4 — Analyse": ["applicability", "remarks_n", "coverage", "remarks_p", "remarks_r"],
    "Bloc 5 — Jalons projet": MILESTONE_KEYS,
}
