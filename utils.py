from __future__ import annotations

import csv
import getpass
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable

from config import FIELD_KEYS


def sanitize_excel_text(value: str | None) -> str:
    if value is None:
        return ""
    clean = str(value).replace("\x00", "").strip()
    if clean.startswith(("=", "+", "-", "@")):
        clean = "'" + clean
    return clean


def sanitize_payload(payload: Dict[str, str]) -> Dict[str, str]:
    return {key: sanitize_excel_text(payload.get(key, "")) for key in FIELD_KEYS}


def validate_required(payload: Dict[str, str], required_keys: Iterable[str]) -> list[str]:
    errors = []
    for key in required_keys:
        if not str(payload.get(key, "")).strip():
            errors.append(f"Champ obligatoire manquant : {key}")
    return errors


def append_history(log_file: Path, action: str, row_number: int, identifier: str, sheet_name: str) -> None:
    exists = log_file.exists()
    log_file.parent.mkdir(parents=True, exist_ok=True)
    with log_file.open("a", newline="", encoding="utf-8") as file:
        writer = csv.writer(file)
        if not exists:
            writer.writerow(["timestamp", "user", "action", "sheet", "row", "number"])
        writer.writerow([
            datetime.now().isoformat(timespec="seconds"),
            getpass.getuser(),
            action,
            sheet_name,
            row_number,
            identifier,
        ])
