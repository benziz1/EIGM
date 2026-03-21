from __future__ import annotations

import csv
import getpass
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, List

from config import COLUMNS_BY_KEY, FIELD_KEYS

HISTORY_HEADERS = [
    "timestamp",
    "user",
    "action",
    "sheet",
    "row",
    "number",
    "number_base",
    "changes",
]


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


def requirement_number_base(identifier: str) -> str:
    value = str(identifier or "").strip()
    if not value:
        return ""
    parts = value.split("_")
    if len(parts) >= 4:
        return "_".join(parts[:-1])
    return value


def describe_payload_changes(previous_payload: Dict[str, str], current_payload: Dict[str, str]) -> str:
    changes: List[str] = []
    for key in FIELD_KEYS:
        if key == "number":
            continue
        before = str(previous_payload.get(key, "") or "").strip()
        after = str(current_payload.get(key, "") or "").strip()
        if before != after:
            label = COLUMNS_BY_KEY[key].label
            changes.append(f"{label}: '{before or '∅'}' → '{after or '∅'}'")
    return " | ".join(changes) if changes else "Aucune différence métier détectée"


def append_history(
    log_file: Path,
    action: str,
    row_number: int,
    identifier: str,
    sheet_name: str,
    changes: str = "",
) -> None:
    exists = log_file.exists()
    log_file.parent.mkdir(parents=True, exist_ok=True)
    with log_file.open("a", newline="", encoding="utf-8") as file:
        writer = csv.writer(file)
        if not exists:
            writer.writerow(HISTORY_HEADERS)
        writer.writerow(
            [
                datetime.now().isoformat(timespec="seconds"),
                getpass.getuser(),
                action,
                sheet_name,
                row_number,
                identifier,
                requirement_number_base(identifier),
                changes,
            ]
        )


def read_history(log_file: Path, limit: int = 20, number_base_filter: str = "") -> List[Dict[str, str]]:
    if not log_file.exists():
        return []

    entries: List[Dict[str, str]] = []
    with log_file.open("r", newline="", encoding="utf-8") as file:
        reader = csv.DictReader(file)
        for row in reader:
            normalized = {header: row.get(header, "") for header in HISTORY_HEADERS}
            if number_base_filter and normalized["number_base"] != number_base_filter:
                continue
            entries.append(normalized)

    entries.sort(key=lambda item: item["timestamp"], reverse=True)
    return entries[:limit]
