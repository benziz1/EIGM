from __future__ import annotations

from copy import copy
from pathlib import Path
from typing import Dict, List

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.worksheet.worksheet import Worksheet

from config import (
    ARCHIVE_MARKER,
    COLUMN_DEFINITIONS,
    COLUMNS_BY_KEY,
    DATA_START_ROW,
    DEFAULT_SELECT_OPTIONS,
    FIELD_KEYS,
)


class ExcelHandlerError(Exception):
    """Raised when workbook operations fail."""


class ExcelHandler:
    def __init__(self, workbook_path: Path):
        self.workbook_path = Path(workbook_path)

    def _load_workbook(self):
        if not self.workbook_path.exists():
            raise ExcelHandlerError(f"Fichier introuvable : {self.workbook_path}")
        try:
            return load_workbook(self.workbook_path)
        except PermissionError as exc:
            raise ExcelHandlerError("Le fichier Excel semble verrouillé (PermissionError).") from exc

    def _get_sheet(self, workbook, sheet_name: str) -> Worksheet:
        if sheet_name not in workbook.sheetnames:
            raise ExcelHandlerError(f"Onglet introuvable : {sheet_name}")
        return workbook[sheet_name]

    @staticmethod
    def _last_data_row(sheet: Worksheet) -> int:
        for row in range(sheet.max_row, DATA_START_ROW - 1, -1):
            if any(sheet[f"{cfg.letter}{row}"].value not in (None, "") for cfg in COLUMN_DEFINITIONS):
                return row
        return DATA_START_ROW - 1

    def list_sheets(self) -> List[str]:
        wb = self._load_workbook()
        names = wb.sheetnames
        wb.close()
        return names

    def get_dynamic_options(self, sheet_name: str) -> Dict[str, List[str]]:
        wb = self._load_workbook()
        sheet = self._get_sheet(wb, sheet_name)
        options = {key: values[:] for key, values in DEFAULT_SELECT_OPTIONS.items()}

        last_row = self._last_data_row(sheet)
        for key in [cfg.key for cfg in COLUMN_DEFINITIONS if cfg.dropdown]:
            col = COLUMNS_BY_KEY[key].letter
            seen = set(options.get(key, []))
            dynamic = []
            for row in range(DATA_START_ROW, last_row + 1):
                val = sheet[f"{col}{row}"].value
                if val in (None, ""):
                    continue
                val_str = str(val)
                if val_str not in seen:
                    seen.add(val_str)
                    dynamic.append(val_str)
            options[key] = options.get(key, []) + sorted(dynamic)
        wb.close()
        return options

    @staticmethod
    def _copy_row_style_and_formulas(sheet: Worksheet, source_row: int, target_row: int) -> None:
        for cfg in COLUMN_DEFINITIONS:
            source = sheet[f"{cfg.letter}{source_row}"]
            target = sheet[f"{cfg.letter}{target_row}"]
            if source.has_style:
                target._style = copy(source._style)
            if source.number_format:
                target.number_format = source.number_format
            if source.protection:
                target.protection = copy(source.protection)
            if source.alignment:
                target.alignment = copy(source.alignment)
            if isinstance(source.value, str) and source.value.startswith("="):
                target.value = Translator(source.value, origin=f"{cfg.letter}{source_row}").translate_formula(
                    f"{cfg.letter}{target_row}"
                )

    @staticmethod
    def _apply_payload(sheet: Worksheet, row: int, payload: Dict[str, str]) -> None:
        for key in FIELD_KEYS:
            col = COLUMNS_BY_KEY[key].letter
            sheet[f"{col}{row}"] = payload.get(key, "")

    @staticmethod
    def _next_index(sheet: Worksheet, last_row: int) -> int:
        if last_row < DATA_START_ROW:
            return 1
        current = sheet[f"A{last_row}"].value
        try:
            return int(current) + 1
        except (TypeError, ValueError):
            return last_row - DATA_START_ROW + 2

    def add_requirement(self, sheet_name: str, payload: Dict[str, str]) -> int:
        wb = self._load_workbook()
        sheet = self._get_sheet(wb, sheet_name)
        last_row = self._last_data_row(sheet)
        new_row = max(last_row + 1, DATA_START_ROW)
        if last_row >= DATA_START_ROW:
            self._copy_row_style_and_formulas(sheet, last_row, new_row)

        sheet[f"A{new_row}"] = self._next_index(sheet, last_row)
        self._apply_payload(sheet, new_row, payload)

        wb.save(self.workbook_path)
        wb.close()
        return new_row

    def search_requirements(self, sheet_name: str, keyword: str) -> List[Dict[str, str]]:
        wb = self._load_workbook()
        sheet = self._get_sheet(wb, sheet_name)
        last_row = self._last_data_row(sheet)
        keyword_lower = keyword.lower().strip()
        results: List[Dict[str, str]] = []

        for row in range(DATA_START_ROW, last_row + 1):
            number = str(sheet[f"B{row}"].value or "")
            french_name = str(sheet[f"C{row}"].value or "")
            english_name = str(sheet[f"D{row}"].value or "")
            if keyword_lower in number.lower() or keyword_lower in french_name.lower() or keyword_lower in english_name.lower():
                results.append(
                    {
                        "row": row,
                        "number": number,
                        "french_name": french_name,
                        "english_name": english_name,
                    }
                )
        wb.close()
        return results

    def load_requirement(self, sheet_name: str, row: int) -> Dict[str, str]:
        wb = self._load_workbook()
        sheet = self._get_sheet(wb, sheet_name)
        data = {key: str(sheet[f"{COLUMNS_BY_KEY[key].letter}{row}"].value or "") for key in FIELD_KEYS}
        wb.close()
        return data

    def update_requirement(self, sheet_name: str, row: int, payload: Dict[str, str]) -> None:
        wb = self._load_workbook()
        sheet = self._get_sheet(wb, sheet_name)
        self._apply_payload(sheet, row, payload)
        wb.save(self.workbook_path)
        wb.close()

    def archive_requirement(self, sheet_name: str, row: int) -> None:
        wb = self._load_workbook()
        sheet = self._get_sheet(wb, sheet_name)
        current = str(sheet[f"C{row}"].value or "")
        if not current.startswith(ARCHIVE_MARKER):
            sheet[f"C{row}"] = f"{ARCHIVE_MARKER} {current}".strip()
        wb.save(self.workbook_path)
        wb.close()
