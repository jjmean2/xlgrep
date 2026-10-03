"""Reading workbooks into per-sheet cell grids."""

from __future__ import annotations

import warnings
from collections.abc import Callable, Iterator
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl
from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula

from .text import normalize_formula, value_text


@dataclass
class Cell:
    row: int
    col: int
    formula: str | None = None  # formula text including the leading "=", if any
    value: object = None  # literal value, or the cached result of a formula
    has_cached_value: bool = False

    @property
    def is_formula(self) -> bool:
        return self.formula is not None

    @property
    def value_str(self) -> str:
        return value_text(self.value)

    @property
    def display(self) -> str:
        """What the cell looks like in the formula bar."""
        return self.formula if self.formula is not None else self.value_str


@dataclass
class Sheet:
    name: str
    hidden: bool
    cells: dict[tuple[int, int], Cell] = field(default_factory=dict)

    def get(self, row: int, col: int) -> Cell | None:
        return self.cells.get((row, col))

    def sorted_cells(self) -> list[Cell]:
        return [self.cells[k] for k in sorted(self.cells)]


def _formula_text(raw: object, raw_formula: bool) -> str | None:
    if isinstance(raw, ArrayFormula):
        text = raw.text or ""
    elif isinstance(raw, DataTableFormula):
        args = ",".join(str(a) for a in (raw.r1, raw.r2) if a)
        text = f"=TABLE({args})"
    elif isinstance(raw, str) and raw.startswith("="):
        text = raw
    else:
        return None
    return text if raw_formula else normalize_formula(text)


def read_sheets(
    path: Path,
    *,
    with_values: bool,
    raw_formula: bool = False,
    sheet_filter: Callable[[str, bool], bool] | None = None,
) -> Iterator[Sheet]:
    """Yield the worksheets of ``path`` that pass ``sheet_filter(name, hidden)``.

    Formula cells carry the formula text; with ``with_values`` they also carry the
    value Excel cached the last time the file was saved (requires a second pass
    over the file, since openpyxl reads either formulas or values, not both).
    """
    with warnings.catch_warnings():
        # openpyxl warns about unsupported extensions (data validation, etc.); irrelevant here.
        warnings.simplefilter("ignore")
        wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
        values_wb = openpyxl.load_workbook(path, read_only=True, data_only=True) if with_values else None
    try:
        for ws in wb.worksheets:
            hidden = ws.sheet_state != "visible"
            if sheet_filter and not sheet_filter(ws.title, hidden):
                continue
            sheet = Sheet(ws.title, hidden)
            for row in ws.iter_rows():
                for c in row:
                    if c.value is None or not hasattr(c, "row"):
                        continue
                    formula = _formula_text(c.value, raw_formula)
                    if formula is not None:
                        cell = Cell(c.row, c.column, formula=formula)
                    else:
                        cell = Cell(c.row, c.column, value=c.value, has_cached_value=True)
                    sheet.cells[(c.row, c.column)] = cell
            if values_wb is not None:
                _fill_cached_values(sheet, values_wb[ws.title])
            yield sheet
    finally:
        wb.close()
        if values_wb is not None:
            values_wb.close()


def _fill_cached_values(sheet: Sheet, ws) -> None:
    for row in ws.iter_rows():
        for c in row:
            if c.value is None or not hasattr(c, "row"):
                continue
            cell = sheet.cells.get((c.row, c.column))
            if cell is not None and cell.is_formula:
                cell.value = c.value
                cell.has_cached_value = True
