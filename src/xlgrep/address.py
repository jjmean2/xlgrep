"""Cell address helpers: column letters, A1 ranges and sheet-name quoting."""

from __future__ import annotations

import re
from dataclasses import dataclass

_CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})?\$?(\d+)?$")
_PLAIN_SHEET_RE = re.compile(r"^[^\W\d]\w*(?:\.\w+)*$")
_LOOKS_LIKE_REF_RE = re.compile(r"^(?:[A-Za-z]{1,3}\d+|R\d*C\d*)$", re.IGNORECASE)


def col_letter(col: int) -> str:
    """1 -> 'A', 27 -> 'AA'."""
    letters = ""
    while col > 0:
        col, rem = divmod(col - 1, 26)
        letters = chr(ord("A") + rem) + letters
    return letters


def col_index(letters: str) -> int:
    """'A' -> 1, 'AA' -> 27."""
    index = 0
    for ch in letters.upper():
        index = index * 26 + (ord(ch) - ord("A") + 1)
    return index


def cell_name(row: int, col: int) -> str:
    return f"{col_letter(col)}{row}"


def quote_sheet(name: str) -> str:
    """Quote a sheet name the way Excel does in references ('Raw Data'!A1)."""
    if _PLAIN_SHEET_RE.match(name) and not _LOOKS_LIKE_REF_RE.match(name):
        return name
    return "'" + name.replace("'", "''") + "'"


def qualified(sheet: str, row: int, col: int) -> str:
    return f"{quote_sheet(sheet)}!{cell_name(row, col)}"


@dataclass(frozen=True)
class CellRange:
    """Inclusive cell range. ``None`` bounds are open (e.g. whole columns 'A:C')."""

    min_row: int | None = None
    min_col: int | None = None
    max_row: int | None = None
    max_col: int | None = None

    def contains(self, row: int, col: int) -> bool:
        return (
            (self.min_row is None or row >= self.min_row)
            and (self.max_row is None or row <= self.max_row)
            and (self.min_col is None or col >= self.min_col)
            and (self.max_col is None or col <= self.max_col)
        )

    def intersects(self, other: CellRange) -> bool:
        def overlap(lo1, hi1, lo2, hi2) -> bool:
            return (hi1 is None or lo2 is None or lo2 <= hi1) and (hi2 is None or lo1 is None or lo1 <= hi2)

        return overlap(self.min_row, self.max_row, other.min_row, other.max_row) and overlap(
            self.min_col, self.max_col, other.min_col, other.max_col
        )


def parse_range(text: str) -> CellRange:
    """Parse 'A1:F100', 'B2', 'A:C' or '3:10' into a CellRange."""
    parts = text.strip().split(":")
    if len(parts) == 1:
        parts = parts * 2
    if len(parts) != 2:
        raise ValueError(f"invalid range: {text!r}")
    bounds = []
    for part in parts:
        m = _CELL_RE.match(part.strip())
        if not m or not (m.group(1) or m.group(2)):
            raise ValueError(f"invalid range: {text!r}")
        col = col_index(m.group(1)) if m.group(1) else None
        row = int(m.group(2)) if m.group(2) else None
        bounds.append((row, col))
    (r1, c1), (r2, c2) = bounds
    if r1 is not None and r2 is not None and r1 > r2:
        r1, r2 = r2, r1
    if c1 is not None and c2 is not None and c1 > c2:
        c1, c2 = c2, c1
    return CellRange(r1, c1, r2, c2)
