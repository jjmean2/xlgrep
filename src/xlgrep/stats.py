"""--stats: how big each workbook is, and what would be hard to move elsewhere.

Collection (stats_file, run per file, possibly in a worker process), combining
(file and run totals) and rendering (table, cards, JSON, CSV) are in that order.

Formula metrics (formulas, unique, longest, volatile, custom functions, arrays)
are about cell formulas; formulas in conditional formats, validations and names
are counted under those objects.
"""

from __future__ import annotations

import csv
import json
import re
import sys
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import TextIO

from .funcs import BUILTIN, LAMBDA, VOLATILE, CallScanner, classify
from .objects import count_rules
from .output import Style
from .package import sheet_parts
from .refs import relative_form
from .scope import READ_ERRORS, Scope, ScopedWorkbook
from .text import display_width

ERROR_VALUES = frozenset(
    {
        "#NULL!",
        "#DIV/0!",
        "#VALUE!",
        "#REF!",
        "#NAME?",
        "#NUM!",
        "#N/A",
        "#GETTING_DATA",
        "#SPILL!",
        "#CALC!",
        "#FIELD!",
        "#BLOCKED!",
        "#CONNECT!",
        "#BUSY!",
        "#UNKNOWN!",
    }
)
_EXTERNAL_LINK_RE = re.compile(r"(^|/)externalLinks/externalLink\d+\.xml$")


@dataclass
class Stats:
    """Metrics for one sheet, or summed over sheets/files (see combine)."""

    name: str
    hidden: bool = False
    dimension: str | None = None  # used range Excel recorded
    values: int = 0  # non-empty cells without a formula
    formulas: int = 0
    unique: set[str] = field(default_factory=set)  # formulas in relative (R1C1) form
    longest: int = 0  # characters in the longest formula
    volatile_formulas: int = 0  # formulas calling at least one volatile function
    volatile: Counter = field(default_factory=Counter)  # function -> formulas calling it
    arrays: int = 0
    data_tables: int = 0
    errors: Counter = field(default_factory=Counter)  # "#REF!" -> cells showing it
    custom: set[str] = field(default_factory=set)  # VBA / add-in functions
    lambdas: set[str] = field(default_factory=set)  # LAMBDA defined names called
    cf: int = 0  # conditional formatting rules
    dv: int = 0  # data validations
    names: int = 0  # defined names
    notes: int = 0
    tables: int = 0
    pivots: int = 0
    charts: int = 0

    @property
    def cells(self) -> int:
        return self.values + self.formulas


@dataclass
class FileStats:
    path: str
    size: int
    sheets: list[Stats] = field(default_factory=list)
    hidden_sheets: int = 0
    workbook_names: int = 0  # names not scoped to a sheet
    vba: bool = False
    external_links: int = 0
    connections: bool = False  # data connections (databases, web queries, Power Query)
    errors: list[str] = field(default_factory=list)  # read errors

    @property
    def total(self) -> Stats:
        total = combine(self.path, self.sheets)
        total.names += self.workbook_names
        return total


@dataclass
class StatsConfig:
    scope: Scope
    by: str | None  # None/"file": a row per file; "sheet": a row per sheet


# ---------------------------------------------------------------- collecting


def stats_file(path: Path, cfg: StatsConfig) -> FileStats:
    result = FileStats(str(path), path.stat().st_size if path.exists() else 0)
    calls = CallScanner()
    try:
        # Raw formulas: the _xlfn./_xll. prefixes tell built-ins from custom functions.
        with ScopedWorkbook(path, cfg.scope, raw_formula=True) as book:
            names = {name for _, name in book.defined_names}
            for part in book.parts():
                if part.info is None:  # workbook-scoped names
                    result.workbook_names += len(part.objects)
                    continue
                stats = Stats(part.sheet_name, part.hidden, part.sheet.dimension if part.sheet else None)
                _count_cells(stats, part.cells, calls, names)
                _count_objects(stats, part.objects)
                stats.cf, stats.dv = count_rules(book.zip, part.info)
                stats.tables, stats.pivots, stats.charts = sheet_parts(book.zip, part.info)
                result.sheets.append(stats)
                result.hidden_sheets += part.hidden
            entries = book.zip.namelist()
            result.vba = any(e.endswith("vbaProject.bin") for e in entries)
            result.external_links = sum(1 for e in entries if _EXTERNAL_LINK_RE.search(e))
            result.connections = any(e.endswith("connections.xml") for e in entries)
    except READ_ERRORS as exc:
        result.errors.append(f"{path}: cannot read workbook ({exc})")
    return result


def _count_cells(stats: Stats, cells, calls: CallScanner, names: set[str]) -> None:
    group_forms: dict[str, str] = {}  # shared-formula group -> its relative form
    for cell in cells:
        if isinstance(cell.value, str) and cell.value in ERROR_VALUES:
            stats.errors[cell.value] += 1
        formula = cell.formula
        if formula is None:
            stats.values += 1
            continue
        stats.formulas += 1
        stats.longest = max(stats.longest, len(formula))
        # Cells of a shared-formula group have the same logic: compute the form once.
        form = group_forms.get(cell.shared) if cell.shared else None
        if form is None:
            form = relative_form(formula, cell.row, cell.col)
            if cell.shared:
                group_forms[cell.shared] = form
        stats.unique.add(form)
        if cell.array:
            stats.arrays += 1
        if formula.upper().startswith("=TABLE("):
            stats.data_tables += 1
        volatile = set()
        for raw in calls(formula):
            name, category = classify(raw, names)
            if name in VOLATILE:
                volatile.add(name)
            elif category == LAMBDA:
                stats.lambdas.add(name)
            elif category != BUILTIN:
                stats.custom.add(name)
        if volatile:
            stats.volatile_formulas += 1
            stats.volatile.update(volatile)


def _count_objects(stats: Stats, objects) -> None:
    for obj in objects:
        if obj.object == "name":
            stats.names += 1
        elif obj.object == "note":
            stats.notes += 1


# ---------------------------------------------------------------- combining


def combine(name: str, parts: list[Stats]) -> Stats:
    """Sum sheet (or file) stats. Unique formulas are a union: the same logic on
    two sheets is one formula to reimplement."""
    total = Stats(name)
    for s in parts:
        total.values += s.values
        total.formulas += s.formulas
        total.unique |= s.unique
        total.longest = max(total.longest, s.longest)
        total.volatile_formulas += s.volatile_formulas
        total.volatile.update(s.volatile)
        total.arrays += s.arrays
        total.data_tables += s.data_tables
        total.errors.update(s.errors)
        total.custom |= s.custom
        total.lambdas |= s.lambdas
        for attr in ("cf", "dv", "names", "notes", "tables", "pivots", "charts"):
            setattr(total, attr, getattr(total, attr) + getattr(s, attr))
    return total


# ---------------------------------------------------------------- rendering


def write_stats(files: list[FileStats], by: str | None, mode: str, style: Style, out: TextIO | None = None) -> None:
    out = out if out is not None else sys.stdout
    if mode == "json":
        _write_json(files, by, out)
    elif mode == "csv":
        _write_csv(files, by, out)
    elif mode == "cards":
        _write_cards(files, style, out)
    elif by == "sheet":
        _write_sheet_tables(files, style, out)
    else:
        _write_file_table(files, style, out)


def _size(n: int) -> str:
    for unit in ("B", "KB", "MB"):
        if n < 1024:
            return f"{n} {unit}" if unit == "B" else f"{n:.1f} {unit}"
        n /= 1024
    return f"{n:.1f} GB"


def _num(n: int) -> str:
    return f"{n:,}"


def _table(rows: list[list[str]], style: Style, out: TextIO, indent: str = "", bold_last: bool = False) -> None:
    """Left-align the first column, right-align the rest; first row is the header."""
    widths = [max(display_width(r[i]) for r in rows) for i in range(len(rows[0]))]

    def fmt(cells: list[str]) -> str:
        first = cells[0] + " " * (widths[0] - display_width(cells[0]))
        rest = [c.rjust(w) for c, w in zip(cells[1:], widths[1:], strict=True)]
        return "  ".join([first, *rest]).rstrip()

    out.write(indent + style.dim(fmt(rows[0])) + "\n")
    for i, row in enumerate(rows[1:], 1):
        line = fmt(row)
        out.write(indent + (style.heading(line) if bold_last and i == len(rows) - 1 else line) + "\n")


def _file_row(f: FileStats, t: Stats) -> list[str]:
    sheets = f"{len(f.sheets)}/{f.hidden_sheets}" if f.hidden_sheets else str(len(f.sheets))
    return [
        f.path,
        _size(f.size),
        sheets,
        _num(t.cells),
        _num(t.formulas),
        _num(len(t.unique)),
        _num(t.volatile_formulas),
        _num(t.arrays),
        _num(sum(t.errors.values())),
        "yes" if f.vba else "no",
        str(f.external_links),
    ]


def _write_file_table(files: list[FileStats], style: Style, out: TextIO) -> None:
    rows = [["FILE", "SIZE", "SHEETS", "CELLS", "FORMULAS", "UNIQUE", "VOLATILE", "ARRAY", "ERRORS", "VBA", "EXT"]]
    rows += [_file_row(f, f.total) for f in files]
    if len(files) > 1:
        t = combine("TOTAL", [f.total for f in files])
        rows.append(
            [
                "TOTAL",
                _size(sum(f.size for f in files)),
                str(sum(len(f.sheets) for f in files)),
                _num(t.cells),
                _num(t.formulas),
                _num(len(t.unique)),
                _num(t.volatile_formulas),
                _num(t.arrays),
                _num(sum(t.errors.values())),
                str(sum(f.vba for f in files)),
                str(sum(f.external_links for f in files)),
            ]
        )
    _table(rows, style, out, bold_last=len(files) > 1)


def _write_sheet_tables(files: list[FileStats], style: Style, out: TextIO) -> None:
    for i, f in enumerate(files):
        if i:
            out.write("\n")
        out.write(style.heading(f.path) + "\n")
        rows = [["SHEET", "RANGE", "CELLS", "FORMULAS", "UNIQUE", "VOLATILE", "ARRAY", "ERRORS", "CF", "DV"]]
        for s in f.sheets:
            rows.append(
                [
                    s.name + (" (hidden)" if s.hidden else ""),
                    s.dimension or "",
                    _num(s.cells),
                    _num(s.formulas),
                    _num(len(s.unique)),
                    _num(s.volatile_formulas),
                    _num(s.arrays),
                    _num(sum(s.errors.values())),
                    _num(s.cf),
                    _num(s.dv),
                ]
            )
        _table(rows, style, out, indent="  ")


def _counts(counter: Counter) -> str:
    return ", ".join(f"{k} {v}" for k, v in counter.most_common())


def _write_cards(files: list[FileStats], style: Style, out: TextIO) -> None:
    for i, f in enumerate(files):
        if i:
            out.write("\n")
        t = f.total
        sheet_names = ", ".join(s.name + (" (hidden)" if s.hidden else "") for s in f.sheets)
        lines = [
            (
                "sheets",
                f"{len(f.sheets)}"
                + (f" ({f.hidden_sheets} hidden)" if f.hidden_sheets else "")
                + (f": {sheet_names}" if sheet_names else ""),
            ),
            (
                "cells",
                (
                    f"{_num(t.cells)} (formulas {_num(t.formulas)}, unique {_num(len(t.unique))}, "
                    f"longest {_num(t.longest)} chars)"
                ),
            ),
            ("volatile", f"{_num(t.volatile_formulas)} formulas" + (f": {_counts(t.volatile)}" if t.volatile else "")),
            ("arrays", f"{_num(t.arrays)} array formulas, {_num(t.data_tables)} data tables"),
            ("errors", f"{_num(sum(t.errors.values()))} cells" + (f": {_counts(t.errors)}" if t.errors else "")),
            (
                "code",
                f"VBA {'yes' if f.vba else 'no'}"
                + (f"; custom functions {', '.join(sorted(t.custom))}" if t.custom else "")
                + (f"; LAMBDA {', '.join(sorted(t.lambdas))}" if t.lambdas else ""),
            ),
            (
                "rules",
                (
                    f"{_num(t.cf)} conditional formats, {_num(t.dv)} validations, "
                    f"{_num(t.names)} names, {_num(t.notes)} notes"
                ),
            ),
            ("objects", f"{_num(t.tables)} tables, {_num(t.pivots)} pivot tables, {_num(t.charts)} charts"),
            ("links", f"{f.external_links} external workbooks" + ("; data connections" if f.connections else "")),
        ]
        out.write(f"{style.heading(f.path)}  {style.dim(_size(f.size))}\n")
        out.writelines(f"  {style.header(label.ljust(9))}  {text}\n" for label, text in lines)


def _record(s: Stats) -> dict:
    return {
        "range": s.dimension,
        "cells": s.cells,
        "values": s.values,
        "formulas": s.formulas,
        "unique_formulas": len(s.unique),
        "longest_formula": s.longest,
        "volatile_formulas": s.volatile_formulas,
        "volatile": dict(s.volatile.most_common()),
        "array_formulas": s.arrays,
        "data_tables": s.data_tables,
        "error_cells": sum(s.errors.values()),
        "errors": dict(s.errors.most_common()),
        "custom_functions": sorted(s.custom),
        "lambdas": sorted(s.lambdas),
        "conditional_formats": s.cf,
        "validations": s.dv,
        "names": s.names,
        "notes": s.notes,
        "tables": s.tables,
        "pivot_tables": s.pivots,
        "charts": s.charts,
    }


def _records(files: list[FileStats], by: str | None) -> list[dict]:
    """One record per file, or per sheet with --by sheet; shared by JSON and CSV."""
    records = []
    for f in files:
        if by == "sheet":
            records += [{"file": f.path, "sheet": s.name, "hidden": s.hidden, **_record(s)} for s in f.sheets]
        else:
            record = {
                "file": f.path,
                "size": f.size,
                "sheets": len(f.sheets),
                "hidden_sheets": f.hidden_sheets,
                **_record(f.total),
                "vba": f.vba,
                "external_links": f.external_links,
                "data_connections": f.connections,
            }
            del record["range"]  # a per-sheet notion
            records.append(record)
    return records


def _write_json(files: list[FileStats], by: str | None, out: TextIO) -> None:
    out.writelines(json.dumps(record, ensure_ascii=False) + "\n" for record in _records(files, by))


def _flat(value: object) -> object:
    """A CSV cell for a record value: lists as "a; b", counts as "#REF! 5; #N/A 2"."""
    if isinstance(value, list):
        return "; ".join(value)
    if isinstance(value, dict):
        return "; ".join(f"{k} {v}" for k, v in value.items())
    return value


def _write_csv(files: list[FileStats], by: str | None, out: TextIO) -> None:
    records = [{k: _flat(v) for k, v in r.items()} for r in _records(files, by)]
    writer = csv.DictWriter(out, fieldnames=list(records[0]) if records else ["file"], lineterminator="\n")
    writer.writeheader()
    writer.writerows(records)
