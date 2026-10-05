"""Functions used in formulas: extraction, classification, and --list-funcs."""

from __future__ import annotations

import csv
import json
import re
import sys
from collections.abc import Iterable
from dataclasses import dataclass, field
from typing import TextIO

from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import TokenizerError
from openpyxl.utils.formulas import FORMULAE

from .output import Style
from .text import display_width, mask_string_literals

# Excel writes functions added after Excel 2007 with "_xlfn." (and "_xlws." for some
# worksheet-only ones); XLL add-in functions get "_xll.". FORMULAE is the 2007 set,
# so together they tell built-ins from LAMBDA names and VBA/add-in functions.
_FUTURE_PREFIXES = ("_xlfn.", "_xlws.")
_ADDIN_PREFIX = "_xll."
_EXTRA_BUILTINS = {"TABLE"}  # what-if data tables, {=TABLE(r,c)}
_FALLBACK_RE = re.compile(r"(?<![\w.])([A-Za-z_\\][\w.]*)\s*\(")

BUILTIN, LAMBDA, CUSTOM = "builtin", "lambda", "custom"

# Recalculated on every change anywhere in the workbook, not just when their inputs change.
VOLATILE = frozenset({"NOW", "TODAY", "RAND", "RANDBETWEEN", "RANDARRAY", "OFFSET", "INDIRECT", "CELL", "INFO"})

# Row numbers of cell references (A1 -> A, $B$2 -> $B$). Digits before "(" or inside
# names (LOG10(, DEC2BIN() are kept, so function names never collide.
_SHAPE_RE = re.compile(r"(?<=[A-Za-z$])\d+(?![\w(])")


def formula_shape(formula: str) -> str:
    """The formula without the row numbers of its references.

    Formulas filled down a column share one shape, and so do the functions they call
    and the sheets and workbooks they refer to; caching by shape avoids redoing that
    work for every row.
    """
    return _SHAPE_RE.sub("", formula)


def function_calls(formula: str) -> list[str]:
    """Function names called in ``formula`` (raw, prefixes kept), in order."""
    try:
        tokens = Tokenizer(formula).items
    except TokenizerError:
        return [m.group(1) for m in _FALLBACK_RE.finditer(mask_string_literals(formula))]
    return [t.value[:-1] for t in tokens if t.type == "FUNC" and t.subtype == "OPEN"]


def classify(raw_name: str, defined_names: set[str]) -> tuple[str, str]:
    """(display name, category) for a raw function name from a formula."""
    name, future, addin = raw_name, False, False
    while True:
        lower = name.lower()
        if lower.startswith(_FUTURE_PREFIXES):
            name, future = name[6:], True
        elif lower.startswith(_ADDIN_PREFIX):
            name, addin = name[5:], True
        else:
            break
    upper = name.upper()
    if addin:
        return name, CUSTOM
    if future or upper in FORMULAE or upper in _EXTRA_BUILTINS:
        return upper, BUILTIN
    if upper in defined_names:
        return name, LAMBDA
    return name, CUSTOM


class CallScanner:
    """function_calls() with a cache per formula shape (see _SHAPE_RE)."""

    def __init__(self):
        self._calls: dict[str, list[str]] = {}

    def __call__(self, formula: str) -> list[str]:
        shape = formula_shape(formula)
        calls = self._calls.get(shape)
        if calls is None:
            calls = self._calls[shape] = function_calls(formula)
        return calls


@dataclass
class FuncStat:
    name: str
    category: str
    calls: int = 0
    places: int = 0
    files: set[str] = field(default_factory=set)


class FuncCounter:
    """Aggregates function usage per group (None, a file, or a (file, sheet) pair)."""

    def __init__(self, name_filter):
        self.name_filter = name_filter  # callable(display_name) -> bool
        self.groups: dict[object, dict[str, FuncStat]] = {}
        self._calls = CallScanner()

    def add(self, group: object, file: str, formula: str, defined_names: set[str]) -> None:
        seen: set[str] = set()
        stats = self.groups.setdefault(group, {})
        for raw in self._calls(formula):
            name, category = classify(raw, defined_names)
            if not self.name_filter(name):
                continue
            key = name.upper()
            stat = stats.get(key)
            if stat is None:
                stat = stats[key] = FuncStat(name, category)
            stat.calls += 1
            if key not in seen:
                stat.places += 1
                seen.add(key)
            stat.files.add(file)

    def merge(self, groups: dict[object, dict[str, FuncStat]]) -> None:
        """Fold in another counter's groups (e.g. from a worker process)."""
        for group, stats in groups.items():
            mine = self.groups.setdefault(group, {})
            for key, stat in stats.items():
                if key not in mine:
                    mine[key] = FuncStat(stat.name, stat.category)
                mine[key].calls += stat.calls
                mine[key].places += stat.places
                mine[key].files |= stat.files

    def rows(self, sort: str) -> Iterable[tuple[object, list[FuncStat]]]:
        for group, stats in self.groups.items():
            if not stats:
                continue
            if sort == "name":
                ordered = sorted(stats.values(), key=lambda s: s.name.upper())
            else:
                ordered = sorted(stats.values(), key=lambda s: (-s.calls, -s.places, s.name.upper()))
            yield group, ordered


def write_func_stats(groups, by: str | None, mode: str, style: Style, out: TextIO | None = None) -> bool:
    """Render --list-funcs results. ``groups`` yields (group key, stats). Returns True if any row."""
    out = out if out is not None else sys.stdout
    found = False
    writer = csv.writer(out, lineterminator="\n") if mode == "csv" else None
    if writer is not None:
        keys = {"file": ["file"], "sheet": ["file", "sheet"]}.get(by, [])
        writer.writerow(keys + ["function", "category", "calls", "places"] + ([] if by else ["files"]))

    current_file = None
    for group, stats in groups:
        found = True
        file, sheet = (group if by == "sheet" else (group, None)) if by else (None, None)
        if mode == "json":
            for s in stats:
                record = {"file": file, "sheet": sheet} if by == "sheet" else {"file": file} if by else {}
                record.update(function=s.name, category=s.category, calls=s.calls, places=s.places)
                if not by:
                    record["files"] = len(s.files)
                out.write(json.dumps(record, ensure_ascii=False) + "\n")
        elif writer is not None:
            prefix = [file, sheet or ""] if by == "sheet" else [file] if by else []
            for s in stats:
                writer.writerow(prefix + [s.name, s.category, s.calls, s.places] + ([] if by else [len(s.files)]))
        else:
            indent = ""
            if by:
                if file != current_file:
                    if current_file is not None:
                        out.write("\n")
                    out.write(style.heading(file) + "\n")
                    current_file = file
                indent = "  "
                if by == "sheet":
                    out.write("  " + style.sheet(sheet if sheet is not None else "(workbook)") + "\n")
                    indent = "    "
            _write_func_table(stats, include_files=not by, indent=indent, style=style, out=out)
    return found


def _write_func_table(stats: list[FuncStat], include_files: bool, indent: str, style: Style, out: TextIO) -> None:
    headers = ["FUNCTION", "CALLS", "PLACES"] + (["FILES"] if include_files else [])
    rows = [[s.name, str(s.calls), str(s.places)] + ([str(len(s.files))] if include_files else []) for s in stats]
    widths = [max(display_width(r[i]) for r in [*rows, headers]) for i in range(len(headers))]

    def fmt(cells: list[str]) -> str:
        first = cells[0] + " " * (widths[0] - display_width(cells[0]))
        return "  ".join([first] + [c.rjust(w) for c, w in zip(cells[1:], widths[1:], strict=True)])

    out.write(indent + style.dim(fmt(headers)) + "\n")
    for s, row in zip(stats, rows, strict=True):
        tag = "" if s.category == BUILTIN else "  " + style.header(s.category)
        out.write(indent + fmt(row) + tag + "\n")
