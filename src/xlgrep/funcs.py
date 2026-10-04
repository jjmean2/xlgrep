"""--list-funcs: count the functions used in formulas."""

from __future__ import annotations

import re
from collections.abc import Iterable
from dataclasses import dataclass, field

from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import TokenizerError
from openpyxl.utils.formulas import FORMULAE

from .text import mask_string_literals

# Excel writes functions added after Excel 2007 with "_xlfn." (and "_xlws." for some
# worksheet-only ones); XLL add-in functions get "_xll.". FORMULAE is the 2007 set,
# so together they tell built-ins from LAMBDA names and VBA/add-in functions.
_FUTURE_PREFIXES = ("_xlfn.", "_xlws.")
_ADDIN_PREFIX = "_xll."
_EXTRA_BUILTINS = {"TABLE"}  # what-if data tables, {=TABLE(r,c)}
_FALLBACK_RE = re.compile(r"(?<![\w.])([A-Za-z_\\][\w.]*)\s*\(")

BUILTIN, LAMBDA, CUSTOM = "builtin", "lambda", "custom"

# Row numbers of cell references (A1 -> A, $B$2 -> $B$). Formulas filled down a
# column then share one "shape" and one tokenisation. Digits before "(" or inside
# names (LOG10(, DEC2BIN() are kept, so function names never collide.
_SHAPE_RE = re.compile(r"(?<=[A-Za-z$])\d+(?![\w(])")


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
        self._calls: dict[str, list[str]] = {}  # formula shape -> function_calls()

    def add(self, group: object, file: str, formula: str, defined_names: set[str]) -> None:
        seen: set[str] = set()
        stats = self.groups.setdefault(group, {})
        shape = _SHAPE_RE.sub("", formula)
        calls = self._calls.get(shape)
        if calls is None:
            calls = self._calls[shape] = function_calls(formula)
        for raw in calls:
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
