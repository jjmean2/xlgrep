"""Per-file work, and running it across files (in parallel when worthwhile).

A worker gets a path and a picklable config, and returns the file's result with
its output already rendered, so the parent only prints results in file order.
"""

from __future__ import annotations

import io
import os
import re
import zipfile
from collections.abc import Callable, Iterator
from concurrent.futures import Future, ProcessPoolExecutor
from dataclasses import dataclass, field
from pathlib import Path
from typing import Generic, TypeVar
from xml.etree import ElementTree as ET

from .address import CellRange
from .funcs import FuncCounter, FuncStat
from .matcher import Matcher, merge_spans
from .objects import SheetObject
from .output import FORMATTERS, Match, ObjectMatch, OutputOptions
from .refs import RefFinder, RefHit, Target
from .scope import Scope, ScopedWorkbook
from .workbook import Cell

# What a broken or unreadable workbook raises; reported per file, not fatal.
READ_ERRORS = (zipfile.BadZipFile, KeyError, OSError, ValueError, ET.ParseError)

# Below this much input, starting worker processes costs more than it saves.
PARALLEL_MIN_BYTES = 1_000_000


# ---------------------------------------------------------------- searching


@dataclass
class SearchConfig:
    scope: Scope
    matcher: Matcher
    targets: list[Target]  # --ref ranges
    search_in: str  # "auto" | "formula" | "value"
    invert: bool
    output: OutputOptions
    mode: str  # "line" | "pretty" | "json" | "csv"
    summary_only: bool  # -l / -c / -q: only the count matters
    limit: int | None  # max matching cells/objects per file

    @property
    def formulas_only(self) -> bool:
        """Only formula cells can match: -f alone, or --ref (and no -v)."""
        if self.invert:
            return False
        return bool(self.targets) or (self.matcher.pattern is None and self.matcher.func_pattern is not None)

    @property
    def shows_neighbours(self) -> bool:
        out = self.output
        return self.mode == "pretty" or bool(out.before or out.after or out.row_context or out.header_row)


@dataclass
class FileResult:
    count: int = 0
    output: str = ""
    grouped: bool = False  # printed context groups (line mode); the parent adds "--" between files
    errors: list[str] = field(default_factory=list)


class Searcher:
    """Decides which cells and objects match, and what to highlight."""

    def __init__(self, cfg: SearchConfig):
        self.matcher = cfg.matcher
        self.search_in = cfg.search_in
        self.invert = cfg.invert
        self.targets = cfg.targets
        self.ref_finder: RefFinder | None = None  # per workbook, when --ref is used

    def start_workbook(self, book: ScopedWorkbook) -> None:
        if self.targets:
            self.ref_finder = RefFinder(self.targets, book.sheet_names, book.defined_names)

    def search_cells(self, cells: list[Cell], sheet: str, limit: int | None) -> list[Match]:
        matches: list[Match] = []
        for cell in cells:
            if limit is not None and len(matches) >= limit:
                break
            searched = self._searched_text(cell)
            if searched is None or searched[0] == "":
                continue
            text, kind = searched
            selected = self._select(text, kind, sheet)
            if selected is not None:
                matches.append(Match(cell, text, kind, *selected))
        return matches

    def search_objects(self, objects: list[SheetObject], limit: int | None) -> list[ObjectMatch]:
        matches: list[ObjectMatch] = []
        for obj in objects:
            if limit is not None and len(matches) >= limit:
                break
            kind = "formula" if obj.is_formula else "value"
            if self.search_in not in ("auto", kind):
                continue
            selected = self._select(obj.text, kind, obj.sheet, obj.areas if obj.object in ("cf", "dv") else None)
            if selected is not None:
                matches.append(ObjectMatch(obj, *selected))
        return matches

    def _searched_text(self, cell: Cell) -> tuple[str, str] | None:
        """The (text, kind) to match a cell on, or None if --in rules it out."""
        if self.search_in == "formula":
            return (cell.formula, "formula") if cell.is_formula else None
        if self.search_in == "value":
            return (cell.value_str, "value") if cell.has_cached_value else None
        return (cell.formula, "formula") if cell.is_formula else (cell.value_str, "value")

    def _select(self, text: str, kind: str, sheet: str | None,
                applies_to: list[CellRange] | None = None) -> tuple[list, list[RefHit] | None] | None:
        """(spans to highlight, --ref hits) if the text is selected, else None. Honours -v.

        Patterns and --ref combine as AND: the text must match the patterns (if any)
        and reference a target (if any).
        """
        if self.targets and kind != "formula":
            return None  # references only exist in formulas
        spans = self.matcher.spans(text, is_formula=kind == "formula")
        selected = bool(spans) or self.matcher.empty
        hits = None
        if selected and self.targets:
            hits = self.ref_finder.find(text, sheet, applies_to) if self.ref_finder else []
            selected = bool(hits)
        if selected == self.invert:
            return None
        if self.invert:
            return [], None
        if hits:
            spans = merge_spans(spans + [(h.start, h.end) for h in hits])
        return spans, hits


def search_file(path: Path, cfg: SearchConfig) -> FileResult:
    """Search one workbook; the result carries its rendered output."""
    buf = io.StringIO()
    formatter = FORMATTERS[cfg.mode](cfg.output, path, buf)
    searcher = Searcher(cfg)
    result = FileResult()
    # Cells that can't match are only read if they're shown around a match.
    formulas_only = cfg.formulas_only and not cfg.shows_neighbours
    try:
        with ScopedWorkbook(path, cfg.scope, raw_formula=cfg.scope.raw_formula,
                            formulas_only=formulas_only) as book:
            searcher.start_workbook(book)
            for part in book.parts():
                if _left(cfg.limit, result.count) == 0:
                    break
                cell_matches = searcher.search_cells(part.cells, part.sheet_name, _left(cfg.limit, result.count))
                result.count += len(cell_matches)
                object_matches = searcher.search_objects(part.objects, _left(cfg.limit, result.count))
                result.count += len(object_matches)
                if cfg.summary_only:
                    continue
                if cell_matches:
                    formatter.write_sheet(part.sheet, cell_matches)
                if object_matches:
                    formatter.write_objects(part.sheet_name, part.hidden, object_matches)
    except READ_ERRORS as exc:
        result.errors.append(f"{path}: cannot read workbook ({exc})")
    if result.count and not cfg.summary_only:
        formatter.end(result.count)
    result.output = buf.getvalue()
    result.grouped = formatter.grouped
    return result


def _left(limit: int | None, count: int) -> int | None:
    """How many more matches the per-file limit allows (None: no limit)."""
    return None if limit is None else max(0, limit - count)


# ---------------------------------------------------------------- --list-funcs


@dataclass
class FuncConfig:
    scope: Scope
    pattern: re.Pattern[str] | None  # -e: regex on function names
    wanted: set[str]  # -f: function names, upper-cased
    by: str | None  # None | "file" | "sheet"

    def name_ok(self, name: str) -> bool:
        if self.pattern is None and not self.wanted:
            return True
        return bool(self.pattern is not None and self.pattern.search(name)) or name.upper() in self.wanted


def count_file(path: Path, cfg: FuncConfig) -> tuple[dict[object, dict[str, FuncStat]], list[str]]:
    """--list-funcs for one workbook: (group -> function -> stats, errors)."""
    counter = FuncCounter(cfg.name_ok)
    file = str(path)
    try:
        # Raw formulas: the _xlfn./_xll. prefixes tell built-ins from custom functions.
        with ScopedWorkbook(path, cfg.scope, raw_formula=True, formulas_only=True) as book:
            names = {name for _, name in book.defined_names}
            for part in book.parts():
                group = None if cfg.by is None else file if cfg.by == "file" else (file, part.sheet_name)
                formulas = [c.formula for c in part.cells if c.formula] + [o.text for o in part.objects if o.is_formula]
                for formula in formulas:
                    counter.add(group, file, formula, names)
    except READ_ERRORS as exc:
        return counter.groups, [f"{path}: cannot read workbook ({exc})"]
    return counter.groups, []


# ---------------------------------------------------------------- running


def choose_jobs(requested: int | None, paths: list[Path]) -> int:
    """Worker count: as requested, else parallel only for several, sizable files."""
    if requested is not None:
        return max(1, requested)
    if len(paths) < 2:
        return 1
    try:
        total = sum(p.stat().st_size for p in paths)
    except OSError:
        total = PARALLEL_MIN_BYTES
    if total < PARALLEL_MIN_BYTES:
        return 1
    return max(1, min(len(paths), os.cpu_count() or 1))


C = TypeVar("C")  # config type
R = TypeVar("R")  # result type


class Runner(Generic[C, R]):
    """Applies ``work(path, cfg)`` to paths, yielding results in input order."""

    def __init__(self, work: Callable[[Path, C], R], cfg: C, jobs: int):
        self.work = work
        self.cfg = cfg
        self.jobs = jobs
        self._executor: ProcessPoolExecutor | None = None

    def results(self, paths: list[Path]) -> Iterator[R]:
        if self.jobs <= 1:
            for path in paths:
                yield self.work(path, self.cfg)
            return
        self._executor = ProcessPoolExecutor(max_workers=self.jobs)
        futures: list[Future] = [self._executor.submit(self.work, path, self.cfg) for path in paths]
        for future in futures:
            yield future.result()

    def close(self) -> None:
        """Stop pending work (e.g. -q found a match, or Ctrl-C)."""
        if self._executor is not None:
            self._executor.shutdown(wait=False, cancel_futures=True)
            self._executor = None
