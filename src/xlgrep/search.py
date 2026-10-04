"""Per-file work and running it across files, in parallel when worthwhile.

Everything a worker process needs travels in a picklable config; each file's
output is rendered into a string so the parent can print results in file order.
"""

from __future__ import annotations

import fnmatch
import io
import os
import zipfile
from collections.abc import Callable, Iterator
from concurrent.futures import Future, ProcessPoolExecutor
from dataclasses import dataclass, field
from pathlib import Path
from typing import TypeVar
from xml.etree import ElementTree as ET

from openpyxl.utils.exceptions import InvalidFileException

from .address import CellRange
from .funcs import FuncCounter, FuncStat
from .matcher import Matcher, merge_spans
from .objects import SheetObject, WorkbookObjects, read_objects
from .output import (
    CsvFormatter,
    Formatter,
    JsonFormatter,
    LineFormatter,
    Match,
    ObjectMatch,
    OutputOptions,
    PrettyFormatter,
)
from .refs import RefFinder, RefHit, Target
from .workbook import Cell, Sheet, read_sheets

READ_ERRORS = (zipfile.BadZipFile, InvalidFileException, KeyError, OSError, ET.ParseError)

# Below this much input, starting worker processes (each imports openpyxl) costs
# more than it saves.
PARALLEL_MIN_BYTES = 1_000_000


@dataclass
class Scope:
    """Which sheets, cells and objects to look at."""

    objects: set[str]  # "cell", "name", "cf", "dv", "note"
    sheet_globs: list[str] = field(default_factory=list)
    no_hidden: bool = False
    cell_range: CellRange | None = None
    raw_formula: bool = False

    def sheet_ok(self, name: str, hidden: bool) -> bool:
        if self.no_hidden and hidden:
            return False
        return not self.sheet_globs or any(fnmatch.fnmatch(name, g) for g in self.sheet_globs)

    @property
    def object_kinds(self) -> set[str]:
        return self.objects - {"cell"}


@dataclass
class SearchConfig:
    scope: Scope
    matcher: Matcher
    targets: list[Target]
    search_in: str  # "auto" | "formula" | "value"
    invert: bool
    output: OutputOptions
    mode: str  # "line" | "pretty" | "json" | "csv"
    with_values: bool
    summary_only: bool  # -l / -c / -q: only the count matters
    limit: int | None  # max matching cells/objects per file


@dataclass
class FileResult:
    count: int = 0
    output: str = ""
    errors: list[str] = field(default_factory=list)
    printed_group: bool = False  # line mode with context printed a group (for "--" separators)


class Searcher:
    def __init__(self, cfg: SearchConfig):
        self.matcher = cfg.matcher
        self.search_in = cfg.search_in
        self.invert = cfg.invert
        self.cell_range = cfg.scope.cell_range
        self.object_kinds = cfg.scope.object_kinds
        self.targets = cfg.targets
        self.ref_finder: RefFinder | None = None  # set per workbook when --ref is used

    def start_workbook(self, objects: WorkbookObjects | None) -> None:
        if not self.targets or objects is None:
            return
        names = {(o.sheet, o.ref.upper()): o.text for o in objects.workbook_names}
        names.update({(o.sheet, o.ref.upper()): o.text
                      for lst in objects.by_sheet.values() for o in lst if o.object == "name"})
        self.ref_finder = RefFinder(self.targets, [s.name for s in objects.sheets], names)

    def searched_text(self, cell: Cell) -> tuple[str, str] | None:
        """The (text, kind) to match against, or None if the cell is out of scope."""
        if self.search_in == "formula":
            return (cell.formula, "formula") if cell.is_formula else None
        if self.search_in == "value":
            return (cell.value_str, "value") if cell.has_cached_value else None
        return (cell.formula, "formula") if cell.is_formula else (cell.value_str, "value")

    def _select(self, text: str, kind: str, sheet: str | None,
                applies_to: list[CellRange] | None = None) -> tuple[list, list[RefHit] | None] | None:
        """(spans to highlight, --ref hits) if the text is selected, else None. Honours -v."""
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

    def search_objects(self, objects: list[SheetObject], limit: int | None) -> list[ObjectMatch]:
        matches: list[ObjectMatch] = []
        for obj in objects:
            if limit is not None and len(matches) >= limit:
                break
            if obj.object not in self.object_kinds:
                continue
            if self.search_in == ("value" if obj.is_formula else "formula"):
                continue
            if self.cell_range and not obj.in_range(self.cell_range):
                continue
            selected = self._select(obj.text, "formula" if obj.is_formula else "value", obj.sheet,
                                    obj.areas if obj.object in ("cf", "dv") else None)
            if selected is not None:
                matches.append(ObjectMatch(obj, *selected))
        return matches

    def search(self, sheet: Sheet, limit: int | None) -> list[Match]:
        matches: list[Match] = []
        for cell in sheet.sorted_cells():
            if limit is not None and len(matches) >= limit:
                break
            if self.cell_range and not self.cell_range.contains(cell.row, cell.col):
                continue
            searched = self.searched_text(cell)
            if searched is None or searched[0] == "":
                continue
            text, kind = searched
            selected = self._select(text, kind, sheet.name)
            if selected is not None:
                matches.append(Match(cell, text, kind, *selected))
        return matches


_FORMATTERS = {"line": LineFormatter, "pretty": PrettyFormatter, "json": JsonFormatter, "csv": CsvFormatter}


def search_file(path: Path, cfg: SearchConfig) -> FileResult:
    """Search one workbook; the result carries its rendered output."""
    buf = io.StringIO()
    formatter: Formatter = _FORMATTERS[cfg.mode](cfg.output, buf)
    searcher = Searcher(cfg)
    scope = cfg.scope
    # --ref needs the sheet order and defined names even when names aren't searched.
    read_kinds = scope.object_kinds | ({"name"} if cfg.targets else set())
    result = FileResult()

    def remaining() -> int | None:
        return None if cfg.limit is None else cfg.limit - result.count

    def done() -> bool:
        left = remaining()
        return left is not None and left <= 0

    def emit_objects(sheet_name: str | None, hidden: bool, objects: list[SheetObject]) -> None:
        if not objects or done():
            return
        matches = searcher.search_objects(objects, remaining())
        result.count += len(matches)
        if matches and not cfg.summary_only:
            formatter.write_objects(path, sheet_name, hidden, matches)

    try:
        wb_objects = read_objects(path, read_kinds, raw_formula=scope.raw_formula) if read_kinds else None
        searcher.start_workbook(wb_objects)
        seen: set[str] = set()
        if "cell" in scope.objects:
            for sheet in read_sheets(path, with_values=cfg.with_values and not cfg.summary_only,
                                     raw_formula=scope.raw_formula, sheet_filter=scope.sheet_ok):
                seen.add(sheet.name)
                if done():
                    break
                matches = searcher.search(sheet, remaining())
                result.count += len(matches)
                if matches and not cfg.summary_only:
                    formatter.write_sheet(path, sheet, matches)
                if wb_objects is not None:
                    emit_objects(sheet.name, sheet.hidden, wb_objects.by_sheet.get(sheet.name, []))
        if wb_objects is not None:
            for info in wb_objects.sheets:
                if info.name not in seen and scope.sheet_ok(info.name, info.hidden):
                    emit_objects(info.name, info.hidden, wb_objects.by_sheet[info.name])
            # Workbook-scoped names belong to no sheet or range.
            if "name" in scope.objects and not scope.sheet_globs and not scope.cell_range:
                emit_objects(None, False, wb_objects.workbook_names)
    except READ_ERRORS as exc:
        result.errors.append(f"{path}: cannot read workbook ({exc})")
    if result.count and not cfg.summary_only:
        formatter.end_file(path, result.count)
    result.output = buf.getvalue()
    result.printed_group = getattr(formatter, "_printed_group", False)
    return result


@dataclass
class FuncConfig:
    scope: Scope
    pattern: object  # compiled name regex or None
    wanted: set[str]  # upper-cased -f names
    by: str | None


def count_file(path: Path, cfg: FuncConfig) -> tuple[dict[object, dict[str, FuncStat]], list[str]]:
    """--list-funcs for one workbook: (group -> name -> stats, errors)."""
    def name_filter(name: str) -> bool:
        if cfg.pattern is None and not cfg.wanted:
            return True
        return bool(cfg.pattern is not None and cfg.pattern.search(name)) or name.upper() in cfg.wanted

    counter = FuncCounter(name_filter)
    scope = cfg.scope
    file = str(path)
    kinds = scope.objects & {"name", "cf", "dv"}

    def group(sheet: str | None):
        return None if cfg.by is None else file if cfg.by == "file" else (file, sheet)

    try:
        objs = read_objects(path, kinds | {"name"}, raw_formula=True)
        names = {o.ref.upper() for o in objs.workbook_names}
        names |= {o.ref.upper() for lst in objs.by_sheet.values() for o in lst if o.object == "name"}
        for info in objs.sheets:
            if not scope.sheet_ok(info.name, info.hidden):
                continue
            for obj in objs.by_sheet[info.name]:
                if obj.object in kinds and (not scope.cell_range or obj.in_range(scope.cell_range)):
                    counter.add(group(info.name), file, obj.text, names)
        if "name" in kinds and not scope.sheet_globs and not scope.cell_range:
            for obj in objs.workbook_names:
                counter.add(group(None), file, obj.text, names)
        if "cell" in scope.objects:
            for sheet in read_sheets(path, with_values=False, raw_formula=True, sheet_filter=scope.sheet_ok):
                for cell in sheet.sorted_cells():
                    if cell.is_formula and (not scope.cell_range or scope.cell_range.contains(cell.row, cell.col)):
                        counter.add(group(sheet.name), file, cell.formula, names)
    except READ_ERRORS as exc:
        return counter.groups, [f"{path}: cannot read workbook ({exc})"]
    return counter.groups, []


T = TypeVar("T")
C = TypeVar("C")


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


class Runner:
    """Applies ``work(path, cfg)`` to paths, yielding results in input order."""

    def __init__(self, work: Callable[[Path, C], T], cfg: C, jobs: int):
        self.work = work
        self.cfg = cfg
        self.jobs = jobs
        self._executor: ProcessPoolExecutor | None = None

    def results(self, paths: list[Path]) -> Iterator[T]:
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
