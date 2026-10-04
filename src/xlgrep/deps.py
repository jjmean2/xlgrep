"""--deps: what each workbook depends on — other workbooks, its own sheets, data
connections.

deps_file runs per workbook (possibly in a worker process) and records references
as stored: other workbooks by their [n] number. The parent then resolves those to
files (resolve_target), knowing every file that was searched, and renders.

A dependency is counted once per formula: a formula reading Data!A1 and Data!B2 is
one formula depending on Data. Formulas in cells, conditional formats, validations
and defined names all count, and references through defined names are followed.
"""

from __future__ import annotations

import csv
import json
import os
import re
import sys
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path, PureWindowsPath
from typing import TextIO
from urllib.parse import unquote

from .funcs import formula_shape
from .output import Style
from .package import Connection, ExternalBook, connections, external_books
from .refs import RefContext, scan_refs
from .scope import READ_ERRORS, Scope, ScopedWorkbook
from .text import display_width

# (from sheet, book as written in brackets, to sheet); sheets are None when unknown
# or not applicable (workbook-scoped names, a name in another workbook).
FileEdge = tuple[str | None, str, str | None]
SheetEdge = tuple[str | None, str]  # (from sheet, to sheet) within one workbook


@dataclass
class FileDeps:
    path: str
    books: list[ExternalBook | None] = field(default_factory=list)  # [n] is books[n-1]
    file_edges: Counter = field(default_factory=Counter)  # FileEdge -> formulas
    sheet_edges: Counter = field(default_factory=Counter)  # SheetEdge -> formulas
    connections: list[Connection] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)

    def target(self, book: str) -> str:
        """The stored path of [book], or the bracket text itself ([Rates.xlsx]Sheet!A1)."""
        if book.isdigit() and 0 < int(book) <= len(self.books) and self.books[int(book) - 1] is not None:
            return self.books[int(book) - 1].target
        return book


@dataclass
class DepsConfig:
    scope: Scope
    by: str | None  # None/"file", or "sheet": file dependencies per sheet


# ---------------------------------------------------------------- collecting


class _Targets:
    """Where a formula's references lead: {(book or None, sheet or None)}."""

    def __init__(self, context: RefContext):
        self.context = context
        self._names: dict[tuple[str | None, str], set] = {}
        self._shapes: dict[tuple[str, str | None], set] = {}  # (formula shape, sheet) -> targets

    def of_formula(self, formula: str, sheet: str | None) -> set:
        """Targets of a formula on ``sheet``. Rows don't change where a formula leads, so
        formulas filled down a column are resolved once (see funcs.formula_shape)."""
        key = (formula_shape(formula), sheet)
        if key not in self._shapes:
            self._shapes[key] = self.of(formula, sheet)
        return self._shapes[key]

    def of(self, formula: str, sheet: str | None, stack: frozenset = frozenset()) -> set:
        found = set()
        for ref in scan_refs(formula):
            if ref.name is not None:
                if ref.external:  # [1]!Name: a name defined in another workbook
                    found.add((ref.book, None))
                    continue
                key = self.context.name_key(ref, sheet)
                if key is not None and key not in stack:
                    found |= self._name(key, stack)
            elif ref.external:
                found.add((ref.book, ":".join(ref.sheets) if ref.sheets else None))
            else:
                found |= {(None, s) for s in self.context.sheets(ref, sheet)}
        return found

    def _name(self, key: tuple[str | None, str], stack: frozenset) -> set:
        if key not in self._names:
            self._names[key] = self.of(self.context.names[key], self.context.name_context(key), stack | {key})
        return self._names[key]


def deps_file(path: Path, cfg: DepsConfig) -> FileDeps:
    result = FileDeps(str(path))
    try:
        # Only formulas can depend on anything, so value cells aren't read.
        with ScopedWorkbook(path, cfg.scope, raw_formula=False, formulas_only=True) as book:
            result.books = external_books(book.zip, book.package)
            result.connections = connections(book.zip, book.package)
            targets = _Targets(RefContext(book.sheet_names, book.defined_names))
            for part in book.parts():
                formulas = [c.formula for c in part.cells if c.formula]
                formulas += [o.text for o in part.objects if o.is_formula]
                for formula in formulas:
                    for to_book, to_sheet in targets.of_formula(formula, part.sheet_name):
                        if to_book is not None:
                            result.file_edges[(part.sheet_name, to_book, to_sheet)] += 1
                        elif to_sheet != part.sheet_name:
                            result.sheet_edges[(part.sheet_name, to_sheet)] += 1
    except READ_ERRORS as exc:
        result.errors.append(f"{path}: cannot read workbook ({exc})")
    return result


# ---------------------------------------------------------------- resolving


FOUND, MATCHED, MISSING = "found", "matched by name", "missing"
_WINDOWS_ABSOLUTE = re.compile(r"^(?:[A-Za-z]:[\\/]|\\\\)")


def resolve_target(raw: str, source: Path, searched: list[Path]) -> tuple[str, str]:
    """(path to show, status) for a stored external workbook path.

    Relative paths resolve against the source workbook's folder. A path that doesn't
    exist here (often a Windows path from another computer) is matched by file name
    against the searched workbooks, if exactly one has that name.
    """
    text = unquote(raw)
    if text.lower().startswith("file:///"):
        text = text[len("file:///"):]
    elif text.lower().startswith("file://"):  # file://server/share/x.xlsx
        text = "\\\\" + text[len("file://"):]
    if not _WINDOWS_ABSOLUTE.match(text):
        local = Path(text) if Path(text).is_absolute() else source.parent / text.replace("\\", "/")
        if local.exists():
            return _shown(local), FOUND
    name = PureWindowsPath(text).name.lower()  # understands both / and \\
    matches = [p for p in searched if p.name.lower() == name]
    if len(matches) == 1:
        return _shown(matches[0]), MATCHED
    return text, MISSING


def _shown(path: Path) -> str:
    try:
        return os.path.relpath(path)
    except ValueError:  # another drive on Windows
        return str(path)


@dataclass
class Edge:
    """A rendered dependency: from a workbook (and sheet) to a workbook, sheet or connection."""

    source: str
    from_sheet: str | None
    target: str
    to_sheet: str | None
    kind: str  # "file" | "sheet" | "data"
    formulas: int = 0
    status: str | None = None  # file edges: found / matched by name / missing
    detail: str | None = None  # data edges: connection type


def edges(deps: FileDeps, by: str | None, searched: list[Path]) -> list[Edge]:
    """The dependencies of one workbook, resolved and ordered: files, sheets, data."""
    out: list[Edge] = []
    files: Counter = Counter()
    for (from_sheet, book, to_sheet), n in deps.file_edges.items():
        key = (from_sheet, book, to_sheet) if by == "sheet" else (None, book, None)
        files[key] += n
    for (from_sheet, book, to_sheet), n in files.items():
        target, status = resolve_target(deps.target(book), Path(deps.path), searched)
        out.append(Edge(deps.path, from_sheet, target, to_sheet, "file", n, status))
    out.sort(key=lambda e: (e.target, e.from_sheet or "", e.to_sheet or ""))
    sheets = [Edge(deps.path, a, deps.path, b, "sheet", n) for (a, b), n in deps.sheet_edges.items()]
    out += sorted(sheets, key=lambda e: -e.formulas)
    out += [Edge(deps.path, None, c.name, None, "data", detail=c.kind) for c in deps.connections]
    return out


# ---------------------------------------------------------------- rendering


def write_deps(all_edges: list[list[Edge]], by: str | None, mode: str, style: Style,
               out: TextIO | None = None) -> None:
    """``all_edges`` has one list per workbook (empty if it has no dependencies)."""
    out = out if out is not None else sys.stdout
    if mode == "json":
        for e in (e for file_edges in all_edges for e in file_edges):
            out.write(json.dumps(_record(e), ensure_ascii=False) + "\n")
    elif mode == "csv":
        writer = csv.DictWriter(out, fieldnames=list(_record(Edge("", None, "", None, "")).keys()),
                                lineterminator="\n")
        writer.writeheader()
        writer.writerows(_record(e) for file_edges in all_edges for e in file_edges)
    elif mode == "mermaid":
        _write_mermaid(all_edges, by, out)
    else:
        _write_text(all_edges, by, style, out)


def _record(e: Edge) -> dict:
    return {"from": e.source, "from_sheet": e.from_sheet, "to": e.target, "to_sheet": e.to_sheet,
            "kind": e.kind, "formulas": e.formulas if e.kind != "data" else None,
            "status": e.status, "type": e.detail}


def _formulas(n: int) -> str:
    return f"{n:,} formula{'s' if n != 1 else ''}"


def _write_text(all_edges: list[list[Edge]], by: str | None, style: Style, out: TextIO) -> None:
    shown = [file_edges for file_edges in all_edges if file_edges]
    for i, file_edges in enumerate(shown):
        if i:
            out.write("\n")
        out.write(style.heading(file_edges[0].source) + "\n")
        file_lines = [e for e in file_edges if e.kind == "file"]
        sheet_lines = [e for e in file_edges if e.kind == "sheet"]
        rows = [(_file_label(e, by), e) for e in file_lines]
        if by == "sheet":
            rows += [(_sheet_label(e), e) for e in sheet_lines]
        width = max((display_width(label) for label, _ in rows), default=0)
        for label, e in rows:
            note = "" if e.status in (None, FOUND) else "  " + style.dim(f"({e.status})")
            padding = " " * (width - display_width(label))
            out.write(f"  {label}{padding}  {_formulas(e.formulas).rjust(14)}{note}\n")
        if sheet_lines and by != "sheet":
            summary = " · ".join(f"{_sheet_label(e)} {e.formulas:,}" for e in sheet_lines)
            out.write(f"  {style.header('sheets')}  {summary}\n")
        data = [e for e in file_edges if e.kind == "data"]
        if data:
            out.write(f"  {style.header('data')}    " + " · ".join(f"{e.target} ({e.detail})" for e in data) + "\n")
    if len(all_edges) > 1:
        out.write(style.dim(f"{len(shown)} of {len(all_edges)} workbooks have dependencies") + "\n")


def _file_label(e: Edge, by: str | None) -> str:
    target = e.target + (f" : {e.to_sheet}" if by == "sheet" and e.to_sheet else "")
    return f"{e.from_sheet or '(workbook)'} → {target}" if by == "sheet" else f"→ {target}"


def _sheet_label(e: Edge) -> str:
    return f"{e.from_sheet or '(workbook)'} → {e.to_sheet}"


def _write_mermaid(all_edges: list[list[Edge]], by: str | None, out: TextIO) -> None:
    """A Mermaid flowchart of workbooks; with --by sheet, sheets boxed by workbook."""
    graph = _Mermaid(by == "sheet")
    for e in (e for file_edges in all_edges for e in file_edges):
        src = graph.workbook(e.source) if e.kind == "data" else graph.place(e.source, e.from_sheet or "(workbook)")
        if e.kind == "file":
            missing = e.status == MISSING
            dst = graph.place(e.target, e.to_sheet or "?", missing)
            graph.edge(src, dst, e.formulas, dashed=missing)
        elif e.kind == "sheet" and by == "sheet":
            graph.edge(src, graph.place(e.source, e.to_sheet), e.formulas)
        elif e.kind == "data":
            graph.edge(src, graph.data(e.source, f"{e.target} ({e.detail})"))
    out.write(graph.render())


class _Mermaid:
    """Collects nodes and edges, then writes them (subgraphs need their nodes up front)."""

    def __init__(self, sheets: bool):
        self.sheets = sheets
        self.ids: dict[tuple, str] = {}
        self.nodes: list[tuple[str, str, str | None, bool]] = []  # id, text, workbook box, missing
        self.boxes: dict[str, tuple[str, str]] = {}  # workbook -> (box id, label)
        self.data_ids: set[str] = set()  # drawn as cylinders
        self.edges: list[str] = []

    def _node(self, key: tuple, text: str, box: str | None = None, missing: bool = False) -> str:
        if key not in self.ids:
            self.ids[key] = f"n{len(self.ids)}"
            self.nodes.append((self.ids[key], text, box, missing))
        return self.ids[key]

    def _label(self, workbook: str, missing: bool) -> str:
        return Path(workbook.replace("\\", "/")).name + (" (missing)" if missing else "")

    def workbook(self, path: str, missing: bool = False) -> str:
        """A workbook as a whole: its box with --by sheet, else its node."""
        if self.sheets:
            if path not in self.boxes:
                self.boxes[path] = (f"w{len(self.boxes)}", self._label(path, missing))
            return self.boxes[path][0]
        return self._node(("book", path), self._label(path, missing), missing=missing)

    def place(self, path: str, sheet: str | None, missing: bool = False) -> str:
        """A sheet of a workbook with --by sheet, else the workbook."""
        if not self.sheets:
            return self.workbook(path, missing)
        self.workbook(path, missing)
        return self._node(("sheet", path, sheet), sheet or "?", box=path, missing=missing)

    def data(self, path: str, text: str) -> str:
        node_id = self._node(("data", path, text), text)
        self.data_ids.add(node_id)
        return node_id

    def edge(self, src: str, dst: str, count: int | None = None, dashed: bool = False) -> None:
        arrow = "-.->" if dashed else "-->"
        self.edges.append(f"  {src} {arrow}{f'|{count}|' if count else ''} {dst}")

    def render(self) -> str:
        lines = ["flowchart LR"]

        def shape(node_id: str, text: str, missing: bool) -> str:
            body = f'[("{_mermaid_text(text)}")]' if node_id in self.data_ids else f'["{_mermaid_text(text)}"]'
            return f"{node_id}{body}" + (":::missing" if missing else "")

        for path, (box_id, label) in self.boxes.items():
            lines.append(f'  subgraph {box_id}["{_mermaid_text(label)}"]')
            lines += [f"    {shape(i, text, m)}" for i, text, box, m in self.nodes if box == path]
            lines.append("  end")
        lines += [f"  {shape(i, text, m)}" for i, text, box, m in self.nodes if box is None]
        lines += self.edges
        lines.append("  classDef missing stroke-dasharray: 5 5")
        return "\n".join(lines) + "\n"


def _mermaid_text(text: str) -> str:
    return text.replace('"', "#quot;")
