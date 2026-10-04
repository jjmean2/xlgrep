"""Result formatters: grep-like lines, pretty grids, JSON Lines and CSV."""

from __future__ import annotations

import csv
import datetime as dt
import json
import sys
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
from typing import TextIO

from .address import cell_name, col_letter, qualified
from .funcs import BUILTIN, FuncStat
from .matcher import Span
from .objects import SheetObject
from .refs import RefHit
from .text import char_width, display_width, escape, escape_spans
from .workbook import Cell, Sheet


@dataclass
class Match:
    cell: Cell
    text: str  # the text that was searched
    kind: str  # "formula" or "value"
    spans: list[Span]
    refs: list[RefHit] | None = None  # --ref hits


@dataclass
class ObjectMatch:
    obj: SheetObject
    spans: list[Span]
    refs: list[RefHit] | None = None  # --ref hits

    @property
    def text(self) -> str:
        return self.obj.text

    @property
    def kind(self) -> str:
        return "formula" if self.obj.is_formula else "value"


ContentFn = Callable[[Cell], str]


class Style:
    RESET = "\x1b[0m"

    def __init__(self, enabled: bool):
        self.enabled = enabled

    def _wrap(self, code: str, text: str) -> str:
        return f"\x1b[{code}m{text}{self.RESET}" if self.enabled and text else text

    def path(self, text: str) -> str:
        return self._wrap("35", text)

    def heading(self, text: str) -> str:
        return self._wrap("1;35", text)

    def sheet(self, text: str) -> str:
        return self._wrap("1;36", text)

    def addr(self, text: str) -> str:
        return self._wrap("32", text)

    def header(self, text: str) -> str:
        return self._wrap("33", text)

    def match(self, text: str) -> str:
        return self._wrap("1;31", text)

    def dim(self, text: str) -> str:
        return self._wrap("2", text)

    def sep(self, text: str) -> str:
        return self._wrap("36", text)

    def highlight(self, text: str, spans: list[Span]) -> str:
        if not spans or not self.enabled:
            return text
        out, pos = [], 0
        for start, end in spans:
            out.append(text[pos:start])
            out.append(self.match(text[start:end]))
            pos = end
        out.append(text[pos:])
        return "".join(out)


@dataclass
class OutputOptions:
    content: ContentFn  # how a (context) cell is displayed
    style: Style
    show_value: bool = False
    header_row: int | None = None
    before: int = 0
    after: int = 0
    row_context: bool = False
    col_context: int = 1
    only_matching: bool = False
    with_filename: bool = True
    max_width: int = 40


class Formatter:
    def __init__(self, opts: OutputOptions, out: TextIO | None = None):
        self.opts = opts
        self.out = out if out is not None else sys.stdout

    def write_sheet(self, path: Path, sheet: Sheet, matches: list[Match]) -> None:
        raise NotImplementedError

    def write_objects(self, path: Path, sheet: str | None, hidden: bool, matches: list[ObjectMatch]) -> None:
        """Matches outside cells; ``sheet`` is None for workbook-scoped names."""
        raise NotImplementedError

    def end_file(self, path: Path, count: int) -> None:
        pass

    def finish(self) -> None:
        pass

    def _value_suffix(self, cell: Cell, content: str) -> str:
        if self.opts.show_value and cell.is_formula and cell.has_cached_value and content == cell.formula:
            return " → " + escape(cell.value_str)
        return ""

    def _header_label(self, sheet: Sheet, col: int) -> str | None:
        if self.opts.header_row is None:
            return None
        cell = sheet.get(self.opts.header_row, col)
        if cell is None:
            return None
        return escape(cell.value_str if cell.is_formula and cell.has_cached_value else cell.display)


class LineFormatter(Formatter):
    """One line per cell: ``path:Sheet!B12:content``; context lines use ``-``."""

    def __init__(self, opts: OutputOptions, out: TextIO | None = None):
        super().__init__(opts, out)
        self._printed_group = False

    @property
    def _has_context(self) -> bool:
        return bool(self.opts.before or self.opts.after or self.opts.row_context)

    def write_sheet(self, path: Path, sheet: Sheet, matches: list[Match]) -> None:
        if self.opts.only_matching:
            for m in matches:
                for start, end in m.spans:
                    self._line(path, sheet, m.cell, ":", self.opts.style.match(escape(m.text[start:end])))
            return
        if not self._has_context:
            for m in matches:
                self._match_line(path, sheet, m)
            return

        by_pos = {(m.cell.row, m.cell.col): m for m in matches}
        for group in self._context_groups(sheet, matches):
            if self._printed_group:
                self.out.write(self.opts.style.sep("--") + "\n")
            self._printed_group = True
            for pos in group:
                if pos in by_pos:
                    self._match_line(path, sheet, by_pos[pos])
                else:
                    cell = sheet.cells[pos]
                    content = self.opts.content(cell)
                    self._line(path, sheet, cell, "-", escape(content) + self._value_suffix(cell, content))

    def write_objects(self, path: Path, sheet: str | None, hidden: bool, matches: list[ObjectMatch]) -> None:
        st = self.opts.style
        for m in matches:
            location = st.addr(m.obj.location) + st.header(f"#{m.obj.object}")
            prefix = f"{st.path(str(path))}{st.sep(':')}" if self.opts.with_filename else ""
            if self.opts.only_matching:
                bodies = [st.match(escape(m.text[s:e])) for s, e in m.spans]
            else:
                bodies = [st.highlight(escape(m.text), escape_spans(m.text, m.spans))]
            for body in bodies:
                self.out.write(f"{prefix}{location}{st.sep(':')}{body}\n")

    def _match_line(self, path: Path, sheet: Sheet, m: Match) -> None:
        body = self.opts.style.highlight(escape(m.text), escape_spans(m.text, m.spans))
        self._line(path, sheet, m.cell, ":", body + self._value_suffix(m.cell, m.text))

    def _line(self, path: Path, sheet: Sheet, cell: Cell, sep: str, body: str) -> None:
        st = self.opts.style
        addr = st.addr(qualified(sheet.name, cell.row, cell.col))
        label = self._header_label(sheet, cell.col)
        if label is not None:
            addr += st.header(f"[{label}]")
        prefix = f"{st.path(str(path))}{st.sep(sep)}" if self.opts.with_filename else ""
        self.out.write(f"{prefix}{addr}{st.sep(sep)}{body}\n")

    def _context_groups(self, sheet: Sheet, matches: list[Match]) -> list[list[tuple[int, int]]]:
        """Group match windows whose row ranges overlap; each group in reading order."""
        windows: list[tuple[int, int, set[tuple[int, int]]]] = []
        for m in matches:
            r, c = m.cell.row, m.cell.col
            lo, hi = max(1, r - self.opts.before), r + self.opts.after
            positions = {(row, c) for row in range(lo, hi + 1)}
            if self.opts.row_context:
                positions |= {pos for pos in sheet.cells if pos[0] == r}
            positions = {pos for pos in positions if pos in sheet.cells}
            positions.add((r, c))
            windows.append((lo, hi, positions))
        windows.sort(key=lambda w: w[0])

        groups: list[tuple[int, int, set[tuple[int, int]]]] = []
        for lo, hi, positions in windows:
            if groups and lo <= groups[-1][1] + 1:
                g_lo, g_hi, g_pos = groups[-1]
                groups[-1] = (g_lo, max(g_hi, hi), g_pos | positions)
            else:
                groups.append((lo, hi, positions))
        return [sorted(g[2]) for g in groups]


class PrettyFormatter(Formatter):
    """Group by file and sheet, and render the neighbourhood of matches as a grid."""

    def __init__(self, opts: OutputOptions, out: TextIO | None = None):
        super().__init__(opts, out)
        self._current_file: Path | None = None
        self._current_sheet: tuple[str | None] | None = None
        self._first_file = True

    def _enter(self, path: Path, sheet: str | None, hidden: bool) -> bool:
        """Print file/sheet headings as needed. Returns True if the sheet was already open."""
        st = self.opts.style
        if path != self._current_file:
            if not self._first_file:
                self.out.write("\n")
            self._first_file = False
            self._current_file = path
            self._current_sheet = None
            self.out.write(st.heading(str(path)) + "\n")
        if self._current_sheet == (sheet,):
            return True
        self._current_sheet = (sheet,)
        title = "(workbook)" if sheet is None else sheet + (" (hidden)" if hidden else "")
        self.out.write("  " + st.sheet(title) + "\n")
        return False

    def write_sheet(self, path: Path, sheet: Sheet, matches: list[Match]) -> None:
        self._enter(path, sheet.name, sheet.hidden)
        for i, block in enumerate(self._blocks(sheet, matches)):
            if i:
                self.out.write("\n")
            self._render_block(sheet, *block)

    def write_objects(self, path: Path, sheet: str | None, hidden: bool, matches: list[ObjectMatch]) -> None:
        st = self.opts.style
        if self._enter(path, sheet, hidden):
            self.out.write("\n")
        tags = [f"#{m.obj.object}" for m in matches]
        tag_w = max(len(t) for t in tags)
        ref_w = max(display_width(m.obj.ref) for m in matches)
        for tag, m in zip(tags, matches):
            text, spans = escape(m.text), escape_spans(m.text, m.spans)
            ref = st.addr(m.obj.ref) + " " * (ref_w - display_width(m.obj.ref))
            self.out.write(f"    {st.header(tag.ljust(tag_w))}  {ref}  {st.highlight(text, spans)}\n")

    def end_file(self, path: Path, count: int) -> None:
        if count:
            self.out.write(self.opts.style.dim(f"  {count} match{'es' if count != 1 else ''}") + "\n")

    def _blocks(self, sheet: Sheet, matches: list[Match]):
        """Yield (rows, cols, matches_by_pos) for groups of matches with overlapping row windows."""
        windows = sorted(
            ((max(1, m.cell.row - self.opts.before), m.cell.row + self.opts.after, m) for m in matches),
            key=lambda w: (w[0], w[1]),  # Match itself isn't orderable; same-row matches tie here
        )
        groups: list[list] = []
        for lo, hi, m in windows:
            if groups and lo <= groups[-1][1] + 1:
                groups[-1][1] = max(groups[-1][1], hi)
                groups[-1][2].append(m)
            else:
                groups.append([lo, hi, [m]])

        for lo, hi, group in groups:
            rows = list(range(lo, hi + 1))
            cols: set[int] = set()
            for m in group:
                c = m.cell.col
                cols.update(range(max(1, c - self.opts.col_context), c + self.opts.col_context + 1))
            if self.opts.row_context:
                match_rows = {m.cell.row for m in group}
                cols.update(col for (row, col) in sheet.cells if row in match_rows)
            yield rows, sorted(cols), {(m.cell.row, m.cell.col): m for m in group}

    def _render_block(self, sheet: Sheet, rows: list[int], cols: list[int], by_pos: dict) -> None:
        st = self.opts.style
        header_row = self.opts.header_row
        show_header = header_row is not None and header_row not in rows
        all_rows = ([header_row] if show_header else []) + rows

        # Cell texts as (plain, spans), truncated to max_width.
        grid: dict[tuple[int, int], tuple[str, list[Span]]] = {}
        for r in all_rows:
            for c in cols:
                m = by_pos.get((r, c))
                cell = sheet.get(r, c)
                if m is not None:
                    text, spans = escape(m.text), escape_spans(m.text, m.spans)
                    text += self._value_suffix(m.cell, m.text)
                elif cell is not None:
                    content = self.opts.content(cell)
                    text, spans = escape(content) + self._value_suffix(cell, content), []
                else:
                    text, spans = "", []
                grid[(r, c)] = _truncate(text, spans, self.opts.max_width)

        widths = {
            c: max([display_width(col_letter(c))] + [display_width(grid[(r, c)][0]) for r in all_rows])
            for c in cols
        }
        num_w = max(len(str(r)) for r in all_rows)
        bar = st.dim("│")

        letters = f" {bar} ".join(_pad(col_letter(c), widths[c]) for c in cols)
        self.out.write(f"    {' ' * num_w} {bar} {st.dim(letters)}\n".rstrip() + "\n")

        match_rows = {r for (r, _c) in by_pos}
        for r in all_rows:
            marker = st.match("▶") if r in match_rows else " "
            cells = []
            for c in cols:
                text, spans = grid[(r, c)]
                shown = st.highlight(text, spans)
                if r == header_row:
                    shown = st.header(text)
                cells.append(shown + " " * (widths[c] - display_width(text)))
            line = f"  {marker} {st.addr(str(r).rjust(num_w))} {bar} " + f" {bar} ".join(cells)
            self.out.write(line.rstrip() + "\n")
            if show_header and r == header_row:
                rule = "┼".join("─" * (widths[c] + 2) for c in cols)
                self.out.write(st.dim(f"    {'─' * num_w}─┼{rule}") + "\n")


def _pad(text: str, width: int) -> str:
    return text + " " * (width - display_width(text))


def _truncate(text: str, spans: list[Span], max_width: int) -> tuple[str, list[Span]]:
    if max_width <= 0 or display_width(text) <= max_width:
        return text, spans
    width, cut = 0, 0
    for i, ch in enumerate(text):
        w = char_width(ch)
        if width + w > max_width - 1:
            break
        width += w
        cut = i + 1
    clipped = [(s, min(e, cut)) for s, e in spans if s < cut]
    return text[:cut] + "…", clipped


def _json_value(value: object) -> object:
    if value is None or isinstance(value, (bool, int, float, str)):
        return value
    if isinstance(value, (dt.date, dt.time)):
        return value.isoformat()
    return str(value)


def _add_refs(record: dict, refs: list[RefHit] | None) -> None:
    if refs is not None:
        record["refs"] = [
            {"start": h.start, "end": h.end, "text": h.text, **({"via": h.via} if h.via else {})} for h in refs
        ]


class JsonFormatter(Formatter):
    def write_sheet(self, path: Path, sheet: Sheet, matches: list[Match]) -> None:
        for m in matches:
            cell = m.cell
            record = {
                "file": str(path),
                "object": "cell",
                "sheet": sheet.name,
                "cell": cell_name(cell.row, cell.col),
                "row": cell.row,
                "col": cell.col,
                "kind": m.kind,
                "formula": cell.formula,
                "value": _json_value(cell.value) if cell.has_cached_value else None,
                "text": m.text,
                "matches": [{"start": s, "end": e, "text": m.text[s:e]} for s, e in m.spans],
            }
            if sheet.hidden:
                record["hidden_sheet"] = True
            _add_refs(record, m.refs)
            self.out.write(json.dumps(record, ensure_ascii=False) + "\n")

    def write_objects(self, path: Path, sheet: str | None, hidden: bool, matches: list[ObjectMatch]) -> None:
        for m in matches:
            record = {
                "file": str(path),
                "object": m.obj.object,
                "sheet": sheet,
                "ref": m.obj.ref,
                "kind": m.kind,
                "text": m.text,
                "matches": [{"start": s, "end": e, "text": m.text[s:e]} for s, e in m.spans],
                **m.obj.detail,
            }
            if hidden:
                record["hidden_sheet"] = True
            _add_refs(record, m.refs)
            self.out.write(json.dumps(record, ensure_ascii=False) + "\n")


class CsvFormatter(Formatter):
    def __init__(self, opts: OutputOptions, out: TextIO | None = None):
        super().__init__(opts, out)
        self.writer = csv.writer(self.out, lineterminator="\n")
        self.writer.writerow(["file", "sheet", "object", "location", "kind", "content", "value"])

    def write_sheet(self, path: Path, sheet: Sheet, matches: list[Match]) -> None:
        for m in matches:
            cell = m.cell
            value = cell.value_str if cell.has_cached_value else ""
            row = [str(path), sheet.name, "cell", cell_name(cell.row, cell.col), m.kind, m.text, value]
            self.writer.writerow(row)

    def write_objects(self, path: Path, sheet: str | None, hidden: bool, matches: list[ObjectMatch]) -> None:
        for m in matches:
            self.writer.writerow([str(path), sheet or "", m.obj.object, m.obj.ref, m.kind, m.text, ""])


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
    widths = [max(display_width(r[i]) for r in rows + [headers]) for i in range(len(headers))]

    def fmt(cells: list[str]) -> str:
        first = cells[0] + " " * (widths[0] - display_width(cells[0]))
        return "  ".join([first] + [c.rjust(w) for c, w in zip(cells[1:], widths[1:])])

    out.write(indent + style.dim(fmt(headers)) + "\n")
    for s, row in zip(stats, rows):
        tag = "" if s.category == BUILTIN else "  " + style.header(s.category)
        out.write(indent + fmt(row) + tag + "\n")
