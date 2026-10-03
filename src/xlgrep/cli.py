"""Command-line interface."""

from __future__ import annotations

import argparse
import fnmatch
import os
import re
import sys
import zipfile
from collections.abc import Sequence
from xml.etree import ElementTree as ET
from pathlib import Path

from openpyxl.utils.exceptions import InvalidFileException

from . import __version__
from .address import CellRange, parse_range
from .files import UnsupportedFile, iter_files
from .matcher import Matcher, build_matcher
from .objects import OBJECT_KINDS, SheetObject, WorkbookObjects, read_objects
from .output import (
    CsvFormatter,
    Formatter,
    JsonFormatter,
    LineFormatter,
    Match,
    ObjectMatch,
    OutputOptions,
    PrettyFormatter,
    Style,
)
from .workbook import Cell, Sheet, read_sheets

EXIT_MATCH, EXIT_NO_MATCH, EXIT_ERROR = 0, 1, 2

# --objects names -> internal kinds
OBJECT_CHOICES = {"cells": "cell", "names": "name", "cf": "cf", "dv": "dv", "notes": "note"}


def parse_objects(text: str) -> set[str]:
    kinds = set()
    for item in text.split(","):
        item = item.strip().lower()
        if item == "all":
            kinds.update(OBJECT_CHOICES.values())
        elif item in OBJECT_CHOICES:
            kinds.add(OBJECT_CHOICES[item])
        elif item:
            raise argparse.ArgumentTypeError(
                f"unknown object {item!r} (choose from all, {', '.join(OBJECT_CHOICES)})"
            )
    if not kinds:
        raise argparse.ArgumentTypeError("no objects given")
    return kinds


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="xlgrep",
        usage="xlgrep [OPTIONS] PATTERN [PATH ...]\n       xlgrep [OPTIONS] (-e PATTERN | -f FUNCS)... [PATH ...]",
        description="Search the cells of Excel workbooks (.xlsx/.xlsm/.xltx/.xltm) like grep.",
        epilog=(
            "examples:\n"
            "  xlgrep VLOOKUP reports/            formulas and values containing VLOOKUP\n"
            "  xlgrep -f VLOOKUP,XLOOKUP .        cells calling these functions\n"
            "  xlgrep -p -C1 --header 'Total' .   pretty grid with column headers\n"
            "  xlgrep --in value -F '#N/A' .      cells whose computed value is #N/A\n"
            "\n"
            "exit status is 0 if a cell matched, 1 if none did, 2 if an error occurred."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument("args", nargs="*", metavar="PATTERN/PATH", help=argparse.SUPPRESS)

    g = p.add_argument_group("patterns")
    g.add_argument("-e", "--regexp", action="append", default=[], metavar="PATTERN",
                   help="pattern to search for (repeatable); positional args are then all paths")
    g.add_argument("-f", "--func", action="append", default=[], metavar="NAMES",
                   help="find calls to these functions (comma-separated, case-insensitive, "
                        "ignores text inside string literals)")
    g.add_argument("-F", "--fixed-strings", action="store_true", help="treat patterns as literal strings")
    g.add_argument("-i", "--ignore-case", action="store_true", help="case-insensitive matching")
    g.add_argument("-S", "--smart-case", action="store_true",
                   help="case-insensitive unless the pattern has an uppercase letter")
    g.add_argument("-w", "--word-regexp", action="store_true", help="match whole words only")
    g.add_argument("-v", "--invert-match", action="store_true", help="select non-empty cells that do not match")

    g = p.add_argument_group("what to search")
    g.add_argument("--in", dest="search_in", choices=["auto", "formula", "value"], default="auto",
                   help="auto: formula text for formula cells, value otherwise (default); "
                        "formula: formula cells only; value: displayed values, incl. cached formula results")
    g.add_argument("--objects", type=parse_objects, default=set(OBJECT_CHOICES.values()), metavar="LIST",
                   help="where to search, comma-separated: cells, names (defined names), cf (conditional "
                        "formats), dv (data validations), notes (cell notes), or all (default)")
    g.add_argument("--raw-formula", action="store_true",
                   help="keep _xlfn./_xlws./_xlpm. prefixes as stored in the file")

    g = p.add_argument_group("files and cells")
    g.add_argument("-g", "--glob", action="append", default=[], metavar="GLOB",
                   help="only search files whose name matches GLOB; '!GLOB' excludes (repeatable)")
    g.add_argument("--sheet", action="append", default=[], metavar="GLOB",
                   help="only search sheets whose name matches GLOB (repeatable)")
    g.add_argument("--range", dest="cell_range", metavar="RANGE", help="only search cells in RANGE, e.g. A1:F100, B:D, 2:10")
    g.add_argument("--no-hidden", action="store_true", help="skip hidden sheets")

    g = p.add_argument_group("output")
    mode = g.add_mutually_exclusive_group()
    mode.add_argument("-p", "--pretty", action="store_true", help="group by file/sheet and show a grid around matches")
    mode.add_argument("--json", action="store_true", help="one JSON object per match")
    mode.add_argument("--csv", action="store_true", help="CSV with file,sheet,cell,kind,content,value")
    mode.add_argument("-l", "--files-with-matches", action="store_true", help="print only names of matching files")
    mode.add_argument("-c", "--count", action="store_true", help="print the number of matching cells per file")
    mode.add_argument("-q", "--quiet", action="store_true", help="print nothing; exit status only")
    g.add_argument("-o", "--only-matching", action="store_true", help="print only the matched parts")
    g.add_argument("-m", "--max-count", type=int, metavar="N", help="stop after N matching cells per file")
    g.add_argument("--show-value", action="store_true", help="append the cached result to formula cells")
    g.add_argument("--no-filename", action="store_true", help="don't prefix lines with the file name")
    g.add_argument("--color", choices=["auto", "always", "never"], default="auto", help="colorize output (default: auto)")
    g.add_argument("--max-width", type=int, default=40, metavar="N", help="pretty mode: max cell width (default 40, 0 = unlimited)")

    g = p.add_argument_group("context")
    g.add_argument("-A", "--after-context", type=int, metavar="N", help="show N cells below each match (same column)")
    g.add_argument("-B", "--before-context", type=int, metavar="N", help="show N cells above each match (same column)")
    g.add_argument("-C", "--context", type=int, metavar="N", help="show N cells above and below each match")
    g.add_argument("--row", action="store_true", help="show the other cells in each matching row")
    g.add_argument("--header", type=int, nargs="?", const=1, metavar="ROW",
                   help="show the column header taken from ROW (default 1)")
    g.add_argument("--col-context", type=int, default=1, metavar="N", help="pretty mode: columns shown on each side (default 1)")

    p.add_argument("-V", "--version", action="version", version=f"%(prog)s {__version__}")
    return p


def _use_color(choice: str) -> bool:
    if choice != "auto":
        return choice == "always"
    return sys.stdout.isatty() and "NO_COLOR" not in os.environ and os.environ.get("TERM") != "dumb"


class Searcher:
    def __init__(self, matcher: Matcher, args: argparse.Namespace, cell_range: CellRange | None):
        self.matcher = matcher
        self.search_in = args.search_in
        self.invert = args.invert_match
        self.cell_range = cell_range

    def searched_text(self, cell: Cell) -> tuple[str, str] | None:
        """The (text, kind) to match against, or None if the cell is out of scope."""
        if self.search_in == "formula":
            return (cell.formula, "formula") if cell.is_formula else None
        if self.search_in == "value":
            return (cell.value_str, "value") if cell.has_cached_value else None
        return (cell.formula, "formula") if cell.is_formula else (cell.value_str, "value")

    def _spans(self, text: str, kind: str) -> list | None:
        """Spans to report, or None if the text isn't selected (honours -v)."""
        spans = self.matcher.spans(text, is_formula=kind == "formula")
        return spans if bool(spans) != self.invert else None

    def search_objects(self, objects: list[SheetObject], limit: int | None) -> list[ObjectMatch]:
        matches: list[ObjectMatch] = []
        for obj in objects:
            if limit is not None and len(matches) >= limit:
                break
            if self.search_in == ("value" if obj.is_formula else "formula"):
                continue
            if self.cell_range and not obj.in_range(self.cell_range):
                continue
            spans = self._spans(obj.text, "formula" if obj.is_formula else "value")
            if spans is not None:
                matches.append(ObjectMatch(obj, spans))
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
            spans = self._spans(text, kind)
            if spans is not None:
                matches.append(Match(cell, text, kind, spans))
        return matches


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    funcs = [name.strip() for item in args.func for name in item.split(",") if name.strip()]
    if args.regexp or funcs:
        patterns, paths = args.regexp, args.args
    elif args.args:
        patterns, paths = args.args[:1], args.args[1:]
    else:
        parser.error("no pattern given (use PATTERN, -e or -f)")
    if funcs and args.search_in == "value":
        parser.error("-f/--func searches formulas and can't be combined with --in value")

    try:
        matcher = build_matcher(
            patterns,
            funcs=funcs,
            fixed=args.fixed_strings,
            ignore_case=args.ignore_case,
            smart_case=args.smart_case,
            word=args.word_regexp,
        )
    except re.error as exc:
        parser.error(f"invalid pattern: {exc}")
    try:
        cell_range = parse_range(args.cell_range) if args.cell_range else None
    except ValueError as exc:
        parser.error(str(exc))

    context = args.context or 0
    before = args.before_context if args.before_context is not None else context
    after = args.after_context if args.after_context is not None else context

    def content(cell: Cell) -> str:
        if args.search_in == "value" and cell.has_cached_value:
            return cell.value_str
        return cell.display

    opts = OutputOptions(
        content=content,
        style=Style(_use_color(args.color)),
        show_value=args.show_value,
        header_row=args.header,
        before=before,
        after=after,
        row_context=args.row,
        col_context=args.col_context,
        only_matching=args.only_matching,
        with_filename=not args.no_filename,
        max_width=args.max_width,
    )
    formatter: Formatter
    if args.pretty:
        formatter = PrettyFormatter(opts)
    elif args.json:
        formatter = JsonFormatter(opts)
    elif args.csv:
        formatter = CsvFormatter(opts)
    else:
        formatter = LineFormatter(opts)

    with_values = args.search_in == "value" or args.show_value or args.pretty or args.json or args.csv
    summary_only = args.files_with_matches or args.count or args.quiet
    limit = 1 if (args.files_with_matches or args.quiet) else args.max_count

    def sheet_filter(name: str, hidden: bool) -> bool:
        if args.no_hidden and hidden:
            return False
        return not args.sheet or any(fnmatch.fnmatch(name, g) for g in args.sheet)

    searcher = Searcher(matcher, args, cell_range)
    object_kinds = args.objects & set(OBJECT_KINDS)
    any_match = had_error = False
    try:
        for item in iter_files(paths or ["."], args.glob):
            if isinstance(item, UnsupportedFile):
                print(f"xlgrep: {item}", file=sys.stderr)
                had_error = True
                continue
            try:
                count = 0

                def remaining() -> int | None:
                    return None if limit is None else limit - count

                def emit_objects(sheet_name: str | None, hidden: bool, objects: list[SheetObject]) -> None:
                    nonlocal count
                    if not objects or (remaining() is not None and remaining() <= 0):
                        return
                    matches = searcher.search_objects(objects, remaining())
                    count += len(matches)
                    if matches and not summary_only:
                        formatter.write_objects(item, sheet_name, hidden, matches)

                wb_objects: WorkbookObjects | None = None
                if object_kinds:
                    wb_objects = read_objects(item, object_kinds, raw_formula=args.raw_formula)

                done: set[str] = set()
                if "cell" in args.objects:
                    for sheet in read_sheets(item, with_values=with_values and not summary_only,
                                             raw_formula=args.raw_formula, sheet_filter=sheet_filter):
                        done.add(sheet.name)
                        if remaining() is not None and remaining() <= 0:
                            break
                        matches = searcher.search(sheet, remaining())
                        count += len(matches)
                        if matches and not summary_only:
                            formatter.write_sheet(item, sheet, matches)
                        if wb_objects is not None:
                            emit_objects(sheet.name, sheet.hidden, wb_objects.by_sheet.get(sheet.name, []))
                        if count and args.quiet:
                            return EXIT_MATCH
                if wb_objects is not None:
                    for info in wb_objects.sheets:
                        if info.name not in done and sheet_filter(info.name, info.hidden):
                            emit_objects(info.name, info.hidden, wb_objects.by_sheet[info.name])
                    # Workbook-scoped names belong to no sheet or range.
                    if not args.sheet and not cell_range:
                        emit_objects(None, False, wb_objects.workbook_names)
                if count and args.quiet:
                    return EXIT_MATCH
            except (zipfile.BadZipFile, InvalidFileException, KeyError, OSError, ET.ParseError) as exc:
                print(f"xlgrep: {item}: cannot read workbook ({exc})", file=sys.stderr)
                had_error = True
                continue
            if count:
                any_match = True
                if args.files_with_matches:
                    print(opts.style.path(str(item)))
                elif args.count:
                    print(f"{opts.style.path(str(item))}:{count}")
                else:
                    formatter.end_file(item, count)
        formatter.finish()
        sys.stdout.flush()
    except BrokenPipeError:
        # Output piped into e.g. `head`; exit quietly.
        devnull = os.open(os.devnull, os.O_WRONLY)
        os.dup2(devnull, sys.stdout.fileno())
        return EXIT_MATCH if any_match else EXIT_NO_MATCH
    except KeyboardInterrupt:
        return 130

    if had_error:
        return EXIT_ERROR
    return EXIT_MATCH if any_match else EXIT_NO_MATCH


def run() -> None:
    sys.exit(main())
