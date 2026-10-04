"""Command-line interface."""

from __future__ import annotations

import argparse
import csv
import os
import re
import sys
from collections.abc import Sequence
from pathlib import Path

from . import __version__
from .address import CellRange, parse_range
from .files import UnsupportedFile, iter_files
from .funcs import FuncCounter
from .matcher import build_matcher
from .output import CSV_HEADER, OutputOptions, Style, write_func_stats
from .refs import parse_target
from .search import FuncConfig, Runner, Scope, SearchConfig, choose_jobs, count_file, search_file

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
        usage=(
            "xlgrep [OPTIONS] PATTERN [PATH ...]\n"
            "       xlgrep [OPTIONS] (-e PATTERN | -f FUNCS)... [PATH ...]\n"
            "       xlgrep --list-funcs [--by file|sheet] [-e PATTERN | -f FUNCS]... [PATH ...]"
        ),
        description="Search the cells of Excel workbooks (.xlsx/.xlsm/.xltx/.xltm) like grep.",
        epilog=(
            "examples:\n"
            "  xlgrep VLOOKUP reports/            formulas and values containing VLOOKUP\n"
            "  xlgrep -f VLOOKUP,XLOOKUP .        cells calling these functions\n"
            "  xlgrep -p -C1 --header 'Total' .   pretty grid with column headers\n"
            "  xlgrep --in value -F '#N/A' .      cells whose computed value is #N/A\n"
            "  xlgrep --list-funcs .              which functions are used, and how often\n"
            "  xlgrep --ref 'Data!A:D' .          formulas that reference columns A-D of sheet Data\n"
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
    g.add_argument("--ref", action="append", default=[], metavar="RANGE",
                   help="only formulas referencing RANGE (repeatable): 'Data!A:D', \"'Raw Data'!B2\", "
                        "'Data!' for a whole sheet, 'A1:B5' for any sheet; follows defined names. "
                        "Combines with PATTERN/-e/-f as AND; without them positional args are all paths")

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

    g = p.add_argument_group("function summary")
    g.add_argument("--list-funcs", action="store_true",
                   help="count the functions used in formulas instead of printing matches; "
                        "-e/-f then filter function names and all positional args are paths")
    g.add_argument("--by", choices=["file", "sheet"], help="--list-funcs: break counts down per file or sheet")
    g.add_argument("--sort", choices=["calls", "name"], default="calls",
                   help="--list-funcs: order by call count (default) or name")

    g = p.add_argument_group("context")
    g.add_argument("-A", "--after-context", type=int, metavar="N", help="show N cells below each match (same column)")
    g.add_argument("-B", "--before-context", type=int, metavar="N", help="show N cells above each match (same column)")
    g.add_argument("-C", "--context", type=int, metavar="N", help="show N cells above and below each match")
    g.add_argument("--row", action="store_true", help="show the other cells in each matching row")
    g.add_argument("--header", type=int, nargs="?", const=1, metavar="ROW",
                   help="show the column header taken from ROW (default 1)")
    g.add_argument("--col-context", type=int, default=1, metavar="N", help="pretty mode: columns shown on each side (default 1)")

    p.add_argument("-j", "--threads", type=int, metavar="N",
                   help="files to search in parallel (default: CPU count when there are several "
                        "files totalling over 1 MB, otherwise 1)")
    p.add_argument("-V", "--version", action="version", version=f"%(prog)s {__version__}")
    return p


def _use_color(choice: str) -> bool:
    if choice != "auto":
        return choice == "always"
    return sys.stdout.isatty() and "NO_COLOR" not in os.environ and os.environ.get("TERM") != "dumb"


def _list_items(paths: list[str], globs: list[str]) -> list[Path | UnsupportedFile]:
    return list(iter_files(paths or ["."], globs))


def _scope(args: argparse.Namespace, cell_range: CellRange | None) -> Scope:
    return Scope(objects=set(args.objects), sheet_globs=list(args.sheet), no_hidden=args.no_hidden,
                 cell_range=cell_range, raw_formula=args.raw_formula)


def _quiet_pipe() -> None:
    """Output was piped into e.g. `head` and closed; discard the rest silently."""
    devnull = os.open(os.devnull, os.O_WRONLY)
    os.dup2(devnull, sys.stdout.fileno())


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    funcs = [name.strip() for item in args.func for name in item.split(",") if name.strip()]
    if args.threads is not None and args.threads < 1:
        parser.error("-j/--threads must be at least 1")
    if args.list_funcs:
        return list_funcs(parser, args, funcs)
    if args.by or args.sort != "calls":
        parser.error("--by and --sort only apply to --list-funcs")
    if args.regexp or funcs or args.ref:
        patterns, paths = args.regexp, args.args
    elif args.args:
        patterns, paths = args.args[:1], args.args[1:]
    else:
        parser.error("no pattern given (use PATTERN, -e, -f or --ref)")
    if funcs and args.search_in == "value":
        parser.error("-f/--func searches formulas and can't be combined with --in value")
    if args.ref and args.search_in == "value":
        parser.error("--ref searches formulas and can't be combined with --in value")
    try:
        targets = [parse_target(r) for r in args.ref]
        matcher = build_matcher(
            patterns,
            funcs=funcs,
            fixed=args.fixed_strings,
            ignore_case=args.ignore_case,
            smart_case=args.smart_case,
            word=args.word_regexp,
        )
        cell_range = parse_range(args.cell_range) if args.cell_range else None
    except re.error as exc:
        parser.error(f"invalid pattern: {exc}")
    except ValueError as exc:
        parser.error(str(exc))

    context = args.context or 0
    style = Style(_use_color(args.color))
    output = OutputOptions(
        style=style,
        value_mode=args.search_in == "value",
        show_value=args.show_value,
        header_row=args.header,
        before=args.before_context if args.before_context is not None else context,
        after=args.after_context if args.after_context is not None else context,
        row_context=args.row,
        col_context=args.col_context,
        only_matching=args.only_matching,
        with_filename=not args.no_filename,
        max_width=args.max_width,
    )
    mode = "pretty" if args.pretty else "json" if args.json else "csv" if args.csv else "line"
    cfg = SearchConfig(
        scope=_scope(args, cell_range),
        matcher=matcher,
        targets=targets,
        search_in=args.search_in,
        invert=args.invert_match,
        output=output,
        mode=mode,
        with_values=args.search_in == "value" or args.show_value or mode != "line",
        summary_only=args.files_with_matches or args.count or args.quiet,
        limit=1 if (args.files_with_matches or args.quiet) else args.max_count,
    )

    items = _list_items(paths, args.glob)
    files = [i for i in items if not isinstance(i, UnsupportedFile)]
    runner = Runner(search_file, cfg, choose_jobs(args.threads, files))
    any_match = had_error = False
    printed_any = printed_group = False
    try:
        if mode == "csv":
            csv.writer(sys.stdout, lineterminator="\n").writerow(CSV_HEADER)
        results = runner.results(files)
        for item in items:
            if isinstance(item, UnsupportedFile):
                print(f"xlgrep: {item}", file=sys.stderr)
                had_error = True
                continue
            result = next(results)
            for error in result.errors:
                print(f"xlgrep: {error}", file=sys.stderr)
                had_error = True
            if not result.count:
                continue
            any_match = True
            if args.quiet:
                return EXIT_MATCH
            if args.files_with_matches:
                print(style.path(str(item)))
            elif args.count:
                print(f"{style.path(str(item))}:{result.count}")
            elif result.output:
                # Cross-file separators the per-file formatters can't know about.
                if mode == "pretty" and printed_any:
                    sys.stdout.write("\n")
                if mode == "line" and printed_group and result.printed_group:
                    sys.stdout.write(style.sep("--") + "\n")
                sys.stdout.write(result.output)
                printed_any = True
                printed_group = printed_group or result.printed_group
        sys.stdout.flush()
    except BrokenPipeError:
        _quiet_pipe()
        return EXIT_MATCH if any_match else EXIT_NO_MATCH
    except KeyboardInterrupt:
        return 130
    finally:
        runner.close()

    if had_error:
        return EXIT_ERROR
    return EXIT_MATCH if any_match else EXIT_NO_MATCH


_NOT_WITH_LIST_FUNCS = {
    "files_with_matches": "-l", "count": "-c", "quiet": "-q", "only_matching": "-o", "invert_match": "-v",
    "max_count": "-m", "after_context": "-A", "before_context": "-B", "context": "-C", "row": "--row",
    "header": "--header", "show_value": "--show-value", "ref": "--ref",
}


def list_funcs(parser: argparse.ArgumentParser, args: argparse.Namespace, funcs: list[str]) -> int:
    for dest, flag in _NOT_WITH_LIST_FUNCS.items():
        if getattr(args, dest) not in (None, False, []):
            parser.error(f"{flag} can't be combined with --list-funcs")
    if args.search_in == "value":
        parser.error("--list-funcs counts formulas and can't be combined with --in value")
    try:
        pattern = build_matcher(args.regexp, fixed=args.fixed_strings, ignore_case=args.ignore_case,
                                smart_case=args.smart_case, word=args.word_regexp).pattern
        cell_range = parse_range(args.cell_range) if args.cell_range else None
    except re.error as exc:
        parser.error(f"invalid pattern: {exc}")
    except ValueError as exc:
        parser.error(str(exc))

    cfg = FuncConfig(scope=_scope(args, cell_range), pattern=pattern, wanted={f.upper() for f in funcs}, by=args.by)
    items = _list_items(args.args, args.glob)
    files = [i for i in items if not isinstance(i, UnsupportedFile)]
    runner = Runner(count_file, cfg, choose_jobs(args.threads, files))
    totals = FuncCounter(lambda name: True)
    had_error = False
    try:
        results = runner.results(files)
        for item in items:
            if isinstance(item, UnsupportedFile):
                print(f"xlgrep: {item}", file=sys.stderr)
                had_error = True
                continue
            groups, errors = next(results)
            for error in errors:
                print(f"xlgrep: {error}", file=sys.stderr)
                had_error = True
            totals.merge(groups)
    except KeyboardInterrupt:
        return 130
    finally:
        runner.close()

    style = Style(_use_color(args.color))
    mode = "json" if args.json else "csv" if args.csv else "table"
    try:
        found = write_func_stats(totals.rows(args.sort), args.by, mode, style)
        sys.stdout.flush()
    except BrokenPipeError:
        _quiet_pipe()
        return EXIT_MATCH
    if had_error:
        return EXIT_ERROR
    return EXIT_MATCH if found else EXIT_NO_MATCH


def run() -> None:
    sys.exit(main())
