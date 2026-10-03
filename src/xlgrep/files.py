"""Finding workbook files to search."""

from __future__ import annotations

import fnmatch
import os
from collections.abc import Iterator
from pathlib import Path

SUPPORTED_SUFFIXES = {".xlsx", ".xlsm", ".xltx", ".xltm"}
UNSUPPORTED_SUFFIXES = {".xls", ".xlsb", ".ods"}


class UnsupportedFile(Exception):
    pass


def _is_lock_file(name: str) -> bool:
    # Excel creates "~$book.xlsx" next to an open workbook.
    return name.startswith("~$")


def glob_ok(name: str, globs: list[str]) -> bool:
    """rg-style globs: plain globs include, '!glob' excludes. No include globs = include all."""
    includes = [g for g in globs if not g.startswith("!")]
    excludes = [g[1:] for g in globs if g.startswith("!")]
    if includes and not any(fnmatch.fnmatch(name, g) for g in includes):
        return False
    return not any(fnmatch.fnmatch(name, g) for g in excludes)


def iter_files(paths: list[str], globs: list[str]) -> Iterator[Path | UnsupportedFile]:
    """Yield workbook paths in sorted order.

    Explicitly named files are always searched (or reported if unsupported);
    directories are walked recursively, skipping hidden entries and lock files.
    """
    for raw in paths:
        path = Path(raw)
        if path.is_dir():
            yield from _walk(path, globs)
        elif not path.exists():
            yield UnsupportedFile(f"{path}: No such file or directory")
        elif path.suffix.lower() in SUPPORTED_SUFFIXES:
            yield path
        elif path.suffix.lower() in UNSUPPORTED_SUFFIXES:
            yield UnsupportedFile(f"{path}: unsupported format {path.suffix} (only .xlsx/.xlsm/.xltx/.xltm)")
        else:
            yield UnsupportedFile(f"{path}: not an Excel workbook")


def _walk(root: Path, globs: list[str]) -> Iterator[Path]:
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = sorted(d for d in dirnames if not d.startswith("."))
        for name in sorted(filenames):
            if name.startswith(".") or _is_lock_file(name):
                continue
            if Path(name).suffix.lower() not in SUPPORTED_SUFFIXES:
                continue
            if glob_ok(name, globs):
                yield Path(dirpath) / name
