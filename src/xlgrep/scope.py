"""What a run looks at: which sheets, cells and objects are in scope.

Searching and --list-funcs both walk a workbook the same way; the rules for what
is included (sheet globs, hidden sheets, --range, --objects, workbook-scoped
names) live here so they can't drift apart.
"""

from __future__ import annotations

import fnmatch
import zipfile
from collections.abc import Iterator
from dataclasses import dataclass, field
from pathlib import Path
from xml.etree import ElementTree as ET

from .address import CellRange
from .objects import SheetObject, read_package_objects
from .package import SheetInfo, open_package
from .workbook import Cell, CellReader, Sheet

# What a broken or unreadable workbook raises; reported per file, not fatal.
READ_ERRORS = (zipfile.BadZipFile, KeyError, OSError, ValueError, ET.ParseError)


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
class Part:
    """One sheet's cells and objects in scope; or, with ``sheet_name`` None, the
    workbook-scoped defined names."""

    sheet_name: str | None
    hidden: bool
    info: SheetInfo | None  # the sheet's package entry (None for workbook-scoped names)
    sheet: Sheet | None  # every cell read from the sheet, for showing neighbours
    cells: list[Cell]  # the cells in scope, in reading order
    objects: list[SheetObject]  # the objects in scope


class ScopedWorkbook:
    """An open workbook, walked sheet by sheet within a scope.

    Use as a context manager; the file stays open while ``parts()`` is iterated.
    Defined names are always read, since --ref and --list-funcs need them even
    when names aren't searched.
    """

    def __init__(self, path: Path, scope: Scope, *, raw_formula: bool, formulas_only: bool = False):
        self.scope = scope
        self.formulas_only = formulas_only
        self.zip = zipfile.ZipFile(path)
        try:
            self.package = open_package(self.zip)
            self.objects = read_package_objects(self.zip, self.package, scope.object_kinds | {"name"}, raw_formula)
            self._cells = CellReader(self.zip, self.package, raw_formula) if "cell" in scope.objects else None
        except BaseException:
            self.zip.close()
            raise

    def __enter__(self) -> ScopedWorkbook:  # noqa: PYI034 (typing.Self needs Python 3.11)
        return self

    def __exit__(self, *exc) -> None:
        self.zip.close()

    @property
    def sheet_names(self) -> list[str]:
        return [s.name for s in self.objects.sheets]

    @property
    def defined_names(self) -> dict[tuple[str | None, str], str]:
        """(scope sheet or None, NAME upper-cased) -> the name's formula."""
        names = {(o.sheet, o.ref.upper()): o.text for o in self.objects.workbook_names}
        for objects in self.objects.by_sheet.values():
            names.update({(o.sheet, o.ref.upper()): o.text for o in objects if o.object == "name"})
        return names

    def parts(self) -> Iterator[Part]:
        """Sheets in workbook order, then the workbook-scoped names."""
        scope = self.scope
        kinds = scope.object_kinds
        rng = scope.cell_range
        for info in self.objects.sheets:
            if not scope.sheet_ok(info.name, info.hidden):
                continue
            sheet = cells = None
            if self._cells is not None and info.part is not None:
                sheet = self._cells.read(info, formulas_only=self.formulas_only)
                cells = [c for c in sheet.sorted_cells() if rng is None or rng.contains(c.row, c.col)]
            objects = [
                o for o in self.objects.by_sheet[info.name] if o.object in kinds and (rng is None or o.in_range(rng))
            ]
            yield Part(info.name, info.hidden, info, sheet, cells or [], objects)
        # Workbook-scoped names belong to no sheet or range, so --sheet/--range exclude them.
        if "name" in kinds and not scope.sheet_globs and rng is None:
            yield Part(None, False, None, None, [], list(self.objects.workbook_names))
