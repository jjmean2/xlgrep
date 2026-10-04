# xlgrep

grep for the cells of Excel workbooks. Search formulas and values across many
`.xlsx` / `.xlsm` / `.xltx` / `.xltm` files and get `file:Sheet!Cell:content` lines back.

```
$ xlgrep -f VLOOKUP reports/
reports/sales.xlsx:Summary!B2:=VLOOKUP(A2,Data!A:D,4,FALSE)
reports/sales.xlsx:'Raw Data'!F3:=IFERROR(VLOOKUP(E3,Map,2,0),"")
```

## Install

```
uv tool install .        # or: pipx install .
```

For development:

```
uv sync
uv run xlgrep --help
uv run pytest
```

## Usage

```
xlgrep [OPTIONS] PATTERN [PATH ...]
xlgrep [OPTIONS] (-e PATTERN | -f FUNCS)... [PATH ...]
```

Directories are searched recursively; with no path, the current directory is searched.
Exit status is 0 if a cell matched, 1 if none did, 2 on errors.

| Task | Command |
|---|---|
| Cells calling a function (ignores string literals, case-insensitive) | `xlgrep -f VLOOKUP,XLOOKUP .` |
| Regex over formulas and values | `xlgrep 'Data!\$?A' .` |
| Literal string, case-insensitive | `xlgrep -Fi 'total (krw)' .` |
| Computed values (cached results) only | `xlgrep --in value -F '#N/A' .` |
| Show cached results next to formulas | `xlgrep --show-value -f SUMIFS .` |
| Context: 2 cells above/below, column header | `xlgrep -C2 --header -f SUMIFS .` |
| Rest of the matching row | `xlgrep --row 'TODO' .` |
| Pretty grid grouped by file and sheet | `xlgrep -p -C1 --header -f SUMIFS .` |
| Machine-readable | `xlgrep --json ...`, `xlgrep --csv ...` |
| Formulas referencing a range (follows defined names) | `xlgrep --ref 'Data!A:D' .` |
| Anything referencing a sheet, before deleting it | `xlgrep --ref 'Data!' .` |
| VLOOKUPs that read from a range | `xlgrep -f VLOOKUP --ref 'Data!A:D' .` |
| Which functions are used, and how often | `xlgrep --list-funcs .` |
| Size and migration-relevant metrics per workbook | `xlgrep --stats .`, `--stats -p .`, `--stats --by sheet --csv .` |
| Function usage per file / sheet | `xlgrep --list-funcs --by file -f VLOOKUP,XLOOKUP .` |
| Matching files / counts | `xlgrep -l ...`, `xlgrep -c ...` |
| Parallelism (default: auto) | `-j 8`, `-j 1` for sequential |
| Limit scope | `-g '*.xlsm'`, `-g '!*backup*'`, `--sheet 'Data*'`, `--range B2:F100`, `--no-hidden` |
| Only some places | `--objects cells`, `--objects names,cf,dv` |

### What gets searched

By default (`--in auto`) formula cells are matched on their formula text and other cells
on their value. `--in formula` restricts the search to formula cells; `--in value` matches
what the sheet displays, using the result Excel stored for formulas when the file was last
saved (files written by tools that don't calculate have no stored results).

Besides cells, xlgrep searches formulas outside cells — defined names, conditional formatting
rules and data validations — and the text of cell notes. These print as `Summary!B2:B50#cf:=...`,
`TaxRate#name:=...`, `Summary!C3#note:...`. Restrict with `--objects`.

`--ref RANGE` selects formulas whose references overlap RANGE, highlighting them. Unqualified
references resolve to the formula's own sheet, 3D references expand across sheets, relative
references in conditional formats and validations move across the range they apply to, and
defined names are followed (also through other names). Combined with a pattern or `-f`, both
must match. References built from strings (`INDIRECT`), external workbooks and table references
(`Table1[Col]`) are not resolved.

`--stats` prints, per workbook: size, sheets, non-empty cells, formulas, **unique formulas**
(formulas compared in relative R1C1 form, so a formula filled down a column counts once),
formulas calling volatile functions, array formulas, cells showing errors, VBA and external
links. `-p` shows every metric (custom functions, LAMBDAs, conditional formats, validations,
names, notes, tables, pivot tables, charts, data connections) as cards; `--json`/`--csv`
include them all; `--by sheet` gives a row per sheet.

`--list-funcs` prints a table of functions with their call count, the number of formulas
using them and the number of files. Custom functions are marked `lambda` (a LAMBDA defined
name in the same workbook) or `custom` (VBA / add-in). In this mode `-e`/`-f` filter
function names, and all positional arguments are paths.

Newer functions are stored as `_xlfn.XLOOKUP`, `_xlfn._xlws.FILTER` etc.; xlgrep strips
these prefixes so formulas read as they do in Excel. Use `--raw-formula` to keep them.

Legacy `.xls` files are not supported.

See [DESIGN.md](DESIGN.md) for the design and roadmap, and [ARCHITECTURE.md](ARCHITECTURE.md) for a guide to the code.
