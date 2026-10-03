import json
import os

import pytest

from xlgrep.cli import main


@pytest.fixture
def run(sample_dir, capsys):
    def _run(*args):
        cwd = os.getcwd()
        os.chdir(sample_dir)
        try:
            code = main(["--color", "never", *args])
        finally:
            os.chdir(cwd)
        out, err = capsys.readouterr()
        return code, out, err

    return _run


def test_default_line_output(run):
    code, out, err = run("VLOOKUP")
    assert code == 0, err
    assert out.splitlines() == [
        "sales.xlsx:Summary!B2:=VLOOKUP(A2,Data!A:D,4,FALSE)",
        "sales.xlsx:Summary!B3:=VLOOKUP(A3,Data!A:D,4,FALSE)",
        'sales.xlsx:Summary!B6:="VLOOKUP is not called here"',
        "sub/budget.xlsx:Secret!A1:VLOOKUP note",
    ]


def test_func_search_ignores_strings_and_case(run):
    code, out, _ = run("-f", "vlookup,XLOOKUP")
    assert out.splitlines() == [
        "sales.xlsx:Summary!B2:=VLOOKUP(A2,Data!A:D,4,FALSE)",
        "sales.xlsx:Summary!B3:=VLOOKUP(A3,Data!A:D,4,FALSE)",
        "sales.xlsx:'Raw Data'!B1:=XLOOKUP(A1,C:C,D:D)",
        "sub/budget.xlsx:Sheet1!B1:=vlookup(A1,Rates,2)",
    ]


def test_raw_formula(run):
    _, out, _ = run("--raw-formula", "-f", "XLOOKUP")
    assert out == "sales.xlsx:'Raw Data'!B1:=_xlfn.XLOOKUP(A1,C:C,D:D)\n"


def test_no_match_exit_code(run):
    code, out, _ = run("NOPE")
    assert (code, out) == (1, "")


def test_search_values_and_show_value(run):
    _, out, _ = run("--in", "value", "-F", "#N/A")
    assert out == "sales.xlsx:'Raw Data'!B1:#N/A\n"
    _, out, _ = run("--show-value", "SUM")
    assert out == "sales.xlsx:Summary!B4:=SUM(B2:B3) → 2180\n"


def test_newlines_are_escaped(run):
    _, out, _ = run("two")
    assert out == "sales.xlsx:'Raw Data'!A2:line one\\nline two\n"


def test_context_and_header(run):
    _, out, _ = run("-C", "1", "--header", "--sheet", "Summary", "SUM")
    assert out.splitlines() == [
        "sales.xlsx-Summary!B3[매출액]-=VLOOKUP(A3,Data!A:D,4,FALSE)",
        "sales.xlsx:Summary!B4[매출액]:=SUM(B2:B3)",
    ]


def test_row_context_groups(run):
    _, out, _ = run("--row", "--sheet", "Summary", "-f", "VLOOKUP")
    assert out.splitlines() == [
        "sales.xlsx-Summary!A2-서울",
        "sales.xlsx:Summary!B2:=VLOOKUP(A2,Data!A:D,4,FALSE)",
        "sales.xlsx-Summary!C2-ok",
        "sales.xlsx-Summary!A3-부산",
        "sales.xlsx:Summary!B3:=VLOOKUP(A3,Data!A:D,4,FALSE)",
    ]


def test_filters(run):
    _, out, _ = run("-l", "--no-hidden", "VLOOKUP")
    assert out == "sales.xlsx\n"
    _, out, _ = run("-c", "-g", "budget*", "VLOOKUP")
    assert out == "sub/budget.xlsx:1\n"
    _, out, _ = run("-c", "-i", "-g", "!budget*", "vlookup")
    assert out == "sales.xlsx:3\n"
    _, out, _ = run("--range", "B3:B4", "-o", "[A-Z]+")
    assert out.splitlines() == [
        "sales.xlsx:Summary!B3:VLOOKUP", "sales.xlsx:Summary!B3:A", "sales.xlsx:Summary!B3:D",
        "sales.xlsx:Summary!B3:A", "sales.xlsx:Summary!B3:D", "sales.xlsx:Summary!B3:FALSE",
        "sales.xlsx:Summary!B4:SUM", "sales.xlsx:Summary!B4:B", "sales.xlsx:Summary!B4:B",
    ]


def test_max_count_and_invert(run):
    _, out, _ = run("-m", "1", "-f", "VLOOKUP")
    assert len(out.splitlines()) == 2  # one per file
    _, out, _ = run("-v", "--sheet", "Sheet1", ".")
    assert out == ""


def test_json(run):
    _, out, _ = run("--json", "SUM")
    rec = json.loads(out)
    assert rec == {
        "file": "sales.xlsx", "sheet": "Summary", "cell": "B4", "row": 4, "col": 2,
        "kind": "formula", "formula": "=SUM(B2:B3)", "value": 2180, "text": "=SUM(B2:B3)",
        "matches": [{"start": 1, "end": 4, "text": "SUM"}],
    }


def test_csv(run):
    _, out, _ = run("--csv", "합계")
    assert out.splitlines() == ["file,sheet,cell,kind,content,value", "sales.xlsx,Summary,A4,value,합계,합계"]


def test_pretty(run):
    _, out, _ = run("-p", "--sheet", "Summary", "-f", "SUM")
    assert out.splitlines() == [
        "sales.xlsx",
        "  Summary",
        "      │ A    │ B           │ C",
        "  ▶ 4 │ 합계 │ =SUM(B2:B3) │",
        "  1 match",
    ]


def test_pretty_multiple_matches_in_one_row(run):
    _, out, _ = run("-p", "--sheet", "Summary", "--range", "2:2", ".")
    assert out.splitlines() == [
        "sales.xlsx",
        "  Summary",
        "      │ A    │ B                             │ C  │ D",
        "  ▶ 2 │ 서울 │ =VLOOKUP(A2,Data!A:D,4,FALSE) │ ok │",
        "  3 matches",
    ]


def test_errors(run, sample_dir):
    (sample_dir / "broken.xlsx").write_text("not a zip")
    (sample_dir / "old.xls").write_bytes(b"BIFF")
    code, out, err = run("VLOOKUP", "missing.xlsx", "old.xls", "broken.xlsx")
    assert code == 2 and out == ""
    assert "missing.xlsx: No such file" in err
    assert "unsupported format .xls" in err
    assert "broken.xlsx: cannot read workbook" in err


def test_usage_errors(run):
    with pytest.raises(SystemExit) as exc:
        run()
    assert exc.value.code == 2
    with pytest.raises(SystemExit):
        run("-f", "SUM", "--in", "value")
