import pytest

from xlgrep.address import CellRange, col_index, col_letter, parse_range, quote_sheet
from xlgrep.matcher import build_matcher
from xlgrep.output import _truncate
from xlgrep.text import display_width, escape, escape_spans, mask_string_literals, normalize_formula, value_text


@pytest.mark.parametrize("col,letters", [(1, "A"), (26, "Z"), (27, "AA"), (702, "ZZ"), (703, "AAA")])
def test_col_letters_roundtrip(col, letters):
    assert col_letter(col) == letters
    assert col_index(letters) == col


@pytest.mark.parametrize(
    "name,quoted",
    [("Summary", "Summary"), ("요약", "요약"), ("Raw Data", "'Raw Data'"), ("It's", "'It''s'"),
     ("2024", "'2024'"), ("A1", "'A1'"), ("Q1-Q2", "'Q1-Q2'")],
)
def test_quote_sheet(name, quoted):
    assert quote_sheet(name) == quoted


def test_parse_range():
    assert parse_range("A1:F100") == CellRange(1, 1, 100, 6)
    assert parse_range("$B$2") == CellRange(2, 2, 2, 2)
    assert parse_range("B:D") == CellRange(None, 2, None, 4)
    assert parse_range("10:3") == CellRange(3, None, 10, None)
    assert parse_range("B:D").contains(999, 3)
    with pytest.raises(ValueError):
        parse_range("A1:B2:C3")


def test_normalize_formula_strips_prefixes_outside_strings():
    assert normalize_formula('=_xlfn.XLOOKUP(A1,"_xlfn.x",B:B)') == '=XLOOKUP(A1,"_xlfn.x",B:B)'
    assert normalize_formula("=_xlfn._xlws.FILTER(A:A,A:A>0)") == "=FILTER(A:A,A:A>0)"
    assert normalize_formula("=_xlfn.LAMBDA(_xlpm.x,_xlpm.x+1)") == "=LAMBDA(x,x+1)"


def test_mask_string_literals_keeps_offsets():
    f = '=IF(A1="a ""q"" b",SUM(1),"x")'
    masked = mask_string_literals(f)
    assert len(masked) == len(f)
    assert "SUM(1)" in masked and "a" not in masked[7:19]


def test_value_text():
    import datetime as dt

    assert value_text(3.0) == "3"
    assert value_text(3.25) == "3.25"
    assert value_text(True) == "TRUE"
    assert value_text(dt.datetime(2024, 1, 2)) == "2024-01-02"
    assert value_text(None) == ""


def test_escape_spans():
    text = "a\nVLOOKUP"
    assert escape(text) == "a\\nVLOOKUP"
    assert escape_spans(text, [(2, 9)]) == [(3, 10)]


def test_func_matcher():
    m = build_matcher([], funcs=["VLOOKUP", "XLOOKUP"])
    assert m.spans("=VLOOKUP(A1,B:C,2)", is_formula=True) == [(1, 8)]
    assert m.spans("=vlookup (A1,B:C,2)", is_formula=True) == [(1, 8)]
    assert m.spans('="VLOOKUP("', is_formula=True) == []
    assert m.spans("=MYVLOOKUP(A1)", is_formula=True) == []
    assert m.spans("=VLOOKUP(A1)", is_formula=False) == []
    assert m.spans("=_xlfn.XLOOKUP(A1)", is_formula=True) == [(1, 14)]


def test_pattern_matcher_options():
    assert build_matcher(["a.c"], fixed=True).spans("abc a.c", False) == [(4, 7)]
    assert build_matcher(["sum"], smart_case=True).spans("SUM", False) == [(0, 3)]
    assert build_matcher(["Sum"], smart_case=True).spans("SUM", False) == []
    assert build_matcher(["SUM"], word=True).spans("SUMIF SUM(", False) == [(6, 9)]
    assert build_matcher(["a", "b"]).spans("ab xb", False) == [(0, 2), (4, 5)]


def test_truncate_wide_chars():
    assert display_width("가나다") == 6
    text, spans = _truncate("가나다라마", [(1, 4)], 5)
    assert text == "가나…" and spans == [(1, 2)]
