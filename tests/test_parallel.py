"""Parallel runs (-j) must print exactly what a sequential run prints.

Tests that start worker processes are marked slow (skipped by the pre-commit hook).
"""

import os
from pathlib import Path

import pytest

from xlgrep.cli import main
from xlgrep.search import PARALLEL_MIN_BYTES, choose_jobs


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


@pytest.mark.slow
@pytest.mark.parametrize(
    "args",
    [
        ["VLOOKUP"],
        ["-C", "1", "-i", "vlookup"],  # "--" separators between groups across files
        ["-p", "-f", "VLOOKUP"],  # blank line between files
        ["--csv", "VLOOKUP"],  # single header
        ["--json", "-f", "VLOOKUP"],
        ["-c", "VLOOKUP"],
        ["-l", "VLOOKUP"],
        ["--list-funcs"],
        ["--list-funcs", "--by", "file"],
        ["--stats"],
        ["--stats", "--by", "sheet", "--json"],
    ],
)
def test_parallel_matches_sequential(run, args):
    sequential = run("-j", "1", *args)
    parallel = run("-j", "2", *args)
    assert parallel == sequential
    assert sequential[1]


@pytest.mark.slow
def test_parallel_errors_keep_order(run, sample_dir):
    (sample_dir / "broken.xlsx").write_text("not a zip")
    code, out, err = run("-j", "2", "VLOOKUP", "missing.xlsx", "broken.xlsx", "sales.xlsx")
    assert code == 2
    assert err.splitlines() == [
        "xlgrep: missing.xlsx: No such file or directory",
        "xlgrep: broken.xlsx: cannot read workbook (File is not a zip file)",
    ]
    assert out.startswith("sales.xlsx:")


@pytest.mark.slow
def test_quiet_stops_on_first_match(run):
    assert run("-j", "2", "-q", "VLOOKUP")[:2] == (0, "")


def test_choose_jobs(tmp_path: Path):
    small = [tmp_path / "a.xlsx", tmp_path / "b.xlsx"]
    for p in small:
        p.write_bytes(b"x")
    assert choose_jobs(None, small) == 1
    assert choose_jobs(None, small[:1]) == 1
    assert choose_jobs(3, small) == 3
    big = tmp_path / "big.xlsx"
    big.write_bytes(b"x" * PARALLEL_MIN_BYTES)
    assert choose_jobs(None, [big, small[0]]) == min(2, os.cpu_count() or 1)
