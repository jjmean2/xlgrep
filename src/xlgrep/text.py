"""Turning cell contents into searchable/displayable text."""

from __future__ import annotations

import datetime as dt
import re
import unicodedata

# Prefixes Excel writes into the file for newer functions and LAMBDA parameters,
# e.g. "_xlfn.XLOOKUP", "_xlfn._xlws.FILTER", "_xlpm.x". Excel hides them in the UI.
_FUTURE_PREFIX_RE = re.compile(r"_xl(?:fn|ws|pm)\.", re.IGNORECASE)
_STRING_LITERAL_RE = re.compile(r'"(?:[^"]|"")*"')
_ESCAPES = str.maketrans({"\n": "\\n", "\r": "\\r", "\t": "\\t"})


def normalize_formula(formula: str) -> str:
    """Strip future-function prefixes outside string literals."""
    out = []
    pos = 0
    for m in _STRING_LITERAL_RE.finditer(formula):
        out.append(_FUTURE_PREFIX_RE.sub("", formula[pos : m.start()]))
        out.append(m.group(0))
        pos = m.end()
    out.append(_FUTURE_PREFIX_RE.sub("", formula[pos:]))
    return "".join(out)


def mask_string_literals(formula: str) -> str:
    """Blank out the inside of string literals, keeping offsets intact."""
    return _STRING_LITERAL_RE.sub(lambda m: '"' + " " * (len(m.group(0)) - 2) + '"', formula)


def value_text(value: object) -> str:
    """Render a cell value roughly the way it would read in a sheet."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, float):
        if value.is_integer() and abs(value) < 1e15:
            return str(int(value))
        return repr(value)
    if isinstance(value, dt.datetime):
        if value.time() == dt.time(0):
            return value.date().isoformat()
        return value.isoformat(sep=" ")
    if isinstance(value, (dt.date, dt.time)):
        return value.isoformat()
    if isinstance(value, dt.timedelta):
        return str(value)
    return str(value)


def escape(text: str) -> str:
    """Escape control characters so a cell always fits on one line.

    Each escaped character becomes two characters, so match spans must be
    remapped with :func:`escape_spans`.
    """
    return text.translate(_ESCAPES)


def escape_spans(text: str, spans: list[tuple[int, int]]) -> list[tuple[int, int]]:
    """Map spans in ``text`` to spans in ``escape(text)``."""
    if not spans or not any(ch in "\n\r\t" for ch in text):
        return spans
    offsets = [0] * (len(text) + 1)
    shift = 0
    for i, ch in enumerate(text):
        offsets[i] = i + shift
        if ch in "\n\r\t":
            shift += 1
    offsets[len(text)] = len(text) + shift
    return [(offsets[s], offsets[e]) for s, e in spans]


def char_width(ch: str) -> int:
    if unicodedata.combining(ch):
        return 0
    return 2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1


def display_width(text: str) -> int:
    return sum(char_width(ch) for ch in text)
