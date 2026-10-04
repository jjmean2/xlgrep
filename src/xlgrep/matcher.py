"""Pattern compilation and match-span calculation."""

from __future__ import annotations

import re
from dataclasses import dataclass

from .text import mask_string_literals

Span = tuple[int, int]


@dataclass
class Matcher:
    pattern: re.Pattern[str] | None  # general patterns, applied to any searched text
    func_pattern: re.Pattern[str] | None  # function calls, applied to formulas only

    @property
    def empty(self) -> bool:
        """No patterns at all (e.g. only --ref was given): everything passes."""
        return self.pattern is None and self.func_pattern is None

    def spans(self, text: str, is_formula: bool) -> list[Span]:
        found: list[Span] = []
        if self.pattern is not None:
            found.extend(m.span() for m in self.pattern.finditer(text) if m.end() > m.start())
        if self.func_pattern is not None and is_formula:
            masked = mask_string_literals(text)
            found.extend(m.span() for m in self.func_pattern.finditer(masked))
        return merge_spans(found)


def build_matcher(
    patterns: list[str],
    *,
    funcs: list[str] | None = None,
    fixed: bool = False,
    ignore_case: bool = False,
    smart_case: bool = False,
    word: bool = False,
) -> Matcher:
    """Compile the patterns. Raises ``re.error`` on an invalid regular expression."""
    pattern = None
    if patterns:
        parts = [re.escape(p) if fixed else p for p in patterns]
        if word:
            parts = [rf"(?<!\w)(?:{p})(?!\w)" for p in parts]
        flags = 0
        if ignore_case or (smart_case and not any(ch.isupper() for ch in "".join(patterns))):
            flags |= re.IGNORECASE
        pattern = re.compile("|".join(f"(?:{p})" for p in parts), flags)

    func_pattern = None
    if funcs:
        names = "|".join(re.escape(name) for name in funcs)
        # Optional prefixes only show up with --raw-formula; they're part of the match then.
        func_pattern = re.compile(
            rf"(?<![\w.])(?:_xl(?:fn|ws)\.)*(?:{names})(?=\s*\()",
            re.IGNORECASE,
        )
    return Matcher(pattern, func_pattern)


def merge_spans(spans: list[Span]) -> list[Span]:
    if len(spans) < 2:
        return spans
    spans.sort()
    merged = [spans[0]]
    for start, end in spans[1:]:
        last_start, last_end = merged[-1]
        if start <= last_end:
            merged[-1] = (last_start, max(last_end, end))
        else:
            merged.append((start, end))
    return merged
