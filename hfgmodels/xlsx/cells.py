"""Cell reference helpers."""

from __future__ import annotations

import re


def col_letter(n: int) -> str:
    """1 -> A, 27 -> AA."""
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def col_index(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def ref(row: int, col: int, abs_row: bool = False, abs_col: bool = False) -> str:
    return f"{'$' if abs_col else ''}{col_letter(col)}{'$' if abs_row else ''}{row}"


_REF = re.compile(r"^\$?([A-Z]+)\$?(\d+)$")


def parse_ref(a1: str) -> tuple[int, int]:
    m = _REF.match(a1.upper())
    if not m:
        raise ValueError(f"not a cell reference: {a1}")
    return int(m.group(2)), col_index(m.group(1))


def quote_sheet(name: str) -> str:
    if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.]*", name):
        return name
    return "'" + name.replace("'", "''") + "'"
