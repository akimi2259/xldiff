"""数式中のセル参照を書き換えるユーティリティ。

行を挿入すると Excel はそれ以降を参照する数式を自動で書き換える
(=A5 → =A6)。これをそのまま比較すると大量の偽の差分になるため、
- 行の対応付け用には「自セルからの相対位置」で正規化した数式を、
- 差分判定用には旧ファイルの参照を新ファイルの座標へ写像した数式を
作って比較する。

openpyxl の Tokenizer は大量のセルに使うと遅いため、文字列リテラルを
読み飛ばしつつ正規表現で参照だけを拾う。
"""

from __future__ import annotations

import re
from typing import Callable

from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import column_index_from_string

# 文字列リテラル、または(シート名付きの)セル参照・列範囲・行範囲
_SCAN = re.compile(
    r"""
    (?P<str>"(?:[^"]|"")*")
  | (?<![A-Za-z0-9_.$!\]\[\\'])
    (?P<ref>
      (?:(?P<sheet>'(?:[^']|'')+'|[^\s'"!()\[\],;:+\-*/^&=<>%{}@\#$]+)!)?
      (?:
         (?P<cell>\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)
       | (?P<cols>\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3})
       | (?P<rows>\$?\d+:\$?\d+)
      )
    )
    (?![A-Za-z0-9_(\[.!$:])
    """,
    re.X,
)
_EP = re.compile(r"(\$?)([A-Za-z]{1,3})?(\$?)(\d+)?")
_RE_PLAIN_SHEET = re.compile(r"^[A-Za-z_][A-Za-z0-9_.]*$")

# endpoint: (col, col_abs, row, row_abs)。列全体・行全体の参照では片方が None
Endpoint = tuple[int | None, bool, int | None, bool]
# fn(sheet, endpoint, role) -> 書き換え後の endpoint。None は写像不能
#   role: "single" | "start" | "end"
MapFn = Callable[[str | None, Endpoint, str], Endpoint | None]


class _Unmappable(Exception):
    pass


def _unquote(sheet: str) -> str:
    if sheet.startswith("'") and sheet.endswith("'"):
        return sheet[1:-1].replace("''", "'")
    return sheet


def quote_sheet(name: str) -> str:
    if _RE_PLAIN_SHEET.match(name) or (name and all(ord(ch) > 127 or ch.isalnum() or ch == "_" for ch in name)):
        return name
    return "'" + name.replace("'", "''") + "'"


def _parse_ep(s: str) -> Endpoint:
    m = _EP.fullmatch(s)
    assert m is not None
    ca, col, ra, row = m.groups()
    if col is None:  # 行全体 ($1 など)
        return (None, False, int(row), ca == "$")
    return (
        column_index_from_string(col.upper()),
        ca == "$",
        int(row) if row is not None else None,
        ra == "$",
    )


def _fmt_ep(ep: Endpoint) -> str:
    col, col_abs, row, row_abs = ep
    s = ""
    if col is not None:
        s += ("$" if col_abs else "") + get_column_letter(col)
    if row is not None:
        s += ("$" if row_abs else "") + str(row)
    return s


def _endpoints(m: re.Match) -> list[Endpoint]:
    body = m.group("cell") or m.group("cols") or m.group("rows")
    return [_parse_ep(p) for p in body.split(":")]


def transform(
    formula: str, fn: MapFn, sheet_fn: Callable[[str], str] | None = None
) -> str | None:
    """数式中のすべてのセル参照を fn で書き換える。

    写像できない参照(削除された行を単独で指す等)が1つでもあれば None。
    名前定義・テーブル参照・外部ブック参照はそのまま残す。
    """

    def repl(m: re.Match) -> str:
        if m.group("str") is not None:
            return m.group("str")
        sheet = m.group("sheet")
        sname = _unquote(sheet) if sheet is not None else None
        eps = _endpoints(m)
        if len(eps) == 1:
            new = [fn(sname, eps[0], "single")]
        else:
            new = [fn(sname, eps[0], "start"), fn(sname, eps[1], "end")]
        if any(e is None for e in new):
            raise _Unmappable
        out = ":".join(_fmt_ep(e) for e in new)  # type: ignore[arg-type]
        if sname is not None:
            out = quote_sheet(sheet_fn(sname) if sheet_fn else sname) + "!" + out
        return out

    try:
        return _SCAN.sub(repl, formula)
    except _Unmappable:
        return None


def _identity(sheet, ep, role):
    return ep


def canonical(formula: str, sheet_fn: Callable[[str], str] | None = None) -> str:
    """表記揺れ(列記号の大小・シート名の引用符)を吸収した数式文字列。"""
    out = transform(formula, _identity, sheet_fn)
    return out if out is not None else formula


def relative(formula: str, row: int, col: int, sheet_fn: Callable[[str], str] | None = None) -> str:
    """自セル (row, col) からの相対位置で参照を表した数式(行の対応付け用)。

    =A5 を B5 に置いたものと =A6 を B6 に置いたものが同じ文字列になる。
    絶対参照の番号は捨てて "R$" / "C$" にする。=B2/$B$20 のような式は、合計行より
    上に行を挿入すると全行で $B$21 に変わるため、番号を残すと対応付けが崩れる。
    絶対参照の実際の変更は、対応付け後のセル比較(参照の写像)で検出される。
    """

    def repl(m: re.Match) -> str:
        if m.group("str") is not None:
            return m.group("str")
        parts = []
        for c, ca, r, ra in _endpoints(m):
            s = ""
            if r is not None:
                s += "R$" if ra else f"R[{r - row}]"
            if c is not None:
                s += "C$" if ca else f"C[{c - col}]"
            parts.append(s)
        out = ":".join(parts)
        sheet = m.group("sheet")
        if sheet is not None:
            sname = _unquote(sheet)
            out = (sheet_fn(sname) if sheet_fn else sname) + "!" + out
        return out

    return _SCAN.sub(repl, formula)
