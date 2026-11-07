"""WorkbookDiff を文字列に変換する。

compact: 1 行 1 変更の省トークン形式(既定)。TTY ではカラー表示。
stat   : シートごとの件数サマリ(git diff --stat 相当)。
json   : 機械処理用。値が None の項目は省略する。
"""

from __future__ import annotations

import json
import unicodedata
from dataclasses import fields, is_dataclass

from .model import CellChange, CellContent, ColEvent, RowEvent, SheetDiff, Val, WorkbookDiff

ARROW = "→"


class Style:
    def __init__(self, color: bool) -> None:
        self.color = color

    def _w(self, code: str, s: str) -> str:
        return f"\x1b[{code}m{s}\x1b[0m" if self.color else s

    def red(self, s: str) -> str:
        return self._w("31", s)

    def green(self, s: str) -> str:
        return self._w("32", s)

    def yellow(self, s: str) -> str:
        return self._w("33", s)

    def cyan(self, s: str) -> str:
        return self._w("36", s)

    def bold(self, s: str) -> str:
        return self._w("1", s)


# ---------------------------------------------------------------------------
# compact
# ---------------------------------------------------------------------------


def _q(s: str) -> str:
    s = s.replace("\\", "\\\\").replace('"', '\\"').replace("\n", "\\n").replace("\r", "")
    return f'"{s}"'


class Compact:
    def __init__(self, color: bool = False, max_len: int = 0) -> None:
        self.st = Style(color)
        self.max_len = max_len

    # 値の表示 ------------------------------------------------------------
    def _cut(self, s: str) -> str:
        if self.max_len and len(s) > self.max_len:
            if s.startswith('"') and s.endswith('"'):
                return s[: self.max_len] + '…"'
            return s[: self.max_len] + "…"
        return s

    def val(self, v: Val | None) -> str:
        return "(empty)" if v is None else self._cut(v.display)

    def text(self, v: str | None) -> str:
        return "(none)" if v is None else self._cut(_q(v))

    @staticmethod
    def attr(v) -> str:
        if v is None:
            return "none"
        if isinstance(v, bool):
            return "true" if v else "false"
        return str(v)

    def change(self, old: str, new: str) -> str:
        return f"{self.st.red(old)} {ARROW} {self.st.green(new)}"

    def content(self, cc: CellContent) -> str:
        s = f"{cc.addr}: {self.val(cc.value)}"
        if cc.link is not None and (cc.value is None or cc.value.raw != cc.link):
            s += f" <link {self.text(cc.link)}>"
        if cc.comment is not None:
            s += f" <comment {self.text(cc.comment)}>"
        return s

    # 行 -----------------------------------------------------------------
    def cell_lines(self, ch: CellChange) -> list[str]:
        st = self.st
        where = ch.addr + (f" (was {ch.old_addr})" if ch.old_addr else "")
        tag = st.yellow("~")
        out = []
        if ch.value is not None:
            out.append(f"{tag} {where}: {self.change(self.val(ch.value[0]), self.val(ch.value[1]))}")
        if ch.link is not None:
            out.append(f"{tag} {where} link: {self.change(self.text(ch.link[0]), self.text(ch.link[1]))}")
        if ch.comment is not None:
            out.append(
                f"{tag} {where} comment: {self.change(self.text(ch.comment[0]), self.text(ch.comment[1]))}"
            )
        if ch.style is not None:
            parts = [f"{k}: {self.change(self.attr(a), self.attr(b))}" for k, (a, b) in ch.style.items()]
            out.append(f"{tag} {where} style: " + ", ".join(parts))
        return out

    def dim_line(self, label: str, dim: dict) -> str:
        parts = [f"{k}: {self.change(self.attr(a), self.attr(b))}" for k, (a, b) in dim.items()]
        return f"{self.st.yellow('~')} {label} " + ", ".join(parts)

    def row_lines(self, ev: RowEvent) -> list[str]:
        st = self.st
        if ev.op == "insert":
            body = " | ".join(self.content(c) for c in ev.cells)
            return [st.green(f"+ row {ev.new_row}:" + (f" {body}" if body else " (empty)"))]
        if ev.op == "delete":
            body = " | ".join(self.content(c) for c in ev.cells)
            return [st.red(f"- row {ev.old_row}:" + (f" {body}" if body else " (empty)"))]
        out = []
        for ch in ev.changes:
            out.extend(self.cell_lines(ch))
        if ev.dim:
            label = f"row {ev.new_row}" + (f" (was {ev.old_row})" if ev.old_row != ev.new_row else "")
            out.append(self.dim_line(label, ev.dim))
        return out

    def col_lines(self, ev: ColEvent) -> list[str]:
        st = self.st
        if ev.op == "insert":
            body = " | ".join(self.content(c) for c in ev.cells)
            return [st.green(f"+ col {ev.new_col}:" + (f" {body}" if body else " (empty)"))]
        if ev.op == "delete":
            body = " | ".join(self.content(c) for c in ev.cells)
            return [st.red(f"- col {ev.old_col}:" + (f" {body}" if body else " (empty)"))]
        label = f"col {ev.new_col}" + (f" (was {ev.old_col})" if ev.old_col else "")
        return [self.dim_line(label, ev.dim or {})]

    def sheet_lines(self, sd: SheetDiff) -> list[str]:
        st = self.st
        if sd.status == "added":
            head = f"@@ {sd.name} (added sheet: {sd.size[1]}) @@"
        elif sd.status == "removed":
            return [st.cyan(f"@@ {sd.name} (removed sheet: {sd.size[0]}) @@")]
        elif sd.status == "renamed":
            head = f"@@ {sd.name} (renamed from {sd.old_name}) @@"
        else:
            head = f"@@ {sd.name} @@"
        out = [st.cyan(head)]
        for k, (a, b) in sd.props.items():
            out.append(f"{st.yellow('~')} sheet {k}: {self.change(self.attr(a), self.attr(b))}")
        for ev in sd.cols:
            out.extend(self.col_lines(ev))
        for ev in sd.rows:
            out.extend(self.row_lines(ev))
        for m in sd.merges_removed:
            out.append(st.red(f"- merge {m}"))
        for m in sd.merges_added:
            out.append(st.green(f"+ merge {m}"))
        for x in sd.cond_formats_removed:
            out.append(st.red(f"- condfmt {x}"))
        for x in sd.cond_formats_added:
            out.append(st.green(f"+ condfmt {x}"))
        for x in sd.validations_removed:
            out.append(st.red(f"- validation {x}"))
        for x in sd.validations_added:
            out.append(st.green(f"+ validation {x}"))
        for k, (a, b) in sd.names.items():
            out.append(f"{st.yellow('~')} name {k}: {self.change(self.attr(a), self.attr(b))}")
        return out

    def render(self, wd: WorkbookDiff) -> str:
        if wd.is_empty():
            return ""
        st = self.st
        out = [st.bold(f"--- {wd.old_path}"), st.bold(f"+++ {wd.new_path}")]
        if wd.sheet_order:
            a, b = wd.sheet_order
            out.append(f"{st.yellow('~')} sheet order: {self.change(str(a), str(b))}")
        for k, (a, b) in wd.names.items():
            out.append(f"{st.yellow('~')} name {k}: {self.change(self.attr(a), self.attr(b))}")
        for sd in wd.sheets:
            if not sd.is_empty():
                out.extend(self.sheet_lines(sd))
        return "\n".join(out) + "\n"


# ---------------------------------------------------------------------------
# stat
# ---------------------------------------------------------------------------


def _width(s: str) -> int:
    """端末上の表示幅(全角は 2)。"""
    return sum(2 if unicodedata.east_asian_width(ch) in "WF" else 1 for ch in s)


def render_stat(wd: WorkbookDiff, color: bool = False) -> str:
    st = Style(color)
    lines = []
    changed = 0
    for sd in wd.sheets:
        if sd.is_empty():
            continue
        changed += 1
        if sd.status in ("added", "removed"):
            lines.append((sd.name, f"{sd.status} ({(sd.size[1] if sd.status == 'added' else sd.size[0])})"))
            continue
        s = sd.stats()
        parts = []
        if s["cells"]:
            parts.append(st.yellow(f"~{s['cells']} cells"))
        if s["rows_inserted"]:
            parts.append(st.green(f"+{s['rows_inserted']} rows"))
        if s["rows_deleted"]:
            parts.append(st.red(f"-{s['rows_deleted']} rows"))
        if s["cols_inserted"]:
            parts.append(st.green(f"+{s['cols_inserted']} cols"))
        if s["cols_deleted"]:
            parts.append(st.red(f"-{s['cols_deleted']} cols"))
        if s["styles"]:
            parts.append(f"~{s['styles']} styles")
        if s["links"]:
            parts.append(f"~{s['links']} links")
        if s["comments"]:
            parts.append(f"~{s['comments']} comments")
        other = len(sd.merges_added) + len(sd.merges_removed) + len(sd.props) + len(sd.names)
        other += len(sd.cond_formats_added) + len(sd.cond_formats_removed)
        other += len(sd.validations_added) + len(sd.validations_removed)
        other += sum(1 for ev in sd.rows if ev.dim) + sum(1 for ev in sd.cols if ev.op == "modify")
        if other:
            parts.append(f"~{other} other")
        name = sd.name + (f" (renamed from {sd.old_name})" if sd.old_name else "")
        lines.append((name, ", ".join(parts)))
    if wd.sheet_order:
        lines.append(("(workbook)", "sheet order changed"))
    if wd.names:
        lines.append(("(workbook)", f"~{len(wd.names)} names"))
    if not lines:
        return ""
    w = max(_width(n) for n, _ in lines)
    out = [f" {n}{' ' * (w - _width(n))} | {d}" for n, d in lines]
    out.append(f" {changed} sheet(s) changed")
    return "\n".join(out) + "\n"


# ---------------------------------------------------------------------------
# json
# ---------------------------------------------------------------------------

_PAIR_FIELDS = {"value", "link", "comment"}


def _pair(t):
    return {"old": _to_obj(t[0]), "new": _to_obj(t[1])}


def _to_obj(x):
    if isinstance(x, Val):
        return x.raw
    if is_dataclass(x):
        d = {}
        for f in fields(x):
            v = getattr(x, f.name)
            if v is None or v == [] or v == {}:
                continue
            if f.name in _PAIR_FIELDS and isinstance(v, tuple) and not isinstance(v, Val):
                d[f.name] = _pair(v)
            elif f.name in ("style", "dim", "props", "names") and isinstance(v, dict):
                d[f.name] = {k: _pair(p) for k, p in v.items()}
            elif f.name == "size" and isinstance(v, tuple):
                d[f.name] = _pair(v)
            elif f.name == "sheet_order" and isinstance(v, tuple):
                d[f.name] = _pair(v)
            else:
                d[f.name] = _to_obj(v)
        return d
    if isinstance(x, list):
        return [_to_obj(v) for v in x]
    if isinstance(x, tuple):
        return [_to_obj(v) for v in x]
    if isinstance(x, dict):
        return {k: _to_obj(v) for k, v in x.items()}
    return x


def render_json(wd: WorkbookDiff, indent: int | None = None) -> str:
    obj = _to_obj(wd)
    obj["sheets"] = [_to_obj(s) for s in wd.sheets if not s.is_empty()]
    obj["identical"] = wd.is_empty()
    return json.dumps(obj, ensure_ascii=False, indent=indent, default=str) + "\n"
