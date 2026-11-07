"""差分の内部表現。各レンダラ(compact / json)はこれだけを読む。

座標の規約: 削除系(-)は旧ファイルの座標、挿入系(+)と変更系(~)は新ファイルの座標。
変更系で旧座標が異なる場合のみ old_addr を持つ。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, NamedTuple


class Val(NamedTuple):
    """セルの値。display は compact 表示用(文字列は引用符付き)、raw は JSON 用。"""

    display: str
    raw: Any


@dataclass
class CellContent:
    """挿入・削除された行(列)に含まれていたセルの内容。"""

    addr: str
    value: Val | None
    link: str | None = None
    comment: str | None = None


@dataclass
class CellChange:
    addr: str
    old_addr: str | None = None
    value: tuple[Val | None, Val | None] | None = None
    style: dict[str, tuple] | None = None
    link: tuple[str | None, str | None] | None = None
    comment: tuple[str | None, str | None] | None = None


@dataclass
class RowEvent:
    op: str  # "insert" | "delete" | "modify"
    old_row: int | None
    new_row: int | None
    cells: list[CellContent] = field(default_factory=list)  # insert / delete
    changes: list[CellChange] = field(default_factory=list)  # modify
    dim: dict[str, tuple] | None = None  # {"height": (old, new), "hidden": (old, new)}


@dataclass
class ColEvent:
    op: str  # "insert" | "delete" | "modify"
    old_col: str | None
    new_col: str | None
    cells: list[CellContent] = field(default_factory=list)
    dim: dict[str, tuple] | None = None


@dataclass
class SheetDiff:
    name: str
    status: str  # "modified" | "added" | "removed" | "renamed"
    old_name: str | None = None
    size: tuple[str, str] | None = None  # 追加・削除シートの使用範囲 (旧, 新)
    rows: list[RowEvent] = field(default_factory=list)
    cols: list[ColEvent] = field(default_factory=list)
    merges_added: list[str] = field(default_factory=list)
    merges_removed: list[str] = field(default_factory=list)
    props: dict[str, tuple] = field(default_factory=dict)
    cond_formats_added: list[str] = field(default_factory=list)
    cond_formats_removed: list[str] = field(default_factory=list)
    validations_added: list[str] = field(default_factory=list)
    validations_removed: list[str] = field(default_factory=list)
    names: dict[str, tuple] = field(default_factory=dict)

    def is_empty(self) -> bool:
        return self.status == "modified" and not (
            self.rows
            or self.cols
            or self.merges_added
            or self.merges_removed
            or self.props
            or self.cond_formats_added
            or self.cond_formats_removed
            or self.validations_added
            or self.validations_removed
            or self.names
        )

    def stats(self) -> dict[str, int]:
        s = {
            "cells": 0,
            "styles": 0,
            "links": 0,
            "comments": 0,
            "rows_inserted": 0,
            "rows_deleted": 0,
            "cols_inserted": 0,
            "cols_deleted": 0,
        }
        for ev in self.rows:
            if ev.op == "insert":
                s["rows_inserted"] += 1
            elif ev.op == "delete":
                s["rows_deleted"] += 1
            for ch in ev.changes:
                s["cells"] += ch.value is not None
                s["styles"] += ch.style is not None
                s["links"] += ch.link is not None
                s["comments"] += ch.comment is not None
        for ev in self.cols:
            if ev.op == "insert":
                s["cols_inserted"] += 1
            elif ev.op == "delete":
                s["cols_deleted"] += 1
        return s


@dataclass
class WorkbookDiff:
    old_path: str
    new_path: str
    sheets: list[SheetDiff] = field(default_factory=list)
    sheet_order: tuple[list[str], list[str]] | None = None
    names: dict[str, tuple] = field(default_factory=dict)

    def is_empty(self) -> bool:
        return (
            all(s.is_empty() for s in self.sheets) and self.sheet_order is None and not self.names
        )
