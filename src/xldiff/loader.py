"""ワークブックを比較用の正規化スナップショットに変換する。

openpyxl のオブジェクトをそのまま比較すると、保存し直しただけで変わる
スタイル ID や型の揺れ(日付↔数値)がノイズになる。ここで値・書式・
付随情報を比較可能な素朴なデータに落とし込む。
"""

from __future__ import annotations

import datetime as dt
import warnings
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.cell.cell import Cell, MergedCell
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import column_index_from_string
from openpyxl.utils.datetime import to_excel


class LoadError(Exception):
    """入力ファイルを読み込めないときに送出する。"""


# ---------------------------------------------------------------------------
# データ構造
# ---------------------------------------------------------------------------


@dataclass(slots=True)
class CellData:
    key: tuple | None  # 値の比較キー。None は空セル
    display: str | None  # 表示用の文字列(文字列は引用符付き)
    raw: object  # JSON 出力用の生の値(日付は ISO 文字列)
    formula: str | None  # 数式セルなら数式テキスト(先頭の = 付き)
    style: dict  # 平坦化した書式(同一書式は同一オブジェクトを共有)
    link: str | None
    comment: str | None


@dataclass
class SheetData:
    name: str
    cells: dict[tuple[int, int], CellData]
    max_row: int
    max_col: int
    merged: list[tuple[int, int, int, int]]  # (min_row, min_col, max_row, max_col)
    row_dims: dict[int, tuple]  # row -> (height, hidden)
    col_dims: dict[int, tuple]  # col -> (width, hidden)
    default_style: dict
    props: dict
    cond_formats: list[str]
    validations: list[str]
    names: dict[str, str] = field(default_factory=dict)


@dataclass
class WorkbookData:
    path: Path
    sheets: list[SheetData]
    names: dict[str, str]

    def sheet(self, name: str) -> SheetData | None:
        for s in self.sheets:
            if s.name == name:
                return s
        return None


# ---------------------------------------------------------------------------
# 値の正規化
# ---------------------------------------------------------------------------


def _num_key(x: float) -> float:
    # 日付→シリアル値変換の丸め誤差を吸収する
    return round(float(x), 9)


def _format_number(x) -> str:
    if isinstance(x, float) and x.is_integer() and abs(x) < 1e15:
        return str(int(x))
    return repr(x)


def _format_datetime(v) -> str:
    if isinstance(v, dt.datetime):
        if v.time() == dt.time(0, 0):
            return v.date().isoformat()
        return v.isoformat(sep=" ")
    return v.isoformat()


def _quote(s: str) -> str:
    s = s.replace("\\", "\\\\").replace('"', '\\"')
    s = s.replace("\r\n", "\\n").replace("\n", "\\n").replace("\r", "\\n").replace("\t", "\\t")
    return f'"{s}"'


def normalize_value(cell: Cell) -> tuple[tuple | None, str | None, object, str | None]:
    """(比較キー, 表示文字列, 生の値, 数式テキスト) を返す。"""
    v = cell.value
    if v is None:
        return None, None, None, None

    # 数式(配列数式・データテーブルを含む)
    if cell.data_type == "f":
        if isinstance(v, str):
            f = v if v.startswith("=") else "=" + v
            return ("f", f), f, f, f
        text = getattr(v, "text", None)  # ArrayFormula
        if isinstance(text, str):
            f = text if text.startswith("=") else "=" + text
            return ("f", "{" + f + "}"), "{" + f + "}", "{" + f + "}", f
        s = f"<{type(v).__name__} {getattr(v, 'ref', '')}>"  # DataTableFormula など
        return ("f", s), s, s, None

    if isinstance(v, bool):
        return ("b", v), "TRUE" if v else "FALSE", v, None
    if isinstance(v, (int, float)):
        return ("n", _num_key(v)), _format_number(v), v, None
    if isinstance(v, (dt.datetime, dt.date, dt.time)):
        # 表示形式を「標準」に変えると日付が数値として読まれるため、シリアル値で比較する
        shown = _format_datetime(v)
        return ("n", _num_key(to_excel(v))), shown, shown, None
    if isinstance(v, dt.timedelta):
        return ("n", _num_key(v.total_seconds() / 86400)), str(v), str(v), None
    if isinstance(v, str):
        if v == "":
            return None, None, None, None
        if cell.data_type == "e":
            return ("e", v), v, v, None
        return ("s", v), _quote(v), v, None
    s = str(v)
    return ("s", s), _quote(s), s, None


# ---------------------------------------------------------------------------
# 書式の平坦化
# ---------------------------------------------------------------------------


def color_str(c) -> str | None:
    if c is None:
        return None
    try:
        t = c.type
        if t == "rgb":
            v = c.rgb
            return v if isinstance(v, str) else None
        if t == "theme":
            tint = c.tint or 0
            return f"theme{c.theme}" + (f"{tint:+.2f}" if tint else "")
        if t == "indexed":
            return f"indexed{c.indexed}"
        if t == "auto":
            return "auto"
    except Exception:  # 不正な色指定は文字列化して比較する
        return str(c)
    return None


def _num_or_none(x):
    if x is None:
        return None
    x = float(x)
    return int(x) if x.is_integer() else round(x, 2)


def flatten_style(cell: Cell) -> dict:
    d: dict = {}
    f = cell.font
    if f is not None:
        d["font.name"] = f.name
        d["font.size"] = _num_or_none(f.sz)
        d["font.bold"] = bool(f.b)
        d["font.italic"] = bool(f.i)
        d["font.underline"] = f.u or None
        d["font.strike"] = bool(f.strike)
        d["font.vertAlign"] = f.vertAlign or None
        d["font.color"] = color_str(f.color)
    fill = cell.fill
    if fill is not None:
        ptype = getattr(fill, "patternType", None)
        if ptype is not None:
            d["fill.pattern"] = ptype
            d["fill.fgColor"] = color_str(fill.fgColor)
            if ptype != "solid":
                d["fill.bgColor"] = color_str(fill.bgColor)
        elif getattr(fill, "type", None) == "gradient" or fill.tagname == "gradientFill":
            stops = ",".join(f"{s.position}:{color_str(s.color)}" for s in (fill.stop or []))
            d["fill.gradient"] = f"{fill.type}/{fill.degree}/{stops}"
        else:
            d["fill.pattern"] = None
    b = cell.border
    if b is not None:
        for side_name in ("left", "right", "top", "bottom", "diagonal"):
            side = getattr(b, side_name, None)
            style = getattr(side, "style", None) if side is not None else None
            d[f"border.{side_name}"] = (
                f"{style}:{color_str(side.color) or 'auto'}" if style else None
            )
    a = cell.alignment
    if a is not None:
        d["align.horizontal"] = a.horizontal
        d["align.vertical"] = a.vertical
        d["align.wrap"] = bool(a.wrap_text)
        d["align.shrink"] = bool(a.shrink_to_fit)
        d["align.indent"] = _num_or_none(a.indent) or 0
        d["align.rotation"] = _num_or_none(a.text_rotation) or 0
    d["numFmt"] = cell.number_format
    p = cell.protection
    if p is not None:
        d["protect.locked"] = bool(p.locked) if p.locked is not None else True
        d["protect.hidden"] = bool(p.hidden)
    return d


class _StyleCache:
    """同じ StyleArray のセルで平坦化結果を共有し、メモリと時間を節約する。"""

    def __init__(self) -> None:
        self._cache: dict[tuple, dict] = {}

    def get(self, cell: Cell) -> dict:
        k = tuple(cell._style) if cell._style is not None else ()  # None は既定書式
        d = self._cache.get(k)
        if d is None:
            d = flatten_style(cell)
            self._cache[k] = d
        return d


# ---------------------------------------------------------------------------
# 読み込み
# ---------------------------------------------------------------------------


def _check_path(path: Path) -> None:
    if not path.exists():
        raise LoadError(f"ファイルが見つかりません: {path}")
    if path.is_dir():
        raise LoadError(f"ディレクトリが指定されました: {path}")
    if path.name.startswith("~$"):
        raise LoadError(f"Excel のロックファイル(~$...)です。元のファイルを指定してください: {path}")
    suffix = path.suffix.lower()
    if suffix == ".xls":
        raise LoadError(
            f".xls(Excel 97-2003 形式)には対応していません。.xlsx で保存し直してください: {path}"
        )
    if not zipfile.is_zipfile(path):
        raise LoadError(
            f"xlsx として読めません(パスワード保護または別形式の可能性があります): {path}"
        )


def _comment_text(cell: Cell) -> str | None:
    c = cell.comment
    if c is None:
        return None
    return (c.text or "").strip() or None


def _link_text(cell: Cell) -> str | None:
    h = cell.hyperlink
    if h is None:
        return None
    target = h.target
    if target:
        return target + (f"#{h.location}" if h.location else "")
    if h.location:
        return f"#{h.location}"
    return None


def _cond_formats(ws) -> list[str]:
    out = []
    try:
        for cf in ws.conditional_formatting:
            for rule in cf.rules:
                parts = [str(cf.sqref), rule.type or ""]
                if rule.operator:
                    parts.append(rule.operator)
                if rule.formula:
                    parts.append(" ".join(rule.formula))
                if rule.text:
                    parts.append(_quote(rule.text))
                dxf = rule.dxf
                if dxf is not None:
                    if dxf.font is not None:
                        parts.append(f"font={color_str(dxf.font.color)}/b={bool(dxf.font.b)}")
                    if dxf.fill is not None:
                        parts.append(
                            f"fill={color_str(getattr(dxf.fill, 'bgColor', None)) or color_str(getattr(dxf.fill, 'fgColor', None))}"
                        )
                out.append(" ".join(p for p in parts if p))
    except Exception:
        pass
    return sorted(out)


def _validations(ws) -> list[str]:
    out = []
    try:
        for dv in ws.data_validations.dataValidation:
            parts = [str(dv.sqref), dv.type or "any"]
            if dv.operator:
                parts.append(dv.operator)
            if dv.formula1:
                parts.append(str(dv.formula1))
            if dv.formula2:
                parts.append(str(dv.formula2))
            out.append(" ".join(parts))
    except Exception:
        pass
    return sorted(out)


def _defined_names(container) -> dict[str, str]:
    out = {}
    try:
        for name, dn in container.defined_names.items():
            out[name] = str(dn.attr_text)
    except Exception:
        pass
    return out


def _load_sheet(ws, styles: _StyleCache) -> SheetData:
    default_cell = Cell(ws, row=1, column=1)  # ws._cells に登録されない既定セル
    default_style = flatten_style(default_cell)

    cells: dict[tuple[int, int], CellData] = {}
    max_row = max_col = 0
    for (r, c), cell in ws._cells.items():
        if isinstance(cell, MergedCell):
            continue
        key, display, raw, formula = normalize_value(cell)
        style = styles.get(cell)
        link = _link_text(cell)
        comment = _comment_text(cell)
        if key is None and link is None and comment is None and style == default_style:
            continue  # 何も情報を持たない空セルは無視(再保存での増減を吸収)
        cells[(r, c)] = CellData(key, display, raw, formula, style, link, comment)
        max_row = max(max_row, r)
        max_col = max(max_col, c)

    merged = []
    for rng in ws.merged_cells.ranges:
        merged.append((rng.min_row, rng.min_col, rng.max_row, rng.max_col))
        max_row = max(max_row, rng.max_row)
        max_col = max(max_col, rng.max_col)

    row_dims: dict[int, tuple] = {}
    for r, rd in ws.row_dimensions.items():
        ht = _num_or_none(rd.ht) if rd.customHeight or rd.ht is not None else None
        hidden = bool(rd.hidden)
        if ht is None and not hidden:
            continue
        row_dims[r] = (ht, hidden)
        max_row = max(max_row, r)

    col_dims: dict[int, tuple] = {}
    for letter, cd in ws.column_dimensions.items():
        width = _num_or_none(cd.width) if cd.customWidth else None
        hidden = bool(cd.hidden)
        if width is None and not hidden:
            continue
        lo = cd.min or column_index_from_string(letter)
        hi = cd.max or lo
        if hi - lo > 1000:  # 「残り全列」指定などは先頭だけ見る
            hi = lo
        for c in range(lo, hi + 1):
            col_dims[c] = (width, hidden)
            max_col = max(max_col, c)

    props = {
        "state": ws.sheet_state,
        "freeze": ws.freeze_panes,
        "autoFilter": ws.auto_filter.ref if ws.auto_filter is not None else None,
        "tabColor": color_str(ws.sheet_properties.tabColor) if ws.sheet_properties else None,
    }

    return SheetData(
        name=ws.title,
        cells=cells,
        max_row=max_row,
        max_col=max_col,
        merged=sorted(merged),
        row_dims=row_dims,
        col_dims=col_dims,
        default_style=default_style,
        props=props,
        cond_formats=_cond_formats(ws),
        validations=_validations(ws),
        names=_defined_names(ws),
    )


def load(path: str | Path) -> WorkbookData:
    path = Path(path)
    _check_path(path)
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")  # 「Data Validation extension is not supported」等を抑止
            wb = load_workbook(path, data_only=False, rich_text=False)
    except Exception as e:  # noqa: BLE001
        raise LoadError(f"読み込みに失敗しました: {path}: {e}") from e

    styles = _StyleCache()
    sheets = []
    for ws in wb.worksheets:  # グラフシート(chartsheet)は対象外
        sheets.append(_load_sheet(ws, styles))
    return WorkbookData(path=path, sheets=sheets, names=_defined_names(wb))


def col_letter(c: int) -> str:
    return get_column_letter(c)


def addr(r: int, c: int) -> str:
    return f"{get_column_letter(c)}{r}"
