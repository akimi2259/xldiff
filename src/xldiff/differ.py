"""2つのワークブックのスナップショットを比較し、WorkbookDiff を組み立てる。"""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field

from . import formula as fml
from .align import AxisMap, align
from .loader import CellData, SheetData, WorkbookData, addr, col_letter
from .model import CellChange, CellContent, ColEvent, RowEvent, SheetDiff, Val, WorkbookDiff


@dataclass
class Options:
    style: bool = True  # 書式の差分を出すか
    sheets: list[str] = field(default_factory=list)  # 対象シートの絞り込み(空なら全シート)


# ---------------------------------------------------------------------------
# シートの対応付け
# ---------------------------------------------------------------------------


def _value_set(s: SheetData) -> set:
    return {c.key for c in s.cells.values() if c.key is not None}


def _jaccard(a: set, b: set) -> float:
    if not a and not b:
        return 1.0
    return len(a & b) / len(a | b)


def match_sheets(old: WorkbookData, new: WorkbookData) -> list[tuple[SheetData | None, SheetData | None]]:
    """同名シートを対応付け、残りは内容の類似度で「名前変更」として対応付ける。"""
    new_by_name = {s.name: s for s in new.sheets}
    pairs: dict[str, str] = {}
    for s in old.sheets:
        if s.name in new_by_name:
            pairs[s.name] = s.name
    rem = [s for s in old.sheets if s.name not in pairs]
    add = [s for s in new.sheets if s.name not in pairs.values()]
    cands = []
    for i, o in enumerate(rem):
        ov = _value_set(o)
        for j, n in enumerate(add):
            cands.append((_jaccard(ov, _value_set(n)), i, j))
    used_o, used_n = set(), set()
    for sim, i, j in sorted(cands, key=lambda t: (-t[0], t[1], t[2])):
        if sim < 0.5 or i in used_o or j in used_n:
            continue
        used_o.add(i)
        used_n.add(j)
        pairs[rem[i].name] = add[j].name

    result: list[tuple[SheetData | None, SheetData | None]] = []
    old_by_name = {s.name: s for s in old.sheets}
    inv = {v: k for k, v in pairs.items()}
    # 新ファイルのシート順で並べ、削除されたシートは旧ファイルでの位置付近に差し込む
    for n in new.sheets:
        if n.name in inv:
            result.append((old_by_name[inv[n.name]], n))
        else:
            result.append((None, n))
    for idx, o in enumerate(old.sheets):
        if o.name not in pairs:
            result.insert(min(idx, len(result)), (o, None))
    return result


# ---------------------------------------------------------------------------
# 行・列の対応付け
# ---------------------------------------------------------------------------


@dataclass
class SheetAlignment:
    col_pairs: list  # [(old_col|None, new_col|None)] 1 始まり
    row_pairs: list
    cols: AxisMap
    rows: AxisMap


def _align_key(cell: CellData | None, r: int, c: int, sheet_fn) -> object:
    if cell is None:
        return None
    if cell.formula is not None and cell.key is not None and cell.key[0] == "f":
        return ("f", fml.relative(cell.formula, r, c, sheet_fn))
    return cell.key


def _rstrip_none(t: list) -> tuple:
    k = len(t)
    while k and t[k - 1] is None:
        k -= 1
    return tuple(t[:k])


def align_sheet(o: SheetData, n: SheetData, sheet_fn_old) -> SheetAlignment:
    # 数式は自セルからの相対表記にしておく(行挿入で =A5→=A6 になっても一致させる)
    okeys = {(r, c): _align_key(cell, r, c, sheet_fn_old) for (r, c), cell in o.cells.items()}
    nkeys = {(r, c): _align_key(cell, r, c, None) for (r, c), cell in n.cells.items()}

    # --- 列: 値の集合の類似度で対応付ける(行の挿入があっても崩れない) ---
    ocols: list[list] = [[None] * o.max_row for _ in range(o.max_col)]
    ncols: list[list] = [[None] * n.max_row for _ in range(n.max_col)]
    for (r, c), k in okeys.items():
        ocols[c - 1][r - 1] = k
    for (r, c), k in nkeys.items():
        ncols[c - 1][r - 1] = k
    oc_keys = [_rstrip_none(col) for col in ocols]
    nc_keys = [_rstrip_none(col) for col in ncols]
    oc_sets = [{k for k in col if k is not None} for col in ocols]
    nc_sets = [{k for k in col if k is not None} for col in ncols]
    col_pairs0 = align(oc_keys, nc_keys, lambda i, j: _jaccard(oc_sets[i], nc_sets[j]), 0.3)
    col_pairs = [(None if i is None else i + 1, None if j is None else j + 1) for i, j in col_pairs0]
    matched_cols = [(i, j) for i, j in col_pairs if i is not None and j is not None]

    # --- 行: 対応付いた列の値だけでキーを作る ---
    def row_tuple(keys, max_row, side):
        rows = [[None] * len(matched_cols) for _ in range(max_row)]
        idx = {pair[side]: k for k, pair in enumerate(matched_cols)}
        for (r, c), k in keys.items():
            pos = idx.get(c)
            if pos is not None:
                rows[r - 1][pos] = k
        return [tuple(x) for x in rows]

    orows = row_tuple(okeys, o.max_row, 0)
    nrows = row_tuple(nkeys, n.max_row, 1)

    def row_sim(i: int, j: int) -> float:
        a, b = orows[i], nrows[j]
        if a == b:
            return 1.0
        nz = eq = 0
        for x, y in zip(a, b):
            if x is None and y is None:
                continue
            nz += 1
            if x == y:
                eq += 1
        return eq / nz if nz else 1.0

    row_pairs0 = align(orows, nrows, row_sim, 0.5)
    row_pairs = [(None if i is None else i + 1, None if j is None else j + 1) for i, j in row_pairs0]
    return SheetAlignment(
        col_pairs=col_pairs,
        row_pairs=row_pairs,
        cols=AxisMap(col_pairs, o.max_col),
        rows=AxisMap(row_pairs, o.max_row),
    )


# ---------------------------------------------------------------------------
# セル比較
# ---------------------------------------------------------------------------


class _Ctx:
    def __init__(self, sheet_map: dict[str, str], aligns: dict[str, SheetAlignment]):
        self.sheet_map = sheet_map  # 旧シート名 → 新シート名
        self._sheet_map_ci = {k.lower(): v for k, v in sheet_map.items()}
        self.aligns = {k.lower(): v for k, v in aligns.items()}  # 旧シート名(小文字) → 対応付け

    def rename(self, old_sheet: str) -> str:
        return self._sheet_map_ci.get(old_sheet.lower(), old_sheet)

    def formulas_equal(self, of: str, nf: str, old_sheet: str) -> bool:
        def fn(sheet, ep, role):
            al = self.aligns.get((sheet or old_sheet).lower())
            if al is None:
                return ep
            c, ca, r, ra = ep
            nc = al.cols.map(c, role) if c is not None else None
            nr = al.rows.map(r, role) if r is not None else None
            if (c is not None and nc is None) or (r is not None and nr is None):
                return None
            return (nc, ca, nr, ra)

        mapped = fml.transform(of, fn, self.rename)
        if mapped is None:
            return False
        if mapped == nf:
            return True
        return fml.canonical(mapped) == fml.canonical(nf)


def _val(cell: CellData | None) -> Val | None:
    if cell is None or cell.key is None:
        return None
    return Val(cell.display or "", cell.raw)


def _style_diff(os: dict, ns: dict) -> dict[str, tuple] | None:
    if os is ns or os == ns:
        return None
    out = {}
    for k in sorted(set(os) | set(ns)):
        a, b = os.get(k), ns.get(k)
        if a != b:
            out[k] = (a, b)
    return out or None


def compare_cell(
    oc: CellData | None,
    nc: CellData | None,
    o: SheetData,
    n: SheetData,
    ctx: _Ctx,
    opts: Options,
) -> dict | None:
    out: dict = {}
    ok = oc.key if oc else None
    nk = nc.key if nc else None
    if ok != nk:
        same = False
        if (
            ok is not None
            and nk is not None
            and ok[0] == "f"
            and nk[0] == "f"
            and oc.formula
            and nc.formula
            and ok[1].startswith("{") == nk[1].startswith("{")
        ):
            same = ctx.formulas_equal(oc.formula, nc.formula, o.name)
        if not same:
            out["value"] = (_val(oc), _val(nc))
    if opts.style:
        sd = _style_diff(oc.style if oc else o.default_style, nc.style if nc else n.default_style)
        if sd:
            out["style"] = sd
    ol, nl = (oc.link if oc else None), (nc.link if nc else None)
    if ol != nl:
        out["link"] = (ol, nl)
    om, nm = (oc.comment if oc else None), (nc.comment if nc else None)
    if om != nm:
        out["comment"] = (om, nm)
    return out or None


def _has_content(cell: CellData) -> bool:
    return cell.key is not None or cell.link is not None or cell.comment is not None


def _content(cell: CellData, a: str) -> CellContent | None:
    if not _has_content(cell):
        return None
    return CellContent(addr=a, value=_val(cell), link=cell.link, comment=cell.comment)


def _dim_diff(a: tuple | None, b: tuple | None, names: tuple[str, str]) -> dict | None:
    a = a or (None, False)
    b = b or (None, False)
    out = {}
    for k, x, y in zip(names, a, b):
        if x != y:
            out[k] = (x, y)
    return out or None


def _range_str(r1, c1, r2, c2) -> str:
    a, b = addr(r1, c1), addr(r2, c2)
    return a if a == b else f"{a}:{b}"


def _dict_diff(a: dict, b: dict) -> dict[str, tuple]:
    out = {}
    for k in sorted(set(a) | set(b)):
        if a.get(k) != b.get(k):
            out[k] = (a.get(k), b.get(k))
    return out


def _by_row(cells: dict) -> dict[int, list[int]]:
    d: dict[int, list[int]] = defaultdict(list)
    for r, c in cells:
        d[r].append(c)
    return d


def diff_sheet(o: SheetData, n: SheetData, al: SheetAlignment, ctx: _Ctx, opts: Options) -> SheetDiff:
    sd = SheetDiff(name=n.name, status="modified" if o.name == n.name else "renamed")
    if o.name != n.name:
        sd.old_name = o.name

    col_fwd = al.cols.fwd
    col_inv = {j: i for i, j in col_fwd.items()}
    inserted_cols = {j for i, j in al.col_pairs if i is None}
    orow_cells = _by_row(o.cells)
    nrow_cells = _by_row(n.cells)
    # 中身(値・リンク・コメント)を持つ最終行。これより下の空行の増減は表示しない
    o_last = max((r for (r, _), c in o.cells.items() if _has_content(c)), default=0)
    n_last = max((r for (r, _), c in n.cells.items() if _has_content(c)), default=0)

    # --- 行ごとの差分(git の hunk と同じ順序で並べる) ---
    for oi, nj in al.row_pairs:
        if oi is None:
            cells = [
                cc
                for c in sorted(nrow_cells.get(nj, []))
                if (cc := _content(n.cells[(nj, c)], addr(nj, c)))
            ]
            if cells or nj <= n_last:
                sd.rows.append(RowEvent("insert", None, nj, cells=cells))
            continue
        if nj is None:
            cells = [
                cc
                for c in sorted(orow_cells.get(oi, []))
                if (cc := _content(o.cells[(oi, c)], addr(oi, c)))
            ]
            if cells or oi <= o_last:
                sd.rows.append(RowEvent("delete", oi, None, cells=cells))
            continue
        changes: list[CellChange] = []
        new_cols = set(nrow_cells.get(nj, []))
        new_cols.update(col_fwd[c] for c in orow_cells.get(oi, []) if c in col_fwd)
        for nc_ in sorted(new_cols):
            if nc_ in inserted_cols:
                continue
            oc_ = col_inv.get(nc_)
            if oc_ is None:
                continue
            res = compare_cell(o.cells.get((oi, oc_)), n.cells.get((nj, nc_)), o, n, ctx, opts)
            if res:
                a_new, a_old = addr(nj, nc_), addr(oi, oc_)
                changes.append(CellChange(addr=a_new, old_addr=a_old if a_old != a_new else None, **res))
        dim = _dim_diff(o.row_dims.get(oi), n.row_dims.get(nj), ("height", "hidden")) if opts.style else None
        if changes or dim:
            sd.rows.append(RowEvent("modify", oi, nj, changes=changes, dim=dim))

    # --- 列の挿入・削除・幅変更 ---
    paired_old_rows = {i for i, j in al.row_pairs if i is not None and j is not None}
    paired_new_rows = {j for i, j in al.row_pairs if i is not None and j is not None}
    for oc_, nc_ in al.col_pairs:
        if oc_ is None:
            cells = [
                cc
                for (r, c), cell in sorted(n.cells.items())
                if c == nc_ and r in paired_new_rows and (cc := _content(cell, addr(r, c)))
            ]
            sd.cols.append(ColEvent("insert", None, col_letter(nc_), cells=cells))
        elif nc_ is None:
            cells = [
                cc
                for (r, c), cell in sorted(o.cells.items())
                if c == oc_ and r in paired_old_rows and (cc := _content(cell, addr(r, c)))
            ]
            sd.cols.append(ColEvent("delete", col_letter(oc_), None, cells=cells))
        elif opts.style:
            dim = _dim_diff(o.col_dims.get(oc_), n.col_dims.get(nc_), ("width", "hidden"))
            if dim:
                ol, nl = col_letter(oc_), col_letter(nc_)
                sd.cols.append(ColEvent("modify", ol if ol != nl else None, nl, dim=dim))

    # --- 結合セル(旧の範囲を新座標へ写像して比較) ---
    new_merges = set(n.merged)
    mapped_old = {}
    for r1, c1, r2, c2 in o.merged:
        m = (
            al.rows.map(r1, "start"),
            al.cols.map(c1, "start"),
            al.rows.map(r2, "end"),
            al.cols.map(c2, "end"),
        )
        mapped_old[(r1, c1, r2, c2)] = m if None not in m else None
    mapped_set = {m for m in mapped_old.values() if m is not None}
    sd.merges_removed = [_range_str(*k) for k, m in mapped_old.items() if m is None or m not in new_merges]
    sd.merges_added = [_range_str(*m) for m in n.merged if m not in mapped_set]

    # --- シート単位の設定 ---
    sd.props = _dict_diff(o.props, n.props)
    if opts.style:
        oset, nset = set(o.cond_formats), set(n.cond_formats)
        sd.cond_formats_removed = sorted(oset - nset)
        sd.cond_formats_added = sorted(nset - oset)
    oset, nset = set(o.validations), set(n.validations)
    sd.validations_removed = sorted(oset - nset)
    sd.validations_added = sorted(nset - oset)
    sd.names = _dict_diff(o.names, n.names)
    return sd


def _sheet_size(s: SheetData) -> str:
    if not s.cells:
        return "empty"
    return f"A1:{addr(s.max_row, s.max_col)} ({len(s.cells)} cells)"


def _added_sheet(n: SheetData) -> SheetDiff:
    sd = SheetDiff(name=n.name, status="added", size=(None, _sheet_size(n)))
    for r, cols in sorted(_by_row(n.cells).items()):
        cells = [cc for c in sorted(cols) if (cc := _content(n.cells[(r, c)], addr(r, c)))]
        if cells:
            sd.rows.append(RowEvent("insert", None, r, cells=cells))
    return sd


def diff_workbooks(old: WorkbookData, new: WorkbookData, opts: Options | None = None) -> WorkbookDiff:
    opts = opts or Options()
    wd = WorkbookDiff(old_path=str(old.path), new_path=str(new.path))
    pairs = match_sheets(old, new)
    sheet_map = {o.name: n.name for o, n in pairs if o is not None and n is not None}
    ctx_tmp = _Ctx(sheet_map, {})

    aligns: dict[str, SheetAlignment] = {}
    for o, n in pairs:
        if o is not None and n is not None:
            aligns[o.name] = align_sheet(o, n, ctx_tmp.rename)
    ctx = _Ctx(sheet_map, aligns)

    wanted = set(opts.sheets)
    for o, n in pairs:
        if wanted and not ({o.name if o else None, n.name if n else None} & wanted):
            continue
        if o is not None and n is not None:
            wd.sheets.append(diff_sheet(o, n, aligns[o.name], ctx, opts))
        elif n is not None:
            wd.sheets.append(_added_sheet(n))
        elif o is not None:
            wd.sheets.append(SheetDiff(name=o.name, status="removed", size=(_sheet_size(o), None)))

    # シートの並び順の変化(追加・削除を除いた相対順序で判定)
    common_old = [s.name for s in old.sheets if s.name in sheet_map]
    common_new_by_old = [sheet_map[x] for x in common_old]
    common_new = [s.name for s in new.sheets if s.name in sheet_map.values()]
    if common_new_by_old != common_new and not wanted:
        wd.sheet_order = ([s.name for s in old.sheets], [s.name for s in new.sheets])

    if not wanted:
        wd.names = _dict_diff(old.names, new.names)
    return wd
