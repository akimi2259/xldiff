from __future__ import annotations

import datetime as dt
import json
from copy import copy
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Font, PatternFill

from xldiff.cli import main
from xldiff.differ import diff_workbooks
from xldiff.loader import load
from xldiff.render import Compact, render_json

def diff(a: Path, b: Path, **kw) -> str:
    return Compact().render(diff_workbooks(load(a), load(b)))


def body(a: Path, b: Path) -> list[str]:
    """ヘッダ(---/+++)を除いた差分行。"""
    return [l for l in diff(a, b).splitlines() if not l.startswith(("---", "+++"))]


def make(path: Path, rows: list[list], title: str = "S") -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = title
    for r in rows:
        ws.append(r)
    wb.save(path)
    return path


# ---------------------------------------------------------------------------
# 不変条件
# ---------------------------------------------------------------------------


def test_same_file_has_no_diff(sample):
    assert diff(sample, sample) == ""


def test_resaved_copy_has_no_diff(tmp_path, sample):
    # openpyxl での保存し直しは描画パーツやスタイル表を書き換えるが、差分は出ないこと
    p = tmp_path / "resave.xlsx"
    load_workbook(sample).save(p)
    assert diff(sample, p) == ""


# ---------------------------------------------------------------------------
# 架空サンプル(conftest.py)を使った基本ケース
# ---------------------------------------------------------------------------


def _edit_sample(src: Path, tmp_path, fn, name="new.xlsx") -> Path:
    wb = load_workbook(src)
    fn(wb, wb.worksheets[0])
    p = tmp_path / name
    wb.save(p)
    return p


def test_cell_edit(tmp_path, sample):
    p = _edit_sample(sample, tmp_path, lambda wb, ws: ws.__setitem__("A3", "第2回"))
    assert body(sample, p) == ["@@ シート1 @@", '~ A3: "第二回" → "第2回"']


def test_hyperlink_change(tmp_path, sample):
    def f(wb, ws):
        ws["C4"].hyperlink = "https://example.com/x"

    lines = body(sample, _edit_sample(sample, tmp_path, f))
    assert len(lines) == 2 and lines[1].startswith("~ C4 link: ") and lines[1].endswith('"https://example.com/x"')


def test_comment_add_and_edit(tmp_path, sample):
    a = _edit_sample(sample, tmp_path, lambda wb, ws: setattr(ws["B5"], "comment", Comment("old", "me")), "a.xlsx")
    b = _edit_sample(sample, tmp_path, lambda wb, ws: setattr(ws["B5"], "comment", Comment("new", "me")), "b.xlsx")
    assert body(sample, a)[1] == '~ B5 comment: (none) → "old"'
    assert body(a, b)[1] == '~ B5 comment: "old" → "new"'


def test_style_changes(tmp_path, sample):
    def f(wb, ws):
        f = copy(ws["A2"].font)
        f.b = True
        ws["A2"].font = f
        ws["A3"].fill = PatternFill("solid", fgColor="FFFF00")
        ws["A4"].number_format = "0.00"

    lines = body(sample, _edit_sample(sample, tmp_path, f))
    assert lines[1] == "~ A2 style: font.bold: false → true"
    assert lines[2].startswith("~ A3 style: fill.fgColor: none → 00FFFF00")
    assert lines[3] == "~ A4 style: numFmt: General → 0.00"


def test_no_style_option(tmp_path, sample):
    p = _edit_sample(sample, tmp_path, lambda wb, ws: setattr(ws["A2"], "font", Font(bold=True)))
    assert main([str(sample), str(p), "--exit-code"]) == 1
    assert main([str(sample), str(p), "--no-style", "--exit-code"]) == 0


def test_merge_added(tmp_path, sample):
    lines = body(sample, _edit_sample(sample, tmp_path, lambda wb, ws: ws.merge_cells("A1:C1")))
    assert lines == ["@@ シート1 @@", "+ merge A1:C1"]


def test_sheet_added_and_renamed(tmp_path, sample):
    def f(wb, ws):
        ws.title = "講義一覧"
        wb.create_sheet("追加")["A1"] = 1

    lines = body(sample, _edit_sample(sample, tmp_path, f))
    assert lines == [
        "@@ 講義一覧 (renamed from シート1) @@",
        "@@ 追加 (added sheet: A1:A1 (1 cells)) @@",
        "+ row 1: A1: 1",
    ]


def test_sheet_removed(tmp_path, sample):
    a = _edit_sample(sample, tmp_path, lambda wb, ws: wb.create_sheet("消す")["A1"].__setattr__("value", "x"), "a.xlsx")
    lines = body(a, sample)
    assert lines == ["@@ 消す (removed sheet: A1:A1 (1 cells)) @@"]


# ---------------------------------------------------------------------------
# 行・列の対応付け
# ---------------------------------------------------------------------------

DATA = [["ID", "名前", "点数"]] + [[i, f"name{i}", i * 10] for i in range(1, 11)]


def test_row_insert_in_middle(tmp_path):
    a = make(tmp_path / "a.xlsx", DATA)
    rows = DATA[:5] + [[99, "new", 0]] + DATA[5:]
    b = make(tmp_path / "b.xlsx", rows)
    assert body(a, b) == ["@@ S @@", '+ row 6: A6: 99 | B6: "new" | C6: 0']


def test_row_delete(tmp_path):
    a = make(tmp_path / "a.xlsx", DATA)
    b = make(tmp_path / "b.xlsx", DATA[:3] + DATA[4:])
    assert body(a, b) == ["@@ S @@", '- row 4: A4: 3 | B4: "name3" | C4: 30']


def test_row_insert_next_to_cell_edit(tmp_path):
    a = make(tmp_path / "a.xlsx", DATA)
    rows = [list(r) for r in DATA]
    rows[5][2] = 999  # 元の 6 行目 (ID=5) の点数
    rows.insert(5, [77, "ins", 1])
    b = make(tmp_path / "b.xlsx", rows)
    assert body(a, b) == [
        "@@ S @@",
        '+ row 6: A6: 77 | B6: "ins" | C6: 1',
        "~ C7 (was C6): 50 → 999",
    ]


def test_column_insert(tmp_path):
    a = make(tmp_path / "a.xlsx", DATA)
    b = make(tmp_path / "b.xlsx", [[r[0], "x" if i else "新列", *r[1:]] for i, r in enumerate(DATA)])
    lines = body(a, b)
    assert lines[0] == "@@ S @@"
    assert lines[1].startswith('+ col B: B1: "新列" | B2: "x"')
    assert len(lines) == 2


def test_column_delete(tmp_path):
    a = make(tmp_path / "a.xlsx", DATA)
    b = make(tmp_path / "b.xlsx", [[r[0], r[2]] for r in DATA])
    lines = body(a, b)
    assert lines[1].startswith('- col B: B1: "名前" | B2: "name1"') and len(lines) == 2


def test_large_sheet_with_duplicate_and_blank_rows(tmp_path):
    # 200 行超 + 同じ内容の行や空行が多い表でも、挿入1行・変更1セルだけが出ること
    rows = []
    for i in range(400):
        if i % 7 == 0:
            rows.append([None, None, None])
        elif i % 5 == 0:
            rows.append(["dup", "dup", "dup"])
        else:
            rows.append([i, f"v{i}", i % 3])
    a = make(tmp_path / "a.xlsx", rows)
    rows2 = [list(r) for r in rows]
    rows2[302][1] = "changed"  # 元の 303 行目 [302, "v302", 2]
    rows2.insert(150, ["inserted", 1, 2])
    b = make(tmp_path / "b.xlsx", rows2)
    lines = body(a, b)
    assert lines == [
        "@@ S @@",
        '+ row 151: A151: "inserted" | B151: 1 | C151: 2',
        '~ B304 (was B303): "v302" → "changed"',
    ]


def test_formulas_below_inserted_row_are_not_reported(tmp_path):
    # 行挿入で Excel が =A5 → =A6 と書き換えても差分にしない
    def build(path, insert: bool):
        wb = Workbook()
        ws = wb.active
        ws.title = "S"
        rows = [[i, i * 2] for i in range(1, 9)]
        if insert:
            rows.insert(3, [100, 200])
        for r, (x, y) in enumerate(rows, start=1):
            ws.cell(r, 1, x)
            ws.cell(r, 2, y)
            ws.cell(r, 3, f"=A{r}*B{r}+$A$1")
        n = len(rows)
        ws.cell(n + 1, 3, f"=SUM(C1:C{n})")
        ws.cell(n + 2, 3, f"=S!A{n}+A1")
        wb.save(path)
        return path

    a = build(tmp_path / "a.xlsx", False)
    b = build(tmp_path / "b.xlsx", True)
    assert body(a, b) == ["@@ S @@", "+ row 4: A4: 100 | B4: 200 | C4: =A4*B4+$A$1"]


def test_real_formula_change_is_reported(tmp_path):
    a = make(tmp_path / "a.xlsx", [[1, 2, "=A1+B1"]])
    b = make(tmp_path / "b.xlsx", [[1, 2, "=A1*B1"]])
    assert body(a, b)[1] == "~ C1: =A1+B1 → =A1*B1"


def test_date_to_general_is_style_only(tmp_path):
    a = make(tmp_path / "a.xlsx", [[dt.datetime(2021, 10, 6)]])
    wb = load_workbook(a)
    wb.active["A1"].number_format = "General"
    b = tmp_path / "b.xlsx"
    wb.save(b)
    lines = body(a, b)
    assert len(lines) == 2 and lines[1].startswith("~ A1 style: numFmt:")


def test_row_height_and_column_width_follow_alignment(tmp_path):
    def build(path, insert):
        wb = Workbook()
        ws = wb.active
        ws.title = "S"
        rows = DATA[:3] + ([[50, "x", 1]] if insert else []) + DATA[3:]
        for r in rows:
            ws.append(r)
        ws.row_dimensions[3].height = 30  # 挿入位置より上
        ws.row_dimensions[6 if insert else 5].height = 40  # 挿入でずれる行
        ws.column_dimensions["B"].width = 20
        wb.save(path)
        return path

    a, b = build(tmp_path / "a.xlsx", False), build(tmp_path / "b.xlsx", True)
    assert body(a, b) == ["@@ S @@", '+ row 4: A4: 50 | B4: "x" | C4: 1']


# ---------------------------------------------------------------------------
# 出力形式・CLI
# ---------------------------------------------------------------------------


def test_json_output(tmp_path, capsys):
    a = make(tmp_path / "a.xlsx", DATA)
    rows = [list(r) for r in DATA]
    rows[1][1] = "改名"
    rows.insert(3, [50, "x", dt.datetime(2021, 1, 2)])
    b = make(tmp_path / "b.xlsx", rows)
    assert main([str(a), str(b), "--format", "json"]) == 0
    obj = json.loads(capsys.readouterr().out)
    assert obj["identical"] is False
    sheet = obj["sheets"][0]
    ops = [(r["op"], r.get("new_row")) for r in sheet["rows"]]
    assert ops == [("modify", 2), ("insert", 4)]
    assert sheet["rows"][0]["changes"][0]["value"] == {"old": "name1", "new": "改名"}
    assert sheet["rows"][1]["cells"][0]["value"] == 50
    assert sheet["rows"][1]["cells"][2]["value"] == "2021-01-02"


def test_stat_output(tmp_path, capsys):
    a = make(tmp_path / "a.xlsx", DATA)
    b = make(tmp_path / "b.xlsx", DATA[:2] + [[0, "z", 0]] + DATA[2:])
    main([str(a), str(b), "--stat"])
    out = capsys.readouterr().out
    assert "S | +1 rows" in out and "1 sheet(s) changed" in out


def test_exit_code(tmp_path):
    a = make(tmp_path / "a.xlsx", DATA)
    b = make(tmp_path / "b.xlsx", DATA[:-1])
    assert main([str(a), str(b)]) == 0
    assert main([str(a), str(b), "--exit-code"]) == 1
    assert main([str(a), str(a), "--exit-code"]) == 0


def test_sheet_filter(tmp_path, capsys, sample):
    def f(wb, ws):
        ws["A2"] = "x"
        wb.create_sheet("別")["A1"] = 1

    p = _edit_sample(sample, tmp_path, f)
    main([str(sample), str(p), "--sheet", "別"])
    out = capsys.readouterr().out
    assert "@@ 別" in out and "シート1" not in out.split("\n", 2)[2]


@pytest.mark.parametrize(
    "name,msg",
    [("missing.xlsx", "見つかりません"), ("old.xls", ".xls"), ("~$lock.xlsx", "ロックファイル")],
)
def test_errors(tmp_path, capsys, name, msg, sample):
    p = tmp_path / name
    if name != "missing.xlsx":
        p.write_bytes(b"dummy")
    assert main([str(p), str(sample)]) == 2
    assert msg in capsys.readouterr().err


def test_render_json_roundtrip_for_sample_variants(tmp_path, sample):
    p = _edit_sample(sample, tmp_path, lambda wb, ws: ws.insert_rows(3))
    json.loads(render_json(diff_workbooks(load(sample), load(p))))


def test_absolute_refs_across_insert_point(tmp_path):
    # 構成比 =B2/$B$20 や RANK の絶対参照は、合計行より上に行を挿入すると全行で書き換わる。
    # それでも挿入1行だけが表示されること
    def build(path, insert):
        vals = [(f"item{i}", i * 7 % 13 + 1) for i in range(1, 19)]
        if insert:
            vals.insert(8, ("new", 5))
        n = len(vals) + 1  # データ最終行
        wb = Workbook()
        ws = wb.active
        ws.title = "S"
        ws.append(["品目", "数量", "構成比", "順位", "平均差"])
        for r, (name, q) in enumerate(vals, start=2):
            ws.append([name, q, f"=B{r}/$B${n + 1}", f"=RANK(B{r},$B$2:$B${n})", f"=B{r}-$B${n + 2}"])
        ws.append(["合計", f"=SUM(B2:B{n})"])
        ws.append(["平均", f"=AVERAGE(B2:B{n})"])
        wb.save(path)
        return path

    a, b = build(tmp_path / "a.xlsx", False), build(tmp_path / "b.xlsx", True)
    assert body(a, b) == [
        "@@ S @@",
        '+ row 10: A10: "new" | B10: 5 | C10: =B10/$B$21 | D10: =RANK(B10,$B$2:$B$20) | E10: =B10-$B$22',
    ]


def test_bulk_column_edit_with_row_insert_on_large_sheet(tmp_path):
    # 全行で1列を書き換え + 1行挿入。完全一致する行がないため大きな区間になるが、
    # 全行削除+全行挿入に崩れず「挿入1行 + 変更1000セル」になること
    rows = [[i, f"name{i}", "未"] for i in range(1, 1001)]
    a = make(tmp_path / "a.xlsx", rows)
    rows2 = [[i, f"name{i}", "済"] for i in range(1, 1001)]
    rows2.insert(500, ["new", "x", "済"])
    b = make(tmp_path / "b.xlsx", rows2)
    lines = body(a, b)
    assert sum(l.startswith("+ row") for l in lines) == 1
    assert sum(l.startswith("- row") for l in lines) == 0
    assert sum(l.startswith("~ C") for l in lines) == 1000
    assert '+ row 501: A501: "new" | B501: "x" | C501: "済"' in lines
