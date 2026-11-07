"""動作確認用の架空サンプルと、その編集版を examples/ に生成する。

    uv run python scripts/make_variants.py
    uv run xldiff examples/base.xlsx examples/06_mixed.xlsx

base.xlsx は講義の録画一覧を模した架空データ(URL は example.com)。
openpyxl の insert_rows / delete_rows はハイパーリンクや結合範囲をずらさないため、
行の挿入・削除は「最終状態を書き直す」方式で作る。
"""

from __future__ import annotations

import copy
import datetime as dt
import sys
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Font, PatternFill

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "examples"
BASE = OUT / "base.xlsx"

NUMS = ["一", "二", "三", "四", "五", "六", "七", "八", "九", "十", "十一", "十二", "十三"]


def date_text(day: dt.date) -> str:
    return f"日付: {day.year}年{day.month}月{day.day}日 10:00 AM"


def build_base(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "シート1"
    ws["A1"] = "トピック: サンプル講義2025"
    for i, n in enumerate(NUMS):
        r = i + 2
        ws.cell(r, 1, f"第{n}回").font = Font(name="Arial", color="FF333333")
        ws.cell(r, 2, date_text(dt.date(2025, 4, 7) + dt.timedelta(weeks=i)))
        url = f"https://example.com/rec/{i + 1:02d}"
        c = ws.cell(r, 3, url)
        c.hyperlink = url
        c.style = "Hyperlink"
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 40
    wb.save(path)


def snapshot_rows(ws):
    """[(values, styles, links)] を行ごとに取り出す(A〜C 列)。"""
    rows = []
    for r in range(1, ws.max_row + 1):
        vals, styles, links = [], [], []
        for c in range(1, 4):
            cell = ws.cell(r, c)
            vals.append(cell.value)
            styles.append(copy.copy(cell._style))
            links.append(cell.hyperlink.target if cell.hyperlink else None)
        rows.append((vals, styles, links))
    return rows


def write_rows(ws, rows, ncols=3):
    """シートの A〜(ncols) 列を rows の内容で書き直す。"""
    for r in range(1, max(ws.max_row, len(rows)) + 1):
        for c in range(1, ncols + 1):
            cell = ws.cell(r, c)
            cell.value = None
            cell.hyperlink = None
            cell.style = "Normal"
    for r, (vals, styles, links) in enumerate(rows, start=1):
        for c, (v, st, ln) in enumerate(zip(vals, styles, links), start=1):
            cell = ws.cell(r, c)
            cell.value = v
            if st is not None:
                cell._style = copy.copy(st)
            if ln:
                cell.hyperlink = ln


def open_base():
    wb = load_workbook(BASE)
    return wb, wb.worksheets[0]


def save(wb, name: str) -> Path:
    p = OUT / name
    wb.save(p)
    return p


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    build_base(BASE)
    made = [BASE]

    # 01: 何も変えずに保存し直す(正規化の確認用。差分ゼロになるべき)
    wb, ws = open_base()
    made.append(save(wb, "01_resave.xlsx"))

    # 02: セルの値を1か所変更(第二回の日付)
    wb, ws = open_base()
    ws["B3"] = date_text(dt.date(2025, 4, 15))
    made.append(save(wb, "02_cell_edit.xlsx"))

    # 03: 第三回の後ろに補講の行を挿入
    wb, ws = open_base()
    rows = snapshot_rows(ws)
    tmpl = rows[3]
    extra_url = "https://example.com/rec/03-extra"
    new_row = (
        ["第三回補講", date_text(dt.date(2025, 4, 23)), extra_url],
        tmpl[1],
        [None, None, extra_url],
    )
    rows.insert(4, new_row)
    write_rows(ws, rows)
    made.append(save(wb, "03_row_insert.xlsx"))

    # 04: 第八回の行を削除
    wb, ws = open_base()
    rows = snapshot_rows(ws)
    del rows[8]
    write_rows(ws, rows)
    made.append(save(wb, "04_row_delete.xlsx"))

    # 05: B 列に「担当」列を挿入(既存の B, C は C, D へ)
    wb, ws = open_base()
    rows = snapshot_rows(ws)
    new_rows = []
    for i, (vals, styles, links) in enumerate(rows):
        tutor = None if i == 0 else ("田中" if i % 2 else "佐藤")
        new_rows.append(
            (
                [vals[0], tutor, vals[1], vals[2]],
                [styles[0], styles[1], styles[1], styles[2]],
                [links[0], None, links[1], links[2]],
            )
        )
    write_rows(ws, new_rows, ncols=4)
    ws.column_dimensions["D"].width = ws.column_dimensions["C"].width
    ws.column_dimensions["C"].width = ws.column_dimensions["B"].width
    ws.column_dimensions["B"].width = 10
    made.append(save(wb, "05_col_insert.xlsx"))

    # 06: 複合(行挿入+セル変更+書式+リンク+コメント+結合+シート追加)
    wb, ws = open_base()
    rows = snapshot_rows(ws)
    rows.insert(4, new_row)
    rows[9][0][1] = date_text(dt.date(2025, 5, 27))  # 第八回(旧 9 行目 → 新 10 行目)の日付
    write_rows(ws, rows)
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"].fill = PatternFill("solid", fgColor="FFFF00")
    ws["C3"].hyperlink = "https://example.com/rec/02-fixed"
    ws["B7"].comment = Comment("音声が途中で切れている", "reviewer")
    ws.merge_cells("A1:C1")
    memo = wb.create_sheet("メモ")
    memo["A1"] = "未視聴"
    memo["B1"] = 3
    memo["A2"] = "合計"
    memo["B2"] = "=COUNTA(シート1!A2:A15)"
    made.append(save(wb, "06_mixed.xlsx"))

    # 07: シート名の変更と、表示形式・列幅の変更
    wb, ws = open_base()
    ws.title = "講義一覧"
    ws.column_dimensions["C"].width = 60
    ws["A1"].number_format = "@"
    made.append(save(wb, "07_rename_format.xlsx"))

    for p in made:
        print(p.relative_to(ROOT).as_posix())


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    main()
