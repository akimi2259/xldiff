"""テスト用の架空サンプルブック。

講義の録画一覧を模した 1 シートのブックを生成する。
A1 にタイトル、2〜14 行目に「回・日付・録画 URL」を置く。
URL セルにはハイパーリンクを設定し、A 列はフォント指定、B・C 列は列幅指定あり。
"""

from __future__ import annotations

import datetime as dt
from pathlib import Path

import pytest
from openpyxl import Workbook
from openpyxl.styles import Font

NUMS = ["一", "二", "三", "四", "五", "六", "七", "八", "九", "十", "十一", "十二", "十三"]


def build_sample(path: Path) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "シート1"
    ws["A1"] = "トピック: サンプル講義2025"
    for i, n in enumerate(NUMS):
        r = i + 2
        day = dt.date(2025, 4, 7) + dt.timedelta(weeks=i)
        ws.cell(r, 1, f"第{n}回").font = Font(name="Arial", color="FF333333")
        ws.cell(r, 2, f"日付: {day.year}年{day.month}月{day.day}日 10:00 AM")
        url = f"https://example.com/rec/{i + 1:02d}"
        c = ws.cell(r, 3, url)
        c.hyperlink = url
        c.style = "Hyperlink"
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 40
    wb.save(path)
    return path


@pytest.fixture(scope="session")
def sample(tmp_path_factory) -> Path:
    return build_sample(tmp_path_factory.mktemp("sample") / "sample.xlsx")
