"""xldiff のコマンドラインインターフェース。"""

from __future__ import annotations

import argparse
import os
import sys

from . import __version__
from .differ import Options, diff_workbooks
from .loader import LoadError, load
from .render import Compact, render_json, render_stat

EPILOG = """\
出力の読み方 (compact 形式):
  @@ シート名 @@          シートの区切り
  ~ B3: 旧 → 新            セルの変更(新ファイルの座標。位置が動いたら "(was B2)")
  ~ B3 style/link/comment  書式・ハイパーリンク・コメントの変更
  + row 5: A5: ... | B5: ...  挿入された行(新ファイルの行番号)
  - row 9: A9: ... | B9: ...  削除された行(旧ファイルの行番号)
  + col D / - col D        挿入・削除された列
  文字列は "..." で囲み、数値・日付・数式(=...)は囲まない。(empty) は空セル。

終了コード: 0=正常(差分の有無に関わらず) / 1=差分あり(--exit-code 指定時のみ) / 2=エラー
"""


def _enable_windows_vt() -> bool:
    if os.name != "nt":
        return True
    try:
        import ctypes

        k = ctypes.windll.kernel32
        h = k.GetStdHandle(-11)
        mode = ctypes.c_uint32()
        if not k.GetConsoleMode(h, ctypes.byref(mode)):
            return False
        return bool(k.SetConsoleMode(h, mode.value | 0x0004))
    except Exception:
        return False


def _use_color(choice: str) -> bool:
    if choice == "always":
        return True
    if choice == "never":
        return False
    if os.environ.get("NO_COLOR"):
        return False
    if not sys.stdout.isatty():
        return False
    return _enable_windows_vt()


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="xldiff",
        description="2つの Excel ファイル (.xlsx/.xlsm) の差分を git diff 風に表示します。",
        epilog=EPILOG,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    p.add_argument("old", help="比較元(変更前)のファイル")
    p.add_argument("new", help="比較先(変更後)のファイル")
    p.add_argument(
        "-f",
        "--format",
        choices=["compact", "json"],
        default="compact",
        help="出力形式 (既定: compact)",
    )
    p.add_argument("--stat", action="store_true", help="シートごとの件数サマリだけを表示")
    p.add_argument("--no-style", action="store_true", help="書式(フォント・塗り・罫線・表示形式・行高・列幅)の差分を無視")
    p.add_argument(
        "-s",
        "--sheet",
        action="append",
        default=[],
        metavar="NAME",
        help="対象シートを絞り込む(複数指定可)",
    )
    p.add_argument(
        "--color",
        choices=["auto", "always", "never"],
        default="auto",
        help="カラー表示 (既定: auto = 端末に出力するときだけ色付け)",
    )
    p.add_argument("--max-len", type=int, default=0, metavar="N", help="値の表示を N 文字で切り詰める (0=切り詰めない)")
    p.add_argument("--indent", type=int, default=None, metavar="N", help="JSON のインデント幅 (既定: 1行に詰める)")
    p.add_argument("--exit-code", action="store_true", help="差分があれば終了コード 1 を返す")
    p.add_argument("-V", "--version", action="version", version=f"xldiff {__version__}")
    return p


def main(argv: list[str] | None = None) -> int:
    # パイプ出力時に cp932 で化けないよう UTF-8 に固定し、改行も OS によらず LF にする
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace", newline="\n")
        except Exception:
            pass

    args = build_parser().parse_args(argv)
    try:
        old = load(args.old)
        new = load(args.new)
    except LoadError as e:
        print(f"xldiff: error: {e}", file=sys.stderr)
        return 2

    wd = diff_workbooks(old, new, Options(style=not args.no_style, sheets=args.sheet))
    wd.old_path, wd.new_path = args.old, args.new
    color = _use_color(args.color)

    if args.format == "json":
        out = render_json(wd, indent=args.indent)
    elif args.stat:
        out = render_stat(wd, color=color)
    else:
        out = Compact(color=color, max_len=args.max_len).render(wd)
    sys.stdout.write(out)
    sys.stdout.flush()

    if args.exit_code and not wd.is_empty():
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
