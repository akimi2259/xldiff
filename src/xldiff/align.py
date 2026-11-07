"""行・列の対応付け(アライメント)。

git diff と同様に、まず完全一致する行(列)を最長共通部分列で対応付ける。
一致しなかった区間は類似度による動的計画法で「変更された行」と
「挿入・削除された行」を見分ける。
"""

from __future__ import annotations

from bisect import bisect_left, bisect_right
from difflib import SequenceMatcher
from typing import Callable, Hashable, Sequence

Pair = tuple[int | None, int | None]
SimFn = Callable[[int, int], float]

# 全探索 DP の計算量上限(区間の行数 m × n)。超えたら帯状 DP にする
DP_LIMIT = 250_000
# 帯状 DP の帯幅の余裕(行数差に加える)
BAND_MARGIN = 32


def align(
    old_keys: Sequence[Hashable],
    new_keys: Sequence[Hashable],
    sim: SimFn | None = None,
    threshold: float = 0.5,
) -> list[Pair]:
    """0 始まりのインデックスの対応表を順序付きで返す。

    (i, j): 対応する要素 / (i, None): 削除 / (None, j): 挿入
    """
    # autojunk=True だと 200 要素超で頻出要素(空行など)が無視され、対応付けが崩れる
    sm = SequenceMatcher(None, old_keys, new_keys, autojunk=False)
    out: list[Pair] = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            out.extend(zip(range(i1, i2), range(j1, j2)))
        elif tag == "delete":
            out.extend((i, None) for i in range(i1, i2))
        elif tag == "insert":
            out.extend((None, j) for j in range(j1, j2))
        else:
            out.extend(_resolve(range(i1, i2), range(j1, j2), sim, threshold))
    return out


def _resolve(olds: range, news: range, sim: SimFn | None, threshold: float) -> list[Pair]:
    m, n = len(olds), len(news)
    if sim is None:
        return _fill_gap(list(olds), list(news))
    if m * n <= DP_LIMIT:
        band = None
    else:
        # 大きな区間(例: 全行で1列を書き換えた上で1行挿入)は、対角線付近だけを探索する
        band = abs(m - n) + BAND_MARGIN
        if m * (2 * band + 2) > DP_LIMIT * 8:
            return _positional(list(olds), list(news))

    def s_at(i: int, j: int) -> float:
        return sim(olds[i], news[j])

    matched = _dp_match(m, n, s_at, threshold, band)

    out: list[Pair] = []
    pi = pj = 0
    for mi, mj in matched + [(m, n)]:
        out.extend(_fill_gap([olds[k] for k in range(pi, mi)], [news[k] for k in range(pj, mj)]))
        if mi < m:
            out.append((olds[mi], news[mj]))
        pi, pj = mi + 1, mj + 1
    return out


_NEG = float("-inf")


def _dp_match(
    m: int, n: int, s_at: Callable[[int, int], float], threshold: float, band: int | None
) -> list[tuple[int, int]]:
    """順序を保ったまま類似度の合計が最大になる対応(区間内の 0 始まり添字)を返す。

    band を指定すると、j が i*n/m ± band の範囲だけを探索する(帯状 DP)。
    score(i, j) は olds[i:], news[j:] での最大合計類似度。帯の外は到達不能(-inf)。
    """
    if band is None:
        lo = [0] * (m + 1)
        hi = [n] * (m + 1)
    else:
        lo, hi = [], []
        for i in range(m + 1):
            c = i * n // m if m else 0
            lo.append(max(0, c - band))
            hi.append(min(n, c + 1 + band))

    def get(row: list[float], L: int, H: int, j: int) -> float:
        return row[j - L] if L <= j <= H else _NEG

    score: list[list[float]] = [[]] * (m + 1)
    score[m] = [0.0] * (hi[m] - lo[m] + 1)  # 残りの新しい行はすべて挿入
    sims: dict[tuple[int, int], float] = {}
    for i in range(m - 1, -1, -1):
        L, H = lo[i], hi[i]
        nxt, nL, nH = score[i + 1], lo[i + 1], hi[i + 1]
        row = [_NEG] * (H - L + 1)
        for j in range(H, L - 1, -1):
            if j == n:
                row[j - L] = 0.0  # 残りの古い行はすべて削除
                continue
            best = get(nxt, nL, nH, j)  # olds[i] を削除
            if j + 1 <= H and row[j + 1 - L] > best:  # news[j] を挿入
                best = row[j + 1 - L]
            diag = get(nxt, nL, nH, j + 1)
            if diag > _NEG:
                s = s_at(i, j)
                sims[(i, j)] = s
                if s >= threshold and s + diag >= best:
                    best = s + diag
            row[j - L] = best
        score[i] = row

    matched: list[tuple[int, int]] = []
    i = j = 0
    while i < m and j < n:
        cur = get(score[i], lo[i], hi[i], j)
        s = sims.get((i, j), 0.0)
        diag = get(score[i + 1], lo[i + 1], hi[i + 1], j + 1)
        if s >= threshold and diag > _NEG and abs(cur - (s + diag)) < 1e-9:
            matched.append((i, j))
            i += 1
            j += 1
        elif abs(cur - get(score[i + 1], lo[i + 1], hi[i + 1], j)) < 1e-9:
            i += 1
        else:
            j += 1
    return matched


def _positional(olds: list[int], news: list[int]) -> list[Pair]:
    """探索しきれないほど大きな区間の最終手段。先頭から位置で対応付ける。"""
    k = min(len(olds), len(news))
    return (
        list(zip(olds[:k], news[:k]))
        + [(i, None) for i in olds[k:]]
        + [(None, j) for j in news[k:]]
    )


def _fill_gap(olds: list[int], news: list[int]) -> list[Pair]:
    """類似しない行同士の隙間。同数なら「その場で書き換えた」とみなして対応付ける。"""
    if len(olds) == len(news):
        return list(zip(olds, news))
    return [(i, None) for i in olds] + [(None, j) for j in news]


class AxisMap:
    """旧ファイルの行(列)番号 → 新ファイルの行(列)番号。1 始まり。"""

    def __init__(self, pairs: list[Pair], old_max: int) -> None:
        self.fwd: dict[int, int] = {i: j for i, j in pairs if i is not None and j is not None}
        self.old_max = old_max
        self._sorted = sorted(self.fwd)
        if self._sorted:
            last = self._sorted[-1]
            self._tail = self.fwd[last] - last
        else:
            self._tail = 0

    @classmethod
    def identity(cls) -> "AxisMap":
        return cls([], 0)

    def map(self, i: int, role: str = "single") -> int | None:
        j = self.fwd.get(i)
        if j is not None:
            return j
        if i > self.old_max:
            return i + self._tail
        # 削除された位置。範囲の端点なら Excel と同様に範囲を縮める
        if role == "start":
            k = bisect_right(self._sorted, i)
            return self.fwd[self._sorted[k]] if k < len(self._sorted) else None
        if role == "end":
            k = bisect_left(self._sorted, i) - 1
            return self.fwd[self._sorted[k]] if k >= 0 else None
        return None
