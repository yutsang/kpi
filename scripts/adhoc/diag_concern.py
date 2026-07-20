#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
diag_concern.py — 睇實一個 source_1 檔 + 佢 source_2 overlay，逐個項目 dump
「該關注事項涉及調整金額」(表頭 Z) 相關嘅所有數字欄，搞清楚 source_2 究竟把
「涉及調整金額」放喺邊條欄（該關注事項涉及調整金額 / 建議調整金額 / 潛在調整合計…），
從而知道 align 寫落 Z 嘅值點解「唔跟 source_2」。

用法（Windows）：
    python scripts\\adhoc\\diag_concern.py ^
        "ad-hoc\\workspace\\source_1\\旅遊局\\SJM-投資計劃執行情況表二（旅遊局）.xlsx" ^
        "ad-hoc\\workspace\\source_2\\0714\\旅遊局\\SJM-投資計劃執行情況表二（旅遊局）.xlsx"

第 2 個參數（source_2 overlay 檔）可以唔俾 → 會由 --root 自動揾（預設 ad-hoc\\workspace）。
    python scripts\\adhoc\\diag_concern.py "…source_1\\旅遊局\\SJM-…（旅遊局）.xlsx" --root ad-hoc\\workspace
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import align_to_header as A


def _colof(col_gs, label):
    k = A.nkey(label)
    return next((c for c, (g, s) in col_gs.items() if s == k), None)


def _rowvals(ws, col, r0, r1):
    if col is None:
        return []
    out = []
    for r in range(r0, r1 + 1):
        v = ws.cell(r, col).value
        if v is not None and A._s(v) != "":
            out.append((r, v))
    return out


def _sumnum(ws, col, r0, r1):
    """逐行去重加 abs（同 _concern_sum 邏輯），dump 用。"""
    if col is None:
        return None
    tot, prev, found = 0.0, None, False
    for r in range(r0, r1 + 1):
        a = A.num(ws.cell(r, col).value)
        if a is None or a == prev:
            continue
        prev = a
        tot += abs(a)
        found = True
    return tot if found else None


# 想睇嘅數字欄（source_2 側可能其中一條先係「涉及調整金額」真身）
_MONEY_SUBS = [
    "該關注事項涉及調整金額",
    "建議調整金額",
    "建議調整後金額",
    "潛在調整合計",
    "調整後投資金額",
    "申報投資金額",
    "跨司工作組確認投資金額",
    "建議接納之調整後金額",
]


def dump_side(tag, ws, projs, col_gs):
    print(f"\n  ── {tag}：欄位偵測 ──")
    for lab in _MONEY_SUBS:
        c = _colof(col_gs, lab)
        print(f"      {lab:<18} → {A.get_column_letter(c)+f'({c})' if c else '（冇呢欄）'}")
    print(f"\n  ── {tag}：逐項目 ──")
    for p in projs:
        sk = A._seqkey(p.seq)
        print(f"\n    ▸ [{sk}] {A._s(p.seq)[:40]}  rows {p.r0}-{p.r1}")
        for lab in _MONEY_SUBS:
            c = _colof(col_gs, lab)
            if c is None:
                continue
            rv = _rowvals(ws, c, p.r0, p.r1)
            sm = _sumnum(ws, c, p.r0, p.r1)
            if rv:
                cells = ", ".join(f"r{r}={A._s(v)[:18]}" for r, v in rv)
                print(f"        {lab:<18} 加總={A.fmt_amt(sm) if sm is not None else '—':<12} [{cells}]")
        # align 對呢項目計出嘅 source_1 側 Z fallback
        t = A._adj_total(p)
        print(f"        └ source_1 abs_total(|潛在調整合計|) = "
              f"{A.fmt_amt(abs(t)) if t is not None else 'None'}   has_adj={p.has_adj()}")


def main():
    args = [a for a in sys.argv[1:]]
    root = Path("ad-hoc/workspace")
    if "--root" in args:
        i = args.index("--root")
        root = Path(args[i + 1])
        del args[i:i + 2]
    if not args:
        print("俾 source_1 檔路徑（第 2 個參數 = source_2 overlay 檔，可省）")
        return
    s1 = Path(args[0])
    s2 = Path(args[1]) if len(args) > 1 else None

    log = lambda *a, **k: None  # extract 內部 log 靜音
    print(f"{'='*74}\n# source_1: {s1.name}")
    wb1 = A.load_wb(s1)
    for sn in wb1.sheetnames:
        ws1 = wb1[sn]
        projs1, subrow, anchor, gm, maxcol, col_gs1, maxrow = A.extract(ws1, log)
        if not anchor or not projs1:
            continue
        print(f"\n{'-'*66}\n# source_1 sheet {sn!r}（{len(projs1)} 項目）")
        dump_side("source_1", ws1, projs1, col_gs1)

    # ── source_2 side ──
    if s2 is None:
        scope, company = A.infer_scope_company(s1.relative_to(root / "source_1").as_posix())
        s2 = A.find_overlay_file(root, scope, company)
        print(f"\n（自動揾 overlay：scope={scope} company={company} → {s2}）")
    if not s2 or not Path(s2).exists():
        print("\n✗ 冇 source_2 overlay 檔 → 淨係得 source_1 側")
        return
    print(f"\n{'='*74}\n# source_2 overlay: {Path(s2).name}")
    wb2 = A.load_wb(Path(s2))
    for sn in wb2.sheetnames:
        ws2 = wb2[sn]
        projs2, subrow, anchor, gm, maxcol, col_gs2, maxrow = A.extract(ws2, log)
        if not anchor or not projs2:
            continue
        print(f"\n{'-'*66}\n# source_2 sheet {sn!r}（{len(projs2)} 項目）")
        dump_side("source_2", ws2, projs2, col_gs2)
        # 每項目：_concern_sum 實際返乜（align 會攞呢個做 override）
        print(f"\n  ── source_2 sheet {sn!r}：_concern_sum（align 會寫落 Z override）──")
        for p in projs2:
            cs = A._concern_sum(ws2, p, col_gs2)
            print(f"      [{A._seqkey(p.seq)}] _concern_sum = "
                  f"{A.fmt_amt(cs) if cs is not None else 'None（→ Z 跌返 source_1）'}")


if __name__ == "__main__":
    main()
