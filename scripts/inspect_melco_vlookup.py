r"""inspect_melco_vlookup —— melco adj 缺 建築設施設備劃分/KP识别表演演出，
能唔能由舊 melco_2025.xlsx 用 KPMG唯一識別碼 1:1 vlookup 過去？

user 2026-06-24：melco adj 已經係項目組嘅；查現有(舊)melco 有冇嗰 2 條值 + 對唔對到 uid。
報告：adj 行數、adj uid 喺舊檔(可 carry)、adj uid 唔喺舊檔(新行→2欄會空)、影響金額、2 欄 value 分佈。

Run（Windows）:  python scripts\inspect_melco_vlookup.py
出 results\inspect_melco_vlookup.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import warnings
warnings.filterwarnings("ignore")
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "results" / "inspect_melco_vlookup.txt"
OLD = ROOT / "data" / "melco" / "raw" / "melco_2025.xlsx"
ADJ = ROOT / "data" / "melco" / "raw" / "melco_2025_adj.xlsx"
UID = "KPMG唯一識別碼"
COLS = ["建築設施設備劃分", "KP识别表演演出"]
AMT = "Amount - Amended"


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    L = ["# inspect_melco_vlookup —— 舊 melco 2 欄能否由 uid vlookup 去 adj"]
    if not OLD.exists() or not ADJ.exists():
        L.append(f"!! 揾唔到 OLD={OLD.exists()} ADJ={ADJ.exists()}"); _w(L); return
    old = pd.read_excel(OLD, sheet_name="Data", dtype=object)
    adj = pd.read_excel(ADJ, sheet_name="Data", dtype=object)
    L.append(f"  舊 melco_2025.xlsx[Data] 行={len(old):,} | adj[Data] 行={len(adj):,}")

    if UID not in old.columns or UID not in adj.columns:
        L.append(f"  !! uid「{UID}」唔喺 old={UID in old.columns} adj={UID in adj.columns}"); _w(L); return
    miss_old = [c for c in COLS if c not in old.columns]
    if miss_old:
        L.append(f"  !! 舊檔都冇呢啲欄：{miss_old}（咁 vlookup 唔到，要項目組）"); _w(L); return

    ou = _s(old[UID]); au = _s(adj[UID])
    oset = set(ou[ou.ne("")])
    # adj uid 重複/空
    a_blank = int(au.eq("").sum())
    a_dup = int(au[au.ne("")].duplicated().sum())
    matched = au.isin(oset) & au.ne("")
    L.append(f"\n  ── uid 對賬 ──")
    L.append(f"     adj uid 空={a_blank} | adj uid 重複={a_dup}")
    L.append(f"     adj 行對到舊檔 uid = {int(matched.sum()):,} / {len(adj):,}  ({matched.mean()*100:.1f}%)")
    unm = adj[~matched]
    L.append(f"     ⚠ adj 對唔到(新行/uid 不符) = {len(unm):,} 行 → 呢啲 2 欄會空 → H 跌返 default")
    if AMT in adj.columns and len(unm):
        _a = pd.to_numeric(unm[AMT].astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0).abs().sum() / 1e4
        L.append(f"        對唔到行金額 |Amount-Amended| = {_a:,.0f}萬")

    # 舊檔 2 欄 value 分佈（睇有幾多 row 真係有 label）
    L.append(f"\n  ── 舊檔 2 欄 value 分佈（vlookup 帶過去嘅內容）──")
    for c in COLS:
        v = _s(old[c]); nonblank = int(v.ne("").sum())
        L.append(f"     {c}：非空 {nonblank:,}/{len(old):,}")
        for val, n in v[v.ne("")].value_counts().head(8).items():
            L.append(f"         {val[:30]:<30}{n:>8,}")

    # 判定
    L.append(f"\n  ── 判定 ──")
    if matched.mean() > 0.98:
        L.append(f"     ✅ uid 對得好齊（{matched.mean()*100:.1f}%）→ 用 VLOOKUP({UID}) 由舊檔搬 2 欄入 adj 就得")
    else:
        L.append(f"     ⚠ 有 {len(unm):,} 行對唔到 → 嗰啲 2 欄會空(H 用 account/desc fallback)；睇上面金額決定影響")
    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
