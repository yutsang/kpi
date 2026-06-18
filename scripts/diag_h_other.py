r"""Inspect H reasonableness（user 2026-06-18）：
  ① H_OTHER（其他）逐家拆 account_desc（signed）—— 答 wynn 負數 / galaxy·vml 過大，究竟入面係咩
  ② capex 唔合理 H —— capex 行但 H ∈ {人工/藝人/其他/comp/廣告} 嘅，by account（capex 應該係建設/器具）
  ③ sjm comp（活動場地/餐飲/房間…）總額 + 邊啲 account 似 comp 但冇標

Run:  python scripts/diag_h_other.py
出 results/diag_h_other.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
COMP = {"H_VENUE", "H_FNB", "H_HOTEL_ROOM", "H_COMP_TICKET", "H_COMP_OTHER"}
CAPEX_BAD = {"H_LABOR", "H_PERFORMER", "H_OTHER", "H_FNB", "H_HOTEL_ROOM",
             "H_COMP_OTHER", "H_COMP_TICKET", "H_ADVERTISING", "H_VENUE"}


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    L = ["# diag_h_other — ① H_OTHER 拆解 ② capex 唔合理 H ③ sjm comp"]
    if not CSV.exists():
        L.append("!! csv 揾唔到"); _w(L); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    df["amt"] = pd.to_numeric(df["amount_mop"], errors="coerce").fillna(0.0) / 1e4
    ent = df["entity"].astype(str)
    hid = _s(df["horizontal_id"]); co = _s(df["final_capex_opex"])
    acc = _s(df["account_desc"]) if "account_desc" in df.columns else pd.Series("", index=df.index)

    # ① H_OTHER 拆解
    L.append("\n① H_OTHER（其他）逐家拆 account_desc（signed 萬）：")
    for e in sorted(ent.unique()):
        m = (ent == e) & hid.eq("H_OTHER")
        if not m.any():
            continue
        tot = df.loc[m, "amt"].sum()
        L.append(f"\n   ── {e}  H_OTHER Σ={tot:,.0f}萬 ──")
        by = df[m].groupby(acc[m])["amt"].sum().sort_values()
        neg = by[by < -50]
        pos = by[by > 50].sort_values(ascending=False)
        if len(neg):
            L.append("     負數: " + " | ".join(f"{str(k)[:20]}={v:,.0f}" for k, v in neg.head(6).items()))
        L.append("     正數top: " + " | ".join(f"{str(k)[:20]}={v:,.0f}" for k, v in pos.head(8).items()))

    # ② capex 唔合理 H
    L.append("\n\n② capex 行但 H 唔似 capex（應係建設/器具）by account 萬：")
    for e in sorted(ent.unique()):
        m = (ent == e) & co.eq("Capex") & hid.isin(CAPEX_BAD)
        if not m.any():
            continue
        s = df.loc[m, "amt"].sum()
        L.append(f"\n   ── {e}  capex-唔合理H Σ={s:,.0f}萬 ──")
        by = df[m].groupby([hid[m], acc[m]])["amt"].sum().sort_values(ascending=False)
        for (h, a), v in by.head(10).items():
            if abs(v) > 100:
                L.append(f"     {h:<14} {str(a)[:30]:<30} {v:,.0f}")

    # ③ sjm comp
    L.append("\n\n③ sjm comp 現況（活動場地/餐飲/房間/贈品/其他comp）：")
    sm = ent.eq("sjm")
    for h in ["H_VENUE", "H_FNB", "H_HOTEL_ROOM", "H_COMP_TICKET", "H_COMP_OTHER"]:
        v = df.loc[sm & hid.eq(h), "amt"].sum()
        L.append(f"     {h:<14} {v:,.0f}萬")
    L.append("   sjm 邊啲 account_desc 似 comp 但標咗第二樣（top 非-comp by account 含 comp/餐/房/票 字眼）：")
    sus = sm & ~hid.isin(COMP) & acc.str.contains("comp|Comp|餐|房|客房|門票|贈|食|飲|F&B|Hotel|Room", case=False, regex=True, na=False)
    if sus.any():
        by = df[sus].groupby([acc[sus], hid[sus]])["amt"].sum().sort_values(ascending=False)
        for (a, h), v in by.head(10).items():
            if abs(v) > 50:
                L.append(f"     {str(a)[:28]:<28} [{h}] {v:,.0f}萬")
    _w(L)


def _w(L):
    out = ROOT / "results" / "diag_h_other.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
