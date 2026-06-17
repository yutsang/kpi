r"""答「為什麼會有 blank V」—— 讀每家 kpi_report.parquet（prep_tableau 填之前），睇 vertical 空白嘅行係咩。

blank V = 我哋 vertical 分類解唔到嗰啲行。逐家拆：by NG / by H(開支類型) / top account_desc。
通常大舊係 comp 行（Comp 餐飲/房間）—— comp 冇對應一個明確非博彩 vertical。

Run:  python scripts/diag_blank_v.py
出 results/diag_blank_v.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
ENT = {"galaxy": 1, "sjm": 2, "wynn": 3, "vml": 4, "melco": 5, "mgm": 6}
AMT = ["amount_mop", "Amount - Amended", "ledger_amount"]


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    L = ["# diag_blank_v — blank V 根因（prep_tableau 填之前，逐家拆 by NG / H / account）"]
    for ent, n in ENT.items():
        p = ROOT / "data" / ent / "output" / f"company_{n}_kpi_report.parquet"
        L.append(f"\n{'='*60}\n── {ent} ──")
        if not p.exists():
            L.append("   !! parquet 揾唔到"); continue
        df = pd.read_parquet(p)
        amtc = next((c for c in AMT if c in df.columns), None)
        amt = pd.to_numeric(df[amtc], errors="coerce").fillna(0.0) / 1e4 if amtc else pd.Series(0.0, index=df.index)
        vid = _s(df["vertical_id"]) if "vertical_id" in df.columns else pd.Series("", index=df.index)
        blank = vid == ""
        L.append(f"   blank V: {int(blank.sum()):,}/{len(df):,} 行 ({blank.mean()*100:.0f}%)  Σ={amt[blank].sum()/1e2:,.0f}M")
        if not blank.any():
            continue
        # by NG label
        ngl = _s(df["ng_label"]) if "ng_label" in df.columns else _s(df.get("ng_code", pd.Series("", index=df.index)))
        by_ng = amt[blank].groupby(ngl[blank]).sum().abs().sort_values(ascending=False).head(6)
        L.append("   by NG: " + "  ".join(f"{k or '(空)'}={v/1e2:,.0f}M" for k, v in by_ng.items()))
        # by H label（開支類型 — 多數係 comp？）
        if "horizontal_label" in df.columns:
            hl = _s(df["horizontal_label"])
            by_h = amt[blank].groupby(hl[blank]).sum().abs().sort_values(ascending=False).head(6)
            L.append("   by H: " + "  ".join(f"{k or '(空)'}={v/1e2:,.0f}M" for k, v in by_h.items()))
        # top account_desc
        adc = next((c for c in ("account_desc", "ledger_account", "Cost element descr.", "Nature of Expenses") if c in df.columns), None)
        if adc:
            ad = _s(df[adc])
            by_a = amt[blank].groupby(ad[blank]).sum().abs().sort_values(ascending=False).head(8)
            L.append("   top account: " + "  ".join(f"{str(k)[:14]}={v/1e2:,.0f}M" for k, v in by_a.items()))
    _w(L)


def _w(L):
    out = ROOT / "results" / "diag_blank_v.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
