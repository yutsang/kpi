r"""讀 vml tagged_rows，拆 H_LABOR(opex) by 分類1 / 分類1.1 / 會計科目分類 —— 睇邊個 mark 值谷大人工。
（分類1 冇 carry 落 CSV，所以讀 data\vml\interim\company_4_tagged_rows.parquet）

HQ vml staff 細（790/1457/1253），但我哋人工 16,107萬。懷疑 conf 將 管理成本分攤/部門間成本分攤
（cost allocation）map 咗去 H_LABOR。呢個 dump 證實。

Run（Windows）:  python scripts\inspect_vml_div1.py
出 results\inspect_vml_div1.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
TR = ROOT / "data/vml/interim/company_4_tagged_rows.parquet"
L: list[str] = ["# inspect_vml_div1 —— vml H_LABOR(opex) by 分類1（萬，tagged_rows raw amount）"]


def main():
    if not TR.exists():
        L.append(f"!! {TR} 唔存在（跑過 vml kedro 先有）"); _w(); return
    df = pd.read_parquet(TR)
    L.append(f"  tagged_rows {len(df):,} 行；欄含分類: " + ", ".join(c for c in df.columns if "分類" in str(c) or "科目分類" in str(c)))
    hid = next((c for c in ("horizontal_id",) if c in df.columns), None)
    co = next((c for c in ("final_capex_opex", "CAPEX_OPEX") if c in df.columns), None)
    amtc = next((c for c in ("amount_mop", "MOP Amt", "amount") if c in df.columns), None)
    bk = next((c for c in ("report_period",) if c in df.columns), None)
    if not (hid and amtc):
        L.append(f"!! 揾唔到 horizontal_id / amount 欄（hid={hid} amt={amtc}）"); _w(); return
    df["amt"] = pd.to_numeric(df[amtc], errors="coerce").fillna(0.0) / 1e4
    df["yb"] = df[bk].astype(str) if bk else "?"
    opex = ~df[co].astype(str).str.strip().eq("Capex") if co else pd.Series(True, index=df.index)
    lab = df[(df[hid].astype(str).str.strip() == "H_LABOR") & opex]
    L.append(f"\n  vml H_LABOR(opex) 共 {lab['amt'].sum():,.0f}萬（{len(lab):,}行）")

    for col in ("分類1", "分類1.1", "會計科目分類", "Nature"):
        c = next((x for x in df.columns if str(x).strip() == col), None)
        if not c:
            L.append(f"\n  ── 「{col}」冇呢欄 ──"); continue
        L.append(f"\n  ── H_LABOR by 「{col}」（萬 + 每 bucket）──")
        v = lab[c].astype(str).str.strip().replace({"nan": "", "None": ""})
        g = lab.assign(_v=v).groupby("_v")["amt"].sum().sort_values(key=lambda s: s.abs(), ascending=False)
        for k, tot in g.head(12).items():
            per = lab[lab.index.isin(lab.assign(_v=v)[lambda d: d["_v"] == k].index)].groupby("yb")["amt"].sum()
            ps = " ".join(f"{b}={x:,.0f}" for b, x in per.items())
            L.append(f"     {str(k)[:18] or '(空白)':<18} Σ={tot:>8,.0f}萬  | {ps}")
    _w()


def _w():
    out = ROOT / "results" / "inspect_vml_div1.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
