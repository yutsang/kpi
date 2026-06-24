r"""inspect_mgm_adj —— 揾 mgm 嘅調整喺邊(mgm_2025.xlsx data tab 冇調整欄)。

user 2026-06-24：mgm 用我嘅文件 + 再 inspect。golden mgm 調整後 capex 63,272 vs 報告 155,597 (−92,325 巨額)。
搜 mgm_2025.xlsx + mgm_2025_adj.xlsx 全部 tab 嘅「調整/adjust/期後」欄；+ data tab 報告 by capex/opex；
+ adj 數據簿「調整明細表」類 sheet 睇有冇 capex/opex × 調整 summary。

Run（Windows）:  python scripts\inspect_mgm_adj.py
出 results\inspect_mgm_adj.txt  ← paste 返嚟
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
OUT = ROOT / "results" / "inspect_mgm_adj.txt"
FILES = ["mgm_2025.xlsx", "mgm_2025_adj.xlsx"]
ADJ_KW = ["調整", "adjust", "期後", "amend"]


def _num(s):
    return pd.to_numeric(pd.Series(s).astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0.0)


def _co(x):
    s = str(x).strip().lower()
    return "Capex" if s.startswith("cap") else ("Opex" if s.startswith("op") else "其他/空")


def main():
    L = ["# inspect_mgm_adj —— 揾 mgm 調整喺邊"]
    for fn in FILES:
        p = ROOT / "data" / "mgm" / "raw" / fn
        L.append(f"\n{'='*70}\n## {fn}")
        if not p.exists():
            L.append("   !! 揾唔到"); continue
        try:
            sheets = pd.ExcelFile(p).sheet_names
        except Exception as e:
            L.append(f"   !! 開檔失敗：{e}"); continue
        L.append(f"   sheets({len(sheets)})：{sheets[:50]}")
        # 每張 tab 揾調整欄
        L.append(f"   ── 含「調整/adjust/期後/amend」嘅欄（逐 tab）──")
        hit_any = False
        for sh in sheets:
            try:
                cols = [str(c).strip() for c in pd.read_excel(p, sheet_name=sh, nrows=2, dtype=object).columns]
            except Exception:
                continue
            adjcols = [c for c in cols if any(k.lower() in c.lower() for k in ADJ_KW)]
            if adjcols:
                hit_any = True
                L.append(f"      [{sh}] → {adjcols[:8]}")
        if not hit_any:
            L.append("      （全部 tab 都冇調整欄）")

    # mgm_2025.xlsx data tab：報告 by capex/opex + 全部欄
    p = ROOT / "data" / "mgm" / "raw" / "mgm_2025.xlsx"
    if p.exists():
        try:
            df = pd.read_excel(p, sheet_name="data", dtype=object)
            cols = [str(c).strip() for c in df.columns]
            L.append(f"\n{'='*70}\n## mgm_2025.xlsx [data] 全部欄({len(cols)})：\n   {cols}")
            amtc = next((c for c in ["Debit minus Credit", "25_Amt"] if c in df.columns), None)
            coc = next((c for c in ["Capex_Opex", "Capex/Opex"] if c in df.columns), None)
            if amtc and coc:
                amt = _num(df[amtc]) / 1e4
                co = df[coc].map(_co)
                L.append(f"   ── 報告({amtc}) by capex/opex（萬）──")
                for k in ["Capex", "Opex", "其他/空"]:
                    m = co.eq(k)
                    if int(m.sum()):
                        L.append(f"      {k:<8}{amt[m].sum():>12,.0f}萬  {int(m.sum()):>8,}行")
        except Exception as e:
            L.append(f"   !! data tab 讀失敗：{e}")
    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
