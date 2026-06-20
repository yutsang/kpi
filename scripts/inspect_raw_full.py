r"""通用 raw Excel 拆解器 —— 揾每家 data/{ent}/raw/** 嘅 xlsx，dump 結構搵 comp/staff 解決方法。
逐 file × sheet：欄名+nonblank+sample、數值欄(amount候選)總和、似label欄(comp/room/staff/分類/標簽/類別)value分佈。

Run（Windows）:
  python scripts\inspect_raw_full.py vml
  python scripts\inspect_raw_full.py sjm mgm galaxy
出 results\inspect_raw_full_<ent>.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
ENTS = sys.argv[1:] or ["vml"]
LABEL_KW = "comp|room|客房|會場|venue|food|餐|f&b|staff|人工|薪|salar|payroll|分類|標簽|標籤|類別|類型|category|nature|性質|hotel|房|票|ticket|贈"


def num(s):
    return pd.to_numeric(s.astype(str).str.replace(",", "", regex=False), errors="coerce")


def run(ent, L):
    rawdir = ROOT / "data" / ent / "raw"
    files = sorted(rawdir.rglob("*.xlsx")) if rawdir.exists() else []
    L.append(f"\n{'#'*70}\n# {ent}  raw dir={rawdir}  ({len(files)} xlsx)")
    for f in files:
        if f.name.startswith("~$"):
            continue
        try:
            xl = pd.ExcelFile(f)
        except Exception as e:
            L.append(f"\n!! {f.name} 開唔到: {e}"); continue
        L.append(f"\n{'='*64}\n## {f.relative_to(ROOT)}  sheets={xl.sheet_names}")
        for sh in xl.sheet_names:
            try:
                df = pd.read_excel(f, sheet_name=sh, dtype=str, nrows=200000)
            except Exception as e:
                L.append(f"  [{sh}] 讀唔到: {e}"); continue
            df.columns = [str(c).strip() for c in df.columns]
            if not len(df):
                continue
            L.append(f"\n  ── sheet [{sh}]  {df.shape[0]:,}行 × {df.shape[1]}欄 ──")
            numcols = []
            for c in df.columns:
                v = num(df[c])
                if v.notna().sum() >= len(df) * 0.3:
                    numcols.append(c)
            L.append(f"     欄: " + " | ".join(c[:20] for c in df.columns))
            L.append(f"     數值欄總和：")
            for c in numcols:
                L.append(f"        {c[:28]:<28} Σ={num(df[c]).sum():>18,.0f}")
            # label-like 欄 value 分佈
            for c in df.columns:
                if c in numcols:
                    continue
                lc = c.lower()
                import re
                if re.search(LABEL_KW, lc) or df[c].astype(str).str.contains("[Cc]omp|客房|人工|分類|標簽", regex=True, na=False).mean() > 0.1:
                    vc = df[c].astype(str).str.strip().replace({"nan": "", "None": ""}).value_counts()
                    vc = vc[vc.index != ""]
                    if len(vc):
                        top = " | ".join(f"{k[:18]}×{v}" for k, v in vc.head(8).items())
                        L.append(f"     [label?] {c[:24]:<24} ({df[c].astype(str).str.strip().ne('').sum()}非空) top: {top}")


def main():
    L = ["# inspect_raw_full"]
    for ent in ENTS:
        run(ent, L)
        out = ROOT / "results" / f"inspect_raw_full_{ent}.txt"; out.parent.mkdir(exist_ok=True)
        out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote results/inspect_raw_full_*.txt  ← paste 返嚟")


if __name__ == "__main__":
    main()
