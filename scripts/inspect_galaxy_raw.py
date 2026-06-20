r"""直接拆 galaxy raw `data\galaxy\raw\project_team\galaxy_2025.xlsx` tab `Combine(clean)`。
目的：搵 房/票 comp 喺邊個欄 filter 到 golden 值（user 話肉眼睇到）。
dump 所有欄 + 房 comp 行嘅每個 amount 欄 sum + 分類欄 value 分佈。

Run（Windows）:  python scripts\inspect_galaxy_raw.py
出 results\inspect_galaxy_raw.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
F = ROOT / "data" / "galaxy" / "raw" / "project_team" / "galaxy_2025.xlsx"
SHEET = "Combine(clean)"
ROOM = r"room|客房|Lodging|Hotel|住宿|房"
L: list[str] = [f"# inspect_galaxy_raw — {F.name} [{SHEET}]"]


def main():
    if not F.exists():
        L.append(f"!! 揾唔到 {F}"); _w(); return
    df = pd.read_excel(F, sheet_name=SHEET, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    L.append(f"\nshape = {df.shape}")
    L.append(f"\n── 所有欄名（{len(df.columns)}）──")
    for i, c in enumerate(df.columns):
        nonblank = df[c].astype(str).str.strip().replace({"nan": "", "None": ""}).astype(bool).sum()
        samp = " | ".join(df[c].dropna().astype(str).str.strip().head(3).tolist())[:70]
        L.append(f"   [{i:>2}] {c[:26]:<26} nonblank={nonblank:>6}  e.g. {samp}")

    # 數值欄（amount 候選）
    numcols = []
    for c in df.columns:
        v = pd.to_numeric(df[c].astype(str).str.replace(",", "", regex=False), errors="coerce")
        if v.notna().sum() >= len(df) * 0.3:
            numcols.append(c)
    L.append(f"\n── 數值欄（amount 候選 {len(numcols)}）總和 ──")
    for c in numcols:
        v = pd.to_numeric(df[c].astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0)
        L.append(f"   {c[:30]:<30} Σ={v.sum():>18,.0f}")

    # 含房 keyword 嘅 text 欄 → 邊個欄有 room 分類
    L.append(f"\n── 邊啲 text 欄含房 keyword（room 分類欄候選）──")
    for c in df.columns:
        if c in numcols:
            continue
        s = df[c].astype(str)
        m = s.str.contains(ROOM, case=False, regex=True, na=False)
        if m.sum() >= 5:
            vc = s[m].value_counts().head(5)
            L.append(f"   [{c}] {m.sum()} 行含房；top值: " + " | ".join(f"{k[:20]}×{v}" for k, v in vc.items()))

    # 揾「room 分類欄」：揀含 room 最多嘅 text 欄，sum 每個數值欄
    cat_col = None
    best = 0
    for c in df.columns:
        if c in numcols:
            continue
        m = df[c].astype(str).str.contains(r"room|客房|Lodging", case=False, regex=True, na=False)
        if m.sum() > best:
            best = m.sum(); cat_col = c
    if cat_col:
        m = df[cat_col].astype(str).str.contains(r"room|客房|Lodging", case=False, regex=True, na=False)
        L.append(f"\n── 用『{cat_col}』含 room 過濾（{int(m.sum())} 行）→ 每個數值欄 sum（搵邊個 = golden 客房 30,188/32,356）──")
        for c in numcols:
            v = pd.to_numeric(df.loc[m, c].astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0)
            L.append(f"     {c[:30]:<30} Σ={v.sum():>16,.0f}")
        L.append(f"\n── 該 room 行 3 條 sample（全欄）──")
        for _, r in df[m].head(3).iterrows():
            L.append("     " + " || ".join(f"{c}={str(r[c])[:18]}" for c in df.columns if str(r[c]).strip() not in ("", "nan", "None")))
    _w()


def _w():
    out = ROOT / "results" / "inspect_galaxy_raw.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
