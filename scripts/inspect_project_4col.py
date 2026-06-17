r"""Check the 4-layer project-name structure in the Tableau combined output.

項目組要求 4(+1) 欄齊：
  dicj_code     golden DICJ 碼
  項目名稱       golden DICJ 名（roll-up 層 — 最粗）
  project_code  我哋統一碼（subproject 碼）
  project       我哋 project 名（應該比 dicj 名細少少）
  subproject    仲細一層（如 entity 有 sub 欄）

Run after prep_tableau:  python scripts/inspect_project_4col.py
讀 tableau_combined_25.csv → per entity 覆蓋率 + 抽樣睇 dicj名 vs project名 粒度。
出 results/inspect_project_4col.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
# prep_tableau 已改清晰名：dicj code / project(=DICJ大項目名) / subproject code(細碼) / subproject(細名)
COLS = ["dicj code", "project", "subproject code", "subproject"]


def _pop(s):
    s = s.astype(str).str.strip().replace({"nan": "", "None": ""})
    return s, (s != "").mean() * 100


def main():
    L = ["# inspect_project_4col — 4-層 project 名結構 (dicj碼 / dicj名 / project碼 / project名 / subproject)"]
    if not CSV.exists():
        L.append(f"!! {CSV.name} 揾唔到 — 先跑  python scripts/prep_tableau.py"); _w(L); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    L.append(f"   rows={len(df):,}   欄齊？ " + "  ".join(f"{c}{'✓' if c in df.columns else '✗ 冇'}" for c in COLS))
    ecol = "entity" if "entity" in df.columns else None
    ents = sorted(df[ecol].unique()) if ecol else ["(all)"]
    for ent in ents:
        d = df[df[ecol] == ent] if ecol else df
        L.append(f"\n── {ent}  ({len(d):,} 行) ──")
        for c in COLS:
            if c in d.columns:
                s, pop = _pop(d[c])
                L.append(f"     {c:<16} 有值 {pop:5.1f}%   distinct={s[s != ''].nunique():,}")
            else:
                L.append(f"     {c:<16} ✗ 冇呢欄")
        # granularity: per 'dicj code' 有幾多 distinct 'subproject code' (>1 = 細碼真係比 dicj 細)
        if "dicj code" in d.columns and "subproject code" in d.columns:
            dd = d[d["dicj code"].astype(str).str.strip().replace({"nan": "", "None": ""}) != ""]
            g = dd.groupby("dicj code")["subproject code"].nunique()
            if len(g):
                L.append(f"     granularity: {int((g > 1).sum())}/{len(g)} 個 dicj code 有 >1 個 subproject code"
                         f"（即細碼比 dicj 細；最多一個 dicj 拆 {int(g.max())} 個 subproject code）")
        # 抽樣 8 個 distinct，睇 4 層粒度
        scols = [c for c in ["dicj code", "project", "subproject code", "subproject"] if c in d.columns]
        samp = d[scols].drop_duplicates().head(8)
        L.append(f"     抽樣 ({' | '.join(scols)}):")
        for _, r in samp.iterrows():
            L.append("       " + "  |  ".join(str(r[c])[:24] for c in scols))
    _w(L)


def _w(L):
    out = ROOT / "results" / "inspect_project_4col.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
