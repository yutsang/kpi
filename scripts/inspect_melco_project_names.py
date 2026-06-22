r"""inspect_melco_project_names —— 揾 melco 跨NG dicj 嘅 per-(NG) granular 名喺邊條欄。

問題：項目名 ← golden dicj-idxmax（NG-blind），melco 項目13 全標「水上樂園」。
user 話 database 有寫清楚 (NG,code)→名，係我哋回填錯。
查：melco 跨NG dicj，per (dicj_code, ng_code) 睇各 name 候選欄（project/subproject/project_full/項目組V）
   邊條有真 per-NG granular 名 → 就用嗰條取代 golden idxmax。

Run（Windows）:  python scripts\inspect_melco_project_names.py
出 results\inspect_melco_project_names.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
OUT = ROOT / "results" / "inspect_melco_project_names.txt"
DICJS = ["項目12", "項目13", "項目16", "項目17", "項目18", "項目19", "項目28"]


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    L = ["# inspect_melco_project_names —— melco 跨NG dicj per-(NG) 名候選"]
    if not CSV.exists():
        L.append("!! csv 揾唔到"); _w(L); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    df["pre"] = pd.to_numeric(df.get("調整前_萬", 0), errors="coerce").fillna(0.0)
    m = df[df["entity"].astype(str).eq("melco")].copy()
    for c in ("dicj code", "ng_code", "ng_label", "project", "subproject", "project_full", "項目組V", "account_desc"):
        m[c] = _s(m.get(c, pd.Series("", index=m.index)))

    L.append(f"\n  melco 全部欄: {[c for c in df.columns]}")
    for d in DICJS:
        sub = m[m["dicj code"].str.contains(d, na=False)]
        if sub.empty:
            continue
        L.append(f"\n{'='*72}\n## dicj={d}  (Σpre={sub['pre'].sum():,.0f}萬, {len(sub):,}行)")
        for ng, g in sub.groupby("ng_code"):
            L.append(f"\n  ── {ng} ({g['ng_label'].iloc[0]}) Σ={g['pre'].sum():,.0f}萬 ──")
            for col in ("project", "subproject", "project_full", "項目組V"):
                vals = g.groupby(col)["pre"].sum().sort_values(key=lambda s: s.abs(), ascending=False)
                vals = vals[vals.index.astype(str).str.strip().ne("")]
                top = " | ".join(f"{str(k)[:26]}({v:,.0f})" for k, v in vals.head(3).items())
                L.append(f"     {col:<12}: {top if top else '(全空)'}")
    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
