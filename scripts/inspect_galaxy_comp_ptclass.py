r"""Galaxy comp 分類欄位全掃：
  1. 項目組H 含 'Comp' 嘅所有唯一值 + 金額（bucket 24 prefix 調整後，25 prefix 調整前）
  2. 調整一級 / 調整二級 值分佈（comp H 行）
  3. 'Comp-' vs 'Comp|' 金額對比
  目的：找出 galaxy 24 pt_class_H filter "Comp|" = 54,991 vs "Comp-" = 92,516 嘅差異

Run（Windows）:
  python scripts\inspect_galaxy_comp_ptclass.py
出 results\inspect_galaxy_comp_ptclass.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
COMP_H = {"H_HOTEL_ROOM", "H_VENUE", "H_FNB", "H_COMP_TICKET", "H_COMP_OTHER"}


def _amt(sub, yr):
    post = pd.to_numeric(sub.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    pre  = pd.to_numeric(sub.get("調整前_萬",  0), errors="coerce").fillna(0.0)
    yb   = sub["year_bucket"].astype(str)
    y2   = yb.str[:2]
    return post.where(y2.isin(["23", "24"]), pre)


def main():
    L = ["# inspect_galaxy_comp_ptclass — galaxy 項目組H 'Comp*' 分佈 + 調整欄位"]
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    gl = df[df["entity"].astype(str) == "galaxy"].copy()

    ph = gl["項目組H"].astype(str) if "項目組H" in gl.columns else pd.Series("", index=gl.index)
    hid = gl["horizontal_id"].astype(str).str.strip() if "horizontal_id" in gl.columns else pd.Series("", index=gl.index)
    yb = gl["year_bucket"].astype(str)
    y2 = yb.str[:2]

    # ── A: 項目組H 唯一值（含 Comp）per bucket ──
    for yr in ("23", "24", "25"):
        if yr == "23":
            sub = gl[yb == "23"].copy()
        else:
            sub = gl[y2 == yr].copy()
        sub["m"] = _amt(sub, yr)
        sub_ph = ph.reindex(sub.index)

        # 含 Comp 的值
        comp_ph = sub[sub_ph.str.contains("Comp", case=False, na=False)]
        L.append(f"\n{'='*64}\n## galaxy bucket {yr} — 項目組H 含 'Comp' 值 (前20 by |Σ|)")
        if len(comp_ph):
            g = comp_ph.groupby(sub_ph.reindex(comp_ph.index))["m"].sum().reset_index()
            g.columns = ["項目組H", "Σ"]
            g = g.sort_values("Σ", key=lambda s: s.abs(), ascending=False)
            for _, r in g.head(20).iterrows():
                L.append(f"  {str(r['項目組H'])[:50]:<50} {r['Σ']:>10,.0f}萬")
            # 分別統計 "Comp|" vs "Comp-" vs 其他
            pipe_mask = sub_ph.reindex(comp_ph.index).str.contains(r"Comp\|", regex=True, na=False)
            dash_mask = sub_ph.reindex(comp_ph.index).str.contains(r"Comp-|Comp –|Comp —", regex=True, na=False)
            others_mask = ~pipe_mask & ~dash_mask
            L.append(f"\n  summary:")
            L.append(f"    'Comp|' filter  rows={int(pipe_mask.sum()):,}  Σ={comp_ph.loc[pipe_mask,'m'].sum():,.0f}萬")
            L.append(f"    'Comp-' filter  rows={int(dash_mask.sum()):,}  Σ={comp_ph.loc[dash_mask,'m'].sum():,.0f}萬")
            L.append(f"    其他 Comp*      rows={int(others_mask.sum()):,}  Σ={comp_ph.loc[others_mask,'m'].sum():,.0f}萬")
        else:
            L.append("  (無 項目組H 含 Comp 行)")

        # ── B: 現 comp H 行的 調整一級 / 調整二級 分佈 ──
        comp_h = sub[hid.reindex(sub.index, fill_value="").isin(COMP_H)]
        L.append(f"\n## galaxy bucket {yr} comp H — 調整一級/二級 分佈")
        for col in ("調整一級", "調整二級", "adjust_lv1", "adjust_lv2"):
            if col in comp_h.columns:
                g2 = comp_h.groupby(comp_h[col].astype(str))["m"].sum().reset_index()
                g2 = g2[g2["m"].abs() >= 100].sort_values("m", key=lambda s: s.abs(), ascending=False)
                if len(g2):
                    L.append(f"  [{col}]")
                    for _, r in g2.head(10).iterrows():
                        L.append(f"    {str(r[col])[:40]:<40} {r['m']:>10,.0f}萬")

    # ── C: 檢查有冇 "基礎 一級" / "基礎 二級" 之類的欄位 ──
    L.append(f"\n{'='*64}\n## galaxy 所有含 'Comp' 嘅欄名 (column scan)")
    for col in gl.columns:
        unique_vals = gl[col].astype(str).unique()
        comp_vals = [v for v in unique_vals if "comp" in v.lower() and len(v) < 80]
        if comp_vals:
            L.append(f"  col={col}: {comp_vals[:6]}")

    out = ROOT / "results" / "inspect_galaxy_comp_ptclass.txt"
    out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
