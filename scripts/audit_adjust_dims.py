r"""audit_adjust_dims.py — 調整前/調整/調整後 + 調整一級/二級 有冇真係食入 Tableau
Run: python scripts\audit_adjust_dims.py
In : tableau_combined_25.csv
Out: results\audit_adjust_dims.txt
每家 × bucket：三個調整數 Σ、一致性(後=前+調整)、Δ(調整後−amount_mop)、
lv1/lv2 覆蓋率（尤其 調整≠0 但 lv1 空 = 有調整冇分類）、MGM 詳細 by 調整一級。
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
OUT  = ROOT / "results" / "audit_adjust_dims.txt"
CSV  = next((c for c in [ROOT/"tableau_combined_25.csv", Path("tableau_combined_25.csv")] if c.exists()), None)
COLS = ["entity", "year_bucket", "amount_mop", "調整前_萬", "調整_萬", "調整後_萬", "調整一級", "調整二級"]
ENTS = ["galaxy", "sjm", "wynn", "vml", "melco", "mgm"]
BKS  = ["25", "25_24SY", "25_23SY", "24", "24_23SY", "23"]


def _blank(s):
    return s.fillna("").astype(str).str.strip().isin(["", "nan", "None", "NaN", "<NA>", "0", "0.0"])


def main():
    if CSV is None:
        OUT.parent.mkdir(exist_ok=True); OUT.write_text("!! csv 揾唔到", encoding="utf-8"); print("no csv"); return
    parts, mgm_parts = [], []
    for ch in pd.read_csv(CSV, usecols=lambda c: c in COLS, chunksize=400_000, dtype=str, encoding="utf-8-sig"):
        ch = ch.rename(columns=lambda c: c.strip())
        ch["amt"]  = pd.to_numeric(ch["amount_mop"], errors="coerce").fillna(0.0) / 1e4
        for c in ("調整前_萬", "調整_萬", "調整後_萬"):
            ch[c] = pd.to_numeric(ch.get(c), errors="coerce").fillna(0.0)
        ch["_lv1"] = (~_blank(ch.get("調整一級", pd.Series("", index=ch.index)))).astype(int)
        ch["_lv2"] = (~_blank(ch.get("調整二級", pd.Series("", index=ch.index)))).astype(int)
        ch["_adjnz"] = (ch["調整_萬"].abs() > 0.005).astype(int)
        ch["_adjnz_lv1miss"] = ((ch["_adjnz"] == 1) & (ch["_lv1"] == 0)).astype(int)
        ch["rows"] = 1
        g = ch.groupby(["entity", "year_bucket"], as_index=False).agg(
            amt=("amt","sum"), pre=("調整前_萬","sum"), adj=("調整_萬","sum"), post=("調整後_萬","sum"),
            rows=("rows","sum"), adjnz=("_adjnz","sum"), lv1=("_lv1","sum"), lv2=("_lv2","sum"),
            lv1miss=("_adjnz_lv1miss","sum"))
        parts.append(g)
        m = ch[(ch["entity"] == "mgm") & (ch["_adjnz"] == 1)]
        if len(m):
            mgm_parts.append(m.groupby(["year_bucket", "調整一級"], dropna=False, as_index=False)
                              .agg(adj=("調整_萬","sum"), rows=("rows","sum")))
    G = pd.concat(parts, ignore_index=True).groupby(["entity","year_bucket"], as_index=False).sum()

    L = ["# audit_adjust_dims — 調整維度落 Tableau 檢查（萬MOP）", ""]
    for e in ENTS:
        ge = G[G.entity == e]
        if not len(ge): continue
        L += ["", "█" * 76, f"█ {e.upper()}",
              f"   {'bucket':<9}{'amount_mop':>12}{'調整前':>11}{'調整':>10}{'調整後':>11}"
              f"{'後-(前+調)':>10}{'後-amt':>9}{'調整≠0行':>9}{'lv1有':>8}{'lv2有':>8}{'⚠調整≠0冇lv1':>12}"]
        for b in BKS:
            r = ge[ge.year_bucket == b]
            if not len(r): continue
            r = r.iloc[0]
            eq  = r["post"] - (r["pre"] + r["adj"])
            da  = r["post"] - r["amt"]
            L.append(f"   {b:<9}{r['amt']:>12,.0f}{r['pre']:>11,.0f}{r['adj']:>10,.0f}{r['post']:>11,.0f}"
                     f"{eq:>10,.0f}{da:>9,.0f}{int(r['adjnz']):>9,}{int(r['lv1']):>8,}{int(r['lv2']):>8,}"
                     f"{int(r['lv1miss']):>12,}")
        L.append(f"   {'—合計—':<9}{ge['amt'].sum():>12,.0f}{ge['pre'].sum():>11,.0f}{ge['adj'].sum():>10,.0f}"
                 f"{ge['post'].sum():>11,.0f}{'':>10}{'':>9}{int(ge['adjnz'].sum()):>9,}"
                 f"{int(ge['lv1'].sum()):>8,}{int(ge['lv2'].sum()):>8,}{int(ge['lv1miss'].sum()):>12,}")

    if mgm_parts:
        M = pd.concat(mgm_parts, ignore_index=True).groupby(["year_bucket","調整一級"], dropna=False, as_index=False).sum()
        L += ["", "█" * 76, "█ MGM 詳細：bucket × 調整一級 × Σ調整（調整≠0 行）"]
        for b in BKS:
            sub = M[M.year_bucket == b].sort_values("adj")
            if not len(sub): continue
            L.append(f"   ── {b} ──")
            for _, r in sub.iterrows():
                lv = str(r["調整一級"]).strip() or "<空>"
                L.append(f"      {lv[:44]:<46} {r['adj']:>10,.0f}萬 {int(r['rows']):>7,}行")

    L += ["", "解讀：",
          "  後-(前+調) 應=0（三數自洽）；後-amt = 調整後 vs Tableau tie 金額差（wynn 25 已知 +110 accepted）。",
          "  ⚠調整≠0冇lv1 > 0 = 有調整但冇一級分類 → ADJUST_MAP lv1 冇 cover 嗰批行。"]
    OUT.parent.mkdir(exist_ok=True); OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {OUT.relative_to(ROOT)}")

if __name__ == "__main__":
    main()
