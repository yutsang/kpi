r"""inspect_golden_ng_names —— 睇 golden 有冇 per-(dicj,NG) 項目名（決定問題2 NG-aware 名 fill 改唔改到）。

問題2：項目名 ← golden dicj-idxmax（NG-blind）。melco coarse dicj（項目13 跨 NG1/3/7）→ 全標「水上樂園」。
查：golden「Database combine」sheet
  1. 全部欄名 + 有冇 NG/項目性質 欄
  2. melco 跨NG dicj（項目12/13/18/19/32/42…）：dicj ×（項目性質 if 有）× 項目名稱 × Amount
     → 若同一 dicj 唔同 NG 有唔同名 = golden 有 per-NG 名，NG-aware fill 改得到
     → 若同一 dicj 得一個名 = golden 只有 dicj 名，要靠 databook granular / 唔同源
另：CSV 入面 melco 項目13 嘅 subproject/科目明細 等，睇有冇 row-level granular 名。

Run（Windows）:  python scripts\inspect_golden_ng_names.py
出 results\inspect_golden_ng_names.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "results" / "inspect_golden_ng_names.txt"
CSV = ROOT / "tableau_combined_25.csv"

# golden 候選路徑（同 prep_tableau _GCAND 一致）
GCAND = [
    ROOT / "HQ 投資方向_audit_0616.xlsx",
    ROOT / "data" / "HQ 投資方向_audit_0616.xlsx",
]
SPAN_DICJ = ["項目12", "項目13", "項目18", "項目19", "項目14", "項目32", "項目42", "項目16", "項目17", "項目28"]


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    L = ["# inspect_golden_ng_names —— golden 有冇 per-(dicj,NG) 名"]

    # ── golden ────────────────────────────────────────────────────────────────
    gp = next((p for p in GCAND if p.exists()), None)
    if not gp:
        # rglob 揾
        hits = list(ROOT.rglob("HQ 投資方向*.xlsx"))
        gp = hits[0] if hits else None
    if not gp:
        L.append("  ⚠ golden 檔揾唔到（試咗 data/golden, root/, ./）。請話我 golden 檔路徑。")
    else:
        L.append(f"  golden = {gp}")
        try:
            xl = pd.ExcelFile(gp)
            L.append(f"  sheets: {xl.sheet_names}")
            g = pd.read_excel(gp, sheet_name="Database combine", dtype=str)
            g.columns = [str(c).strip() for c in g.columns]
            L.append(f"\n  ── Database combine 欄名({len(g.columns)}) ──\n  {list(g.columns)}")
            ngcol = next((c for c in g.columns if ("投資範疇" in c and "博彩" not in c) or "項目性質" in c or "Section" in c or c.strip() in ("NG", "ng_code", "NG Code")), None)
            dcol = next((c for c in g.columns if c.strip() in ("DICJ Code", "DICJ")), None)
            ncol = next((c for c in g.columns if "項目名稱" in str(c) or "项目名称" in str(c)), None)
            acomp = next((c for c in g.columns if "承批" in str(c)), None)
            L.append(f"\n  偵測：項目性質/NG欄={ngcol!r} ｜ DICJ欄={dcol!r} ｜ 項目名欄={ncol!r} ｜ 承批欄={acomp!r}")
            if ngcol:
                L.append("  ✅ golden 有 NG/項目性質 欄 → NG-aware (dicj,NG)→名 fill 改得到")
            else:
                L.append("  ❌ golden 冇 NG 欄 → 一個 dicj 得一個名，要靠 databook granular / subproject")

            # melco 跨NG dicj 樣本
            if dcol and ncol and acomp:
                gm = g[g[acomp].astype(str).str.contains("新濠|Melco|MELCO|melco", na=False)].copy()
                gm["_d"] = gm[dcol].astype(str).str.strip()
                acol = next((c for c in g.columns if c.strip() == "Amount"), None)
                gm["_amt"] = pd.to_numeric(gm[acol].astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0.0) if acol else 0.0
                for d in SPAN_DICJ:
                    sub = gm[gm["_d"].astype(str).str.contains(d, na=False)]
                    if sub.empty:
                        continue
                    L.append(f"\n  ── golden melco {d} ──")
                    cols = [c for c in [ngcol, ncol] if c]
                    gb = sub.groupby([gm.loc[sub.index, c] for c in cols] if cols else sub.index)["_amt"].sum() if cols else None
                    grp = sub.groupby(cols)["_amt"].sum().reset_index().sort_values("_amt", key=lambda s: s.abs(), ascending=False) if cols else sub
                    for _, r in grp.head(8).iterrows():
                        vals = " | ".join(f"{c}={str(r[c])[:24]}" for c in cols)
                        L.append(f"      {vals}  Σ={r['_amt']:,.0f}")
        except Exception as e:
            L.append(f"  ⚠ 讀 golden 失敗: {e}")

    # ── CSV：melco 項目13 row-level granular 名候選 ────────────────────────────
    if CSV.exists():
        df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
        m = df[(df["entity"].astype(str).eq("melco")) & (df.get("dicj code", pd.Series("", index=df.index)).astype(str).str.contains("項目13", na=False))]
        L.append(f"\n  ── CSV melco dicj=項目13：{len(m):,}行，by ng_code × subproject（睇有冇 granular 名）──")
        if len(m):
            for c in ("ng_code", "ng_label", "subproject", "project", "項目組V", "科目明細"):
                if c not in m.columns:
                    continue
            g2 = m.groupby([m["ng_code"].astype(str), m["subproject"].astype(str)])["調整前_萬"].apply(
                lambda s: pd.to_numeric(s, errors="coerce").fillna(0).sum()).reset_index()
            g2 = g2.sort_values("調整前_萬", key=lambda s: s.abs(), ascending=False)
            for _, r in g2.head(15).iterrows():
                L.append(f"      {r['ng_code']:<6} sub={str(r['subproject'])[:36]:<36} {r['調整前_萬']:>9,.0f}萬")

    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
