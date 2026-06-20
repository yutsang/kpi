r"""V 紅旗 inspect：
 A. sjm project/subproject 名質素（附件3.pdf 等垃圾名；睇係咪成個 sjm 都錯）
 B. 藝術珍寶博物館 點解拆兩個 V（capex/opex? entity?）
 C. 路演 入面嘅 演唱會/水上樂園 係邊個 entity（user 去查）

Run（Windows）:  python scripts\inspect_v_redflags.py
出 results\inspect_v_redflags.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
L = ["# inspect_v_redflags"]


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    if not CSV.exists():
        print("!! csv 揾唔到"); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    df["m"] = pd.to_numeric(df.get("amount_mop", 0), errors="coerce").fillna(0.0) / 1e4
    for c in ("entity", "vertical_label", "subproject", "project", "ng_code", "final_capex_opex", "account_desc", "project_full"):
        df[c] = _s(df[c]) if c in df.columns else ""

    # ── A. sjm 名質素 ──
    sj = df[df.entity == "sjm"]
    L.append(f"\n{'='*66}\n## A. sjm project/subproject 名（Σ={sj['m'].sum():,.0f}萬）")
    L.append("  ── top subproject by amount（睇有冇垃圾名）──")
    g = sj.groupby("subproject")["m"].sum().reset_index().sort_values("m", key=lambda s: s.abs(), ascending=False)
    for _, r in g.head(25).iterrows():
        bad = " ⚠垃圾?" if r["subproject"].lower().endswith(".pdf") or "附件" in r["subproject"] or r["subproject"].strip() in ("", "-", "0") or r["subproject"].replace(".", "").isdigit() else ""
        L.append(f"     {r['subproject'][:50]:<50} {r['m']:>9,.0f}{bad}")
    # garbage 合計
    badm = sj["subproject"].str.lower().str.endswith(".pdf") | sj["subproject"].str.contains("附件", na=False)
    L.append(f"  ── 垃圾名(.pdf/附件) Σ={sj.loc[badm,'m'].sum():,.0f}萬 / {int(badm.sum())}行 ──")
    L.append(f"  ── sjm 有幾多 distinct subproject 名：{sj['subproject'].nunique():,} ──")
    L.append("  ── sjm 同 project 欄比較（睇 project_full / project 有冇更好名）──")
    for col in ("project", "project_full"):
        if col in sj.columns and sj[col].astype(bool).any():
            gg = sj.groupby(col)["m"].sum().reset_index().sort_values("m", key=lambda s: s.abs(), ascending=False)
            L.append(f"     [{col}] top 8:")
            for _, r in gg.head(8).iterrows():
                L.append(f"        {str(r[col])[:46]:<46} {r['m']:>9,.0f}")

    # ── B. 藝術珍寶博物館 ──
    art = df[df.subproject.str.contains("藝術珍寶博物館|珍寶博物館", na=False)]
    L.append(f"\n{'='*66}\n## B. 『藝術珍寶博物館』Σ={art['m'].sum():,.0f}萬 — 點解拆 V？")
    gb = art.groupby(["entity", "vertical_label", "final_capex_opex", "ng_code"])["m"].sum().reset_index().sort_values("m", key=lambda s: s.abs(), ascending=False)
    for _, r in gb.iterrows():
        if abs(r["m"]) >= 10:
            L.append(f"     {r['entity']:<7} V={r['vertical_label'][:14]:<14} {r['final_capex_opex'][:6]:<6} {r['ng_code']:<5} {r['m']:>9,.0f}")

    # ── C. 路演 入面 演唱會/水上樂園 ──
    rs = df[(df.vertical_label == "路演") & df.subproject.str.contains("演唱會|水上樂園", na=False)]
    L.append(f"\n{'='*66}\n## C. 路演 入面 演唱會/水上樂園 Σ={rs['m'].sum():,.0f}萬 — 邊個 entity？")
    gc = rs.groupby(["entity", "subproject", "ng_code", "final_capex_opex"])["m"].sum().reset_index().sort_values("m", key=lambda s: s.abs(), ascending=False)
    for _, r in gc.head(15).iterrows():
        L.append(f"     {r['entity']:<7} {r['ng_code']:<5} {r['final_capex_opex'][:6]:<6} {r['subproject'][:40]:<40} {r['m']:>8,.0f}")

    out = ROOT / "results" / "inspect_v_redflags.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
