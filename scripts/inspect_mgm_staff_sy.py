r"""MGM 24_23SY bucket 嘅 H_LABOR 行拆細 — 理解點解 2,437萬喺 SY rows。

Run（Windows）:  python scripts\inspect_mgm_staff_sy.py
出 results\inspect_mgm_staff_sy.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
OUT = ROOT / "results" / "inspect_mgm_staff_sy.txt"


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    L = ["# inspect_mgm_staff_sy — MGM 24_23SY 人工 rows 拆細（萬）"]
    if not CSV.exists():
        L.append("!! tableau_combined_25.csv 揾唔到"); _w(L); return

    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    e = df[df["entity"].astype(str).eq("mgm")].copy()

    post = pd.to_numeric(e.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    pre  = pd.to_numeric(e.get("調整前_萬",  0), errors="coerce").fillna(0.0)
    yb   = e["year_bucket"].astype(str)
    y2   = yb.str[:2]
    e["m_post"] = post
    e["m_pre"]  = pre
    e["m_pref"] = post.where(y2.isin(["23", "24"]), pre)
    e["yb"] = yb
    e["y2"] = y2
    for c in ("horizontal_id", "final_capex_opex", "account_code", "account_desc", "description"):
        e[c] = _s(e[c]) if c in e.columns else ""
    e["code"] = e["account_code"].str.replace(r"\.0$", "", regex=True)
    opex = ~e["final_capex_opex"].eq("Capex")
    lab  = e[e["horizontal_id"].eq("H_LABOR") & opex]

    # ── 全 prefix 24 總覽 ──────────────────────────────────────────────────────
    p24 = lab[lab.y2.eq("24")]
    L.append(f"\n{'='*64}\n## MGM 24-prefix 人工 Σ={p24['m_pref'].sum():,.0f}萬  HQ=6,829  Δ={p24['m_pref'].sum()-6829:+,.0f}  ({len(p24):,}行)")
    bg = p24.groupby("yb")["m_pref"].sum().reset_index().sort_values("yb")
    for _, r in bg.iterrows():
        L.append(f"   {r['yb']:<16} {r['m_pref']:>9,.0f}萬")

    # ── 24_23SY 詳細 ──────────────────────────────────────────────────────────
    sy = lab[lab.yb.eq("24_23SY")]
    L.append(f"\n{'='*64}\n## MGM 24_23SY bucket  Σ={sy['m_pref'].sum():,.0f}萬  ({len(sy):,}行)")
    if sy.empty:
        L.append("  (empty)")
    else:
        L.append("\n  ── by account_code × account_desc (全部) ──")
        g = sy.groupby(["code", "account_desc"])["m_pref"].sum().reset_index().sort_values("m_pref", key=lambda s: s.abs(), ascending=False)
        for _, r in g.iterrows():
            L.append(f"   {r['code']:<12}{r['account_desc'][:35]:<35} {r['m_pref']:>9,.0f}萬")

        # description keywords to understand what these SY rows represent
        if sy["description"].str.strip().ne("").any():
            L.append("\n  ── top 10 description ──")
            gd = sy.groupby("description")["m_pref"].sum().sort_values(ascending=False).head(10)
            for d, amt in gd.items():
                L.append(f"   {str(d)[:60]:<60}  {amt:>8,.0f}萬")

        # 調整後 vs 調整前 比較（SY rows 係 post 定 pre 主導？）
        L.append(f"\n  ── SY rows: Σ 調整後={sy['m_post'].sum():,.0f}萬  Σ 調整前={sy['m_pre'].sum():,.0f}萬 ──")
        L.append(f"     m_pref 用 調整後（y2=24）= {sy['m_pref'].sum():,.0f}萬")

    # ── exact 24 詳細 ──────────────────────────────────────────────────────────
    ex24 = lab[lab.yb.eq("24")]
    L.append(f"\n{'='*64}\n## MGM exact '24' bucket  Σ={ex24['m_pref'].sum():,.0f}萬  ({len(ex24):,}行)")
    L.append(f"   vs HQ=6,829  Δ={ex24['m_pref'].sum()-6829:+,.0f}萬（如 HQ 係 exact 24 basis）")
    g2 = ex24.groupby(["code", "account_desc"])["m_pref"].sum().reset_index().sort_values("m_pref", key=lambda s: s.abs(), ascending=False)
    for _, r in g2.head(15).iterrows():
        L.append(f"   {r['code']:<12}{r['account_desc'][:35]:<35} {r['m_pref']:>9,.0f}萬")

    # ── 25-prefix 同樣拆法 ─────────────────────────────────────────────────────
    p25 = lab[lab.y2.eq("25")]
    L.append(f"\n{'='*64}\n## MGM 25-prefix 人工 Σ={p25['m_pref'].sum():,.0f}萬  HQ=21,467  Δ={p25['m_pref'].sum()-21467:+,.0f}  ({len(p25):,}行)")
    bg25 = p25.groupby("yb")["m_pref"].sum().reset_index().sort_values("yb")
    for _, r in bg25.iterrows():
        L.append(f"   {r['yb']:<16} {r['m_pref']:>9,.0f}萬")

    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
