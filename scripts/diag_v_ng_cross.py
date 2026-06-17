r"""④ NG×V nonsense — 揾 V 掛咗去唔屬於佢「本家 NG」嘅 combo（user: 美食活動×會議展覽 = 亂）。

V 同 NG 係兩種 taxonomy，但有啲 V 有明確本家 NG（V_FOOD_EVENT 本家 NG8 美食；V_MICE 本家 NG2 會展…）。
本家 NG ≠ 行 NG = cross-NG 可疑。cross-cutting V（物業優化等）任何 NG 都合理，唔 flag。
場館 V（表演/體育/會展場館）只合理 NG2/3/4。

Run:  python scripts/diag_v_ng_cross.py
出 results/diag_v_ng_cross.txt  ← paste 返嚟（按金額排，逐個 combo 決定要唔要 reset）
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"

HOME_NG = {
    "V_GAMING_VENUE": "NG0", "V_GAMING_EQUIP": "NG0",
    "V_OVERSEAS_ROADSHOW": "NG1", "V_OVERSEAS_OFFICE": "NG1", "V_REGIONAL_TEAM": "NG1",
    "V_REGIONAL_SALES": "NG1", "V_INVITE_GUEST": "NG1", "V_INVITE_AGENCY": "NG1",
    "V_OVERSEAS_WEB_SEO": "NG1", "V_PROMO_VIDEO": "NG1",
    "V_MICE": "NG2", "V_CONCERT": "NG3", "V_SPORT_EVENT": "NG4",
    "V_ART_EXHIBITION": "NG5", "V_MUSEUM": "NG5", "V_WELLNESS": "NG6", "V_THEME_PARK": "NG7",
    "V_RESTAURANT": "NG8", "V_FOOD_EVENT": "NG8", "V_COMMUNITY": "NG9", "V_MARITIME": "NG10",
    "V_OTHER": "NG11", "V_FESTIVAL": "NG11", "V_PUBLIC_FACILITY": "NG11",
}
# cross-cutting：本質上任何 NG 都會有，唔當 nonsense（設施/物業/節慶/社區/宣傳片）
CROSS = {"V_PROPERTY_UPGRADE", "V_PUBLIC_FACILITY", "V_FESTIVAL", "V_COMMUNITY", "V_PROMO_VIDEO"}
VENUE = {"V_VENUE_PERF_SPORT_MICE"}           # 只合理 NG2/3/4
VENUE_OK = {"NG2", "NG3", "NG4"}


def _is_cross(vid, ng):
    if not vid or vid in CROSS:
        return False
    if vid in VENUE:
        return ng not in VENUE_OK
    home = HOME_NG.get(vid)
    return bool(home) and ng != home


def main():
    L = ["# diag_v_ng_cross — V 掛錯 NG (本家 NG ≠ 行 NG)，按金額排"]
    if not CSV.exists():
        L.append("!! tableau_combined_25.csv 揾唔到"); _w(L); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    df["amt"] = pd.to_numeric(df["amount_mop"], errors="coerce").fillna(0.0).abs()
    vid = df["vertical_id"].astype(str).str.strip()
    ng = df["ng_code"].astype(str).str.strip()
    flag = [_is_cross(v, n) for v, n in zip(vid.tolist(), ng.tolist())]
    bad = df[pd.Series(flag, index=df.index)]
    L.append(f"   cross-NG 行: {len(bad):,}  Σ={bad['amt'].sum()/1e6:,.1f}M（總 {df['amt'].sum()/1e6:,.0f}M 嘅 {bad['amt'].sum()/df['amt'].sum()*100:.1f}%）")
    g = (bad.groupby(["entity", "ng_code", "ng_label", "vertical_id", "vertical_label"])["amt"]
            .sum().reset_index().sort_values("amt", ascending=False))
    L.append("\n   ── 掛錯 combo（按金額，top 40）：entity | 行NG | V(本家NG) | Σ萬 ──")
    for r in g.head(40).itertuples():
        home = HOME_NG.get(r.vertical_id, "VENUE" if r.vertical_id in VENUE else "?")
        L.append(f"     {r.entity:<7} {r.ng_code:<5}{str(r.ng_label)[:6]:<6} ← {str(r.vertical_label)[:10]:<10}"
                 f"(本家{home})  {r.amt/1e4:>10,.0f}萬")
    # per entity 小結
    L.append("\n   ── per entity cross-NG Σ ──")
    for e, sub in bad.groupby("entity"):
        L.append(f"     {e:<7} {len(sub):>7,} 行  Σ={sub['amt'].sum()/1e6:>8,.1f}M")
    _w(L)


def _w(L):
    out = ROOT / "results" / "diag_v_ng_cross.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
