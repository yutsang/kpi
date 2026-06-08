"""SJM 2023 建設與設施支出 — can we split 設備(設施採購) vs 場地翻新(建設) by Description?

建設 245M is dominated by generic 87999990 'Deposits paid - Renovation' rows whose account can't
tell equipment from construction. Esp. V_GAMING_EQUIP (博彩設施設備優化) 97M sits entirely in 建設 —
gaming EQUIPMENT should be 設施及器具採購. This dumps, for the 建設 rows:
  (A) by vertical_label — Σ|amt|
  (B) V_GAMING_EQUIP 建設 rows: top Cost element descr. + top Description free-text (the detail that
      reveals equipment vs renovation) + amounts
  (C) keyword classification on Description: EQUIP-words vs CONSTRUCTION-words — Σ|amt| each + the
      unmatched residual — so we see how cleanly a desc_contains rule could split it
  (D) same EQUIP/CONSTRUCTION keyword tally for ALL 建設 rows (other V too)

Run (Windows):  python scripts/inspect_sjm_23_construction.py
Output: prints + results/inspect_sjm_23_construction.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
TR = ROOT / "data" / "sjm" / "interim" / "company_2_tagged_rows.parquet"
AMT = "Val/COArea Crcy"
AD = "Cost element descr."
DESC = "Description"
EQUIP_KW = ["smart table", "perfect pay", "egm", "slot", "角子機", "shuffler", "gaming table",
            "gaming chair", "machine", "設備", "器材", "器具", "monitor", "監控", "surveillance",
            "iview", "server", "terminal", "kiosk", "感應", "antenna", "cage", "電子", "system",
            "電腦", "computer", "network", "switch", "camera", "鏡頭", "顯示", "lcd", "led", "tv"]
CONSTR_KW = ["翻新", "renovation", "裝修", "工程", "construction", "場地", "fit out", "fit-out",
             "改造", "重修", "結構", "土建", "機電", "mep", "interior", "lighting", "天花", "ceiling",
             "flooring", "地板", "wall", "牆", "拆", "建築", "裝飾", "decoration", "main contract",
             "function room", "lobby", "hall", "大堂", "走廊", "corridor"]


def main():
    L = ["# inspect_sjm_23_construction — split 設備 vs 場地 in 建設 by Description"]
    if not TR.exists():
        L.append(f"X {TR} missing"); _w(L); return
    df = pd.read_parquet(TR)
    per = next((c for c in ("report_period", "report_year", "years") if c in df.columns), None)
    if per:
        df = df[df[per].astype(str).str.startswith("23")].copy()
    a = pd.to_numeric(df[AMT], errors="coerce").fillna(0.0)
    hl = df["horizontal_label"].astype("string").fillna("")
    vl = df["vertical_label"].astype("string").fillna("")
    con = hl.eq("建設與設施支出")
    L.append(f"\n建設 rows={int(con.sum()):,}  Σ|amt|={a.abs()[con].sum():,.0f}")

    # (A) 建設 by vertical
    L.append("\n## (A) 建設 by vertical_label (Σ|amt| M):")
    for v, x in a.abs()[con].groupby(vl[con]).sum().sort_values(ascending=False).items():
        L.append(f"   {str(v)[:24]:24s} {x/1e6:8.2f}M")

    ad = df[AD].astype("string").fillna("").str.strip() if AD in df.columns else pd.Series("", index=df.index)
    desc = df[DESC].astype("string").fillna("").str.strip() if DESC in df.columns else pd.Series("", index=df.index)

    def kw_split(mask, title):
        sub_a = a.abs()[mask]
        d = desc[mask].str.lower()
        eq = pd.Series(False, index=d.index)
        for k in EQUIP_KW: eq |= d.str.contains(k.lower(), na=False, regex=False)
        cn = pd.Series(False, index=d.index)
        for k in CONSTR_KW: cn |= d.str.contains(k.lower(), na=False, regex=False)
        only_eq = eq & ~cn; only_cn = cn & ~eq; both = eq & cn; none = ~eq & ~cn
        L.append(f"\n## {title}: {int(mask.sum()):,} rows  {sub_a.sum()/1e6:.1f}M")
        L.append(f"   EQUIP-only : {sub_a[only_eq].sum()/1e6:8.2f}M  ({int(only_eq.sum())} rows)")
        L.append(f"   CONSTR-only: {sub_a[only_cn].sum()/1e6:8.2f}M  ({int(only_cn.sum())} rows)")
        L.append(f"   BOTH kw    : {sub_a[both].sum()/1e6:8.2f}M  ({int(both.sum())} rows)")
        L.append(f"   NEITHER    : {sub_a[none].sum()/1e6:8.2f}M  ({int(none.sum())} rows)  <- need eyeball")

    # (B) V_GAMING_EQUIP detail
    geq = con & vl.eq("博彩設施設備優化")
    L.append(f"\n## (B) V_GAMING_EQUIP 建設 — top Cost element descr.:")
    for v, x in a.abs()[geq].groupby(ad[geq]).sum().sort_values(ascending=False).head(10).items():
        L.append(f"   {str(v)[:40]:40s} {x/1e6:7.2f}M")
    L.append(f"\n   top Description (free-text detail):")
    for v, x in a.abs()[geq].groupby(desc[geq].replace('', '(blank)')).sum().sort_values(ascending=False).head(35).items():
        L.append(f"   {str(v)[:62]:62s} {x/1e6:7.3f}M")
    kw_split(geq, "(C) V_GAMING_EQUIP 建設 keyword split")

    # (D) all 建設
    kw_split(con, "(D) ALL 建設 keyword split")
    _w(L)


def _w(L):
    out = ROOT / "results" / "inspect_sjm_23_construction.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
