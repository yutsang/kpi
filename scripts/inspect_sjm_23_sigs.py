"""SJM 2023 — (A) does the SOURCE have row-level detail (Description/vendor/narrative) we dropped in
the build? and (B) dump current account-level sigs with $ + H + V for row-by-row review.

SJM 23 deliverable is skewed (廣告 + capex) and tagged_rows has NO Description/vendor (build_sjm_23.py
only emits 6 cols). If the original sjm_2023.xlsx capex/opex tabs carry a narrative / 摘要 / vendor /
WBS column, adding it to the build makes signatures finer → the 567M 推廣費 / deposit mega-buckets
become splittable row-by-row (and the deliverable columns line up with other years).

(A) lists every column of sjm_2023.xlsx '項目明細賬capex' + '項目明細賬opex' with non-blank% + sample
    — to find detail columns the build is currently discarding.
(B) dumps SJM 23 tagged_rows grouped by (Cost Element, Cost element descr.) → Σ|amt| / capex-opex /
    current H / top V — the sig menu to review.

Run (Windows):  python scripts/inspect_sjm_23_sigs.py
Output: prints + results/inspect_sjm_23_sigs.txt
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


def find_file(name):
    hits = list((ROOT / "data").rglob(name))
    return hits[0] if hits else None


def main():
    L = ["# inspect_sjm_23_sigs — (A) source detail cols  (B) account-level sig dump"]

    # (A) source columns
    fp = find_file("sjm_2023.xlsx")
    L.append(f"\n## (A) source sjm_2023.xlsx = {fp.relative_to(ROOT) if fp else 'NOT FOUND'}")
    if fp:
        for tab in ("項目明細賬capex", "項目明細賬opex"):
            try:
                raw = pd.read_excel(fp, sheet_name=tab, header=0, dtype=object)
            except Exception as e:
                L.append(f"   [{tab}] read failed: {e}"); continue
            L.append(f"\n   tab '{tab}': {len(raw):,} rows, {len(raw.columns)} cols")
            for c in raw.columns:
                s = raw[c].astype("string").fillna("").str.strip()
                nb = s.ne("").mean() * 100
                if nb < 1:
                    continue
                nun = s[s.ne("")].nunique()
                samp = " | ".join(map(str, s[s.ne("")].value_counts().head(2).index))
                flag = "  <== detail?" if any(k in str(c).lower() for k in
                       ("desc", "摘要", "說明", "narr", "memo", "vendor", "供應", "name", "wbs",
                        "object", "備註", "remark", "text", "用途", "明細")) else ""
                L.append(f"      {str(c)[:34]:34s} nb{nb:4.0f}% uniq{nun:>5}  {samp[:46]}{flag}")

    # (B) account-level sig dump from tagged_rows (23)
    if not TR.exists():
        L.append(f"\n## (B) tagged_rows missing: {TR}"); _w(L); return
    df = pd.read_parquet(TR)
    per = next((c for c in ("report_period", "report_year") if c in df.columns), None)
    if per:
        df = df[df[per].astype(str).eq("23")].copy()
    a = pd.to_numeric(df[AMT], errors="coerce").fillna(0.0)
    tot = a.abs().sum() or 1
    ac = df["Cost Element"].astype("string").fillna("") if "Cost Element" in df.columns else pd.Series("", index=df.index)
    ad = df["Cost element descr."].astype("string").fillna("") if "Cost element descr." in df.columns else pd.Series("", index=df.index)
    hl = df["horizontal_label"].astype("string").fillna("")
    vl = df["vertical_label"].astype("string").fillna("")
    co = df["final_capex_opex"].astype("string").fillna("") if "final_capex_opex" in df.columns else pd.Series("", index=df.index)
    key = ac.str.strip() + " | " + ad.str.strip()
    L.append(f"\n## (B) SJM 23 sigs by (Cost Element | descr.) — {len(df):,} rows  Σ|amt|={a.abs().sum():,.0f}")
    g = pd.DataFrame({"k": key, "amt": a.abs(), "h": hl, "v": vl, "co": co})
    agg = g.groupby("k").agg(amt=("amt", "sum"), n=("amt", "size")).sort_values("amt", ascending=False)
    for k, row in agg.head(60).iterrows():
        sub = g[g.k == k]
        h = sub.groupby("h")["amt"].sum().sort_values(ascending=False)
        htop = " ; ".join(f"{hh}:{x/1e6:.1f}M" for hh, x in h.head(3).items())
        cox = sub.groupby("co")["amt"].sum().to_dict()
        capex_pct = (cox.get("Capex", 0) / max(row["amt"], 1)) * 100
        L.append(f"\n   {str(k)[:54]:54s} {row['amt']/1e6:8.2f}M ({row['amt']/tot*100:4.1f}%) {int(row['n'])}r capex{capex_pct:3.0f}%")
        L.append(f"        H: {htop}")
    _w(L)


def _w(L):
    out = ROOT / "results" / "inspect_sjm_23_sigs.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
