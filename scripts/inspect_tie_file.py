r"""Generic tie-file inspector for the NEW per-entity raw files (data/<ent>/raw or data/raw).
Reads ONLY the 'data' tab. Shows columns, picks the year/period col, sums each bucket vs golden
(reference_golden_buckets), checks dup unique_id + amount-tie, and lists 調整/remark-like cols.

Auto-detects amount col (amount_mop / Val/COArea Crcy / 調整後金額 / adjusted_amount / Amount - Amended …)
and year col (year / 項目期間 / years / 是否2025年 / Yr related / Report Years…); pass --amount / --year
to override.

Run (Windows):
  python scripts/inspect_tie_file.py --file data\sjm\raw\sjm_2024.xlsx --entity sjm --report 2024
  python scripts/inspect_tie_file.py --file data\mgm\raw\mgm_2025.xlsx  --entity mgm --report 2025
Output: prints + results/inspect_tie_<stem>.txt
"""
from __future__ import annotations
import argparse, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
GOLDEN_WAN = {
    "galaxy": {"25": 326626, "25_24SY": 18766, "25_23SY": 3, "24": 248981, "24_23SY": 17489, "23": 119395},
    "sjm":    {"25": 124661, "25_24SY": 63594, "25_23SY": 7971, "24": 155980, "24_23SY": 59681, "23": 59346},
    "wynn":   {"25": 209699, "25_24SY": 5611, "25_23SY": 40, "24": 194885, "24_23SY": 4090, "23": 134168},
    "vml":    {"25": 252310, "25_24SY": 39, "25_23SY": 0, "24": 445636, "24_23SY": -841, "23": 134729},
    "melco":  {"25": 243470, "25_24SY": 13125, "25_23SY": 0, "24": 174252, "24_23SY": 26011, "23": 99522},
    "mgm":    {"25": 186142, "25_24SY": 64264, "25_23SY": 5527, "24": 149763, "24_23SY": 63985, "23": 95033},
}
AMOUNT_CANDS = ["amount_mop", "Val/COArea Crcy", "調整後金額", "adjusted_amount", "Amount - Amended",
                "MOP Amt", "本位幣金額", "Entry Voucher Amount/ Expense Amount", "Debit minus Credit"]
YEAR_CANDS = ["year", "項目期間", "years", "是否2025年", "是否2024年", "Yr related",
              "Report Years-重分類2023年期後調整後", "Report Years"]


def _num(s):
    return pd.to_numeric(s.astype("string").str.replace(",", "", regex=False), errors="coerce").fillna(0.0)


def _s(df, c):
    v = df.get(c)
    if v is None: return pd.Series("", index=df.index)
    if isinstance(v, pd.DataFrame): v = v.iloc[:, 0]
    return v.astype("string").fillna("").str.strip()


def _find(df, cands):
    for c in cands:
        if c in df.columns: return c
        hit = next((x for x in df.columns if str(c).lower() in str(x).lower()), None)
        if hit: return hit
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--file", required=True)
    ap.add_argument("--entity", default=None)
    ap.add_argument("--report", default=None, help="report year e.g. 2024 (bucket prefix)")
    ap.add_argument("--amount", default=None)
    ap.add_argument("--year", default=None)
    a = ap.parse_args()
    p = Path(a.file)
    if not p.is_absolute(): p = ROOT / a.file
    L = [f"# inspect_tie_file — {p.name}"]
    xls = pd.ExcelFile(p)
    sh = next((s for s in xls.sheet_names if str(s).strip().lower() == "data"), xls.sheet_names[0])
    df = pd.read_excel(p, sheet_name=sh, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    L.append(f"sheets={xls.sheet_names}  reading={sh!r}  rows={len(df):,}  cols={len(df.columns)}")
    L.append("ALL columns:")
    for i in range(0, len(df.columns), 6):
        L.append("   " + " | ".join(str(c)[:22] for c in df.columns[i:i + 6]))

    amt_c = a.amount or _find(df, AMOUNT_CANDS)
    yr_c = a.year or _find(df, YEAR_CANDS)
    L.append(f"\namount col = {amt_c!r}   year col = {yr_c!r}")
    if not amt_c:
        L.append("!! no amount col — pass --amount"); _w(L, p); return
    amt = _num(_s(df, amt_c))
    L.append(f"Σ amount = {amt.sum()/1e6:,.2f}M   blank amount rows = {int(_num(_s(df,amt_c)).eq(0).sum()):,}")

    gold = {k: v * 1e4 for k, v in (GOLDEN_WAN.get((a.entity or '').lower()) or {}).items()}

    # ── per-bucket amount columns (sjm 24/23: 2024年度金額+2024年度調整; mgm 25: 25_Amt/24_Amt/23_Amt) ──
    BUCKET_COLS = {  # bucket → (amount col substrings, adjust col substrings)
        "25": (["25_Amt"], []), "25_24SY": (["24_Amt"], []), "25_23SY": (["23_Amt"], []),
        "24": (["2024年度金額", "2024年度調整后"], ["2024年度調整"]),
        "24_23SY": (["2023年度金額"], ["2023年度調整"]),
        "23": (["2023年度調整后", "2023年度金額"], []),
    }
    found_bucket = False
    L.append("\n## per-bucket amount-col sums (vs golden):")
    for bk, (acands, adjcands) in BUCKET_COLS.items():
        ac = _find(df, acands)
        if not ac: continue
        found_bucket = True
        s = _num(_s(df, ac))
        adc = _find(df, adjcands) if adjcands else None
        if adc and adc != ac:
            s = s + _num(_s(df, adc))
        g = gold.get(bk)
        star = f"  golden={g/1e6:,.2f}M  Δ={(s.sum()-g)/1e6:,.2f}M" if g is not None else ""
        nb = int((s != 0).sum())
        L.append(f"   {bk:8s} = {ac!r}{('+'+adc) if adc and adc!=ac else ''}: Σ={s.sum()/1e6:>11,.2f}M  ({nb} nonzero rows){star}")
    if not found_bucket:
        L.append("   (no per-bucket amount cols found)")

    if yr_c:
        yv = _s(df, yr_c)
        L.append(f"\n## by {yr_c!r} value (vs golden, report={a.report}):")
        for v in sorted(yv.unique()):
            m = yv.eq(v)
            s = amt[m].sum()
            star = ""
            for b, g in gold.items():
                if g and abs(s - g) <= max(abs(g) * 0.01, 2e4): star += f"  ★≈{b}"
            L.append(f"   {str(v)[:26]:26s} {int(m.sum()):>7,}r  Σ={s/1e6:>11,.2f}M{star}")

    uid = _find(df, ["unique_id", "唯一識別碼", "KPMG唯一識別碼", "RecordID"])
    if uid:
        u = _s(df, uid)
        L.append(f"\nunique_id={uid!r}: non-blank={int(u.ne('').sum()):,} distinct={u[u.ne('')].nunique():,}"
                 f"  exact-dup rows={int(df.duplicated(keep=False).sum()):,}")
    # 調整 / remark cols
    adj = [c for c in df.columns if "調整" in str(c) or "remark" in str(c).lower() or "備註" in str(c)]
    if adj:
        L.append("\n調整/remark-like cols:")
        for c in adj[:14]:
            s = _s(df, c)
            L.append(f"   {str(c)[:40]:40s} nonblank={int(s.ne('').sum()):>6,}  Σnum={_num(s).sum()/1e6:>9.2f}M")
    # capex/opex distinct
    cc = _find(df, ["capex_opex", "報告capex/opex", "Source", "CAPEX/OPEX", "Ledger Type"])
    if cc:
        L.append(f"\ncapex/opex col {cc!r}: " + "  ".join(f"{k}={n}" for k, n in
                 _s(df, cc).replace("", "(blank)").value_counts().head(12).items()))
    _w(L, p)


def _w(L, p):
    out = ROOT / "results" / f"inspect_tie_{p.stem}.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste back")


if __name__ == "__main__":
    main()
