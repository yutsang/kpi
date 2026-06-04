"""audit_health.py — one-shot HEALTH / ANOMALY report across entities, to surface problems.
Per entity × year bucket: Σamt, rows, NG breakdown (+未分類%), top H (+H_OTHER%), top V
(+V_OTHER%), the biggest single sigs (double-count radar), capex/opex × gaming. Flags:
  ⚠ 未分類 > 3%   ⚠ H_OTHER > 5%   ⚠ V_OTHER > 25%   ⚠ a single sig > 5% of the year (double-count)
  ⚠ year amount jump (24 vs 25 ratio > 1.6x or < 0.6x)

Run (Windows):
  python scripts/audit_health.py --all
  python scripts/audit_health.py --entity vml melco
"""
from __future__ import annotations
import argparse, re, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd
import pyarrow.parquet as pq
import yaml

ROOT = Path(__file__).resolve().parent.parent
ENT = {"galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
       "vml": "company_4", "melco": "company_5", "mgm": "company_6"}


def _cn_kw(s) -> str:
    s = str(s)
    if "非博彩" in s:
        return ""
    for kws, ng in [(["博彩", "gaming"], "NG0"), (["海上"], "NG10"),
                    (["外國", "客源", "國際客"], "NG1"), (["會議", "會展", "mice"], "NG2"),
                    (["娛樂", "演唱", "表演"], "NG3"), (["體育", "賽事"], "NG4"),
                    (["文化", "藝術", "文藝"], "NG5"), (["健康", "養生"], "NG6"),
                    (["主題", "遊樂"], "NG7"), (["美食", "餐飲"], "NG8"),
                    (["社區"], "NG9"), (["其他"], "NG11")]:
        if any(k in s or k in s.lower() for k in kws):
            return ng
    return ""


def fuzzy(df, name):
    if not name:
        return None
    if name in df.columns:
        return name
    for c in df.columns:
        if str(c).strip() == str(name).strip():
            return c
    return None


def derive_ng(df, cf, cols):
    names = [cols.get("ng11_category", "")] + [
        (ys.get("columns_override") or {}).get("ng11_category") for ys in (cf.get("yearly_sources") or [])]
    ngcols = []
    for nm in names:
        fc = fuzzy(df, nm)
        if fc and fc not in ngcols:
            ngcols.append(fc)
    for c in df.columns:
        if c not in ngcols and any(k in str(c) for k in ("項目性質", "項目類型", "項目分類", "範疇", "NG11 Category")):
            ngcols.append(c)
    try:
        sys.path.insert(0, str(ROOT / "src"))
        from kpi.lib.conf import load_categories
        from kpi.pipelines.step2_tag_projects._logic import normalize_ng_code
        cats = load_categories()
        def res(x):
            for c in (x, x.upper().replace(" ", "")):
                r = normalize_ng_code(c, cats) or ""
                if r[:2] == "NG" and r[2:].isdigit():
                    return r
            return _cn_kw(x)
    except Exception:
        res = _cn_kw
    out = pd.Series("", index=df.index, dtype="object")
    for fc in ngcols:
        mp = {x: res(x) for x in set(df[fc].astype(str).unique())}
        r = df[fc].astype(str).map(mp).fillna("")
        r = r.where(r.str.fullmatch(r"NG\d+").fillna(False), "")
        out = out.mask(out.eq(""), r)
    return out.replace("", "(未分類)")


def ngnum(s):
    m = re.search(r"NG(\d+)", str(s)); return int(m.group(1)) if m else 99


def audit(ent):
    com = ENT[ent]
    parquet = ROOT / "data" / ent / "output" / f"{com}_kpi_report.parquet"
    if not parquet.exists():
        print(f"\n##### {ent}: X no kpi_report.parquet"); return []
    cf = yaml.safe_load((ROOT / "conf" / com / "parameters.yml").open(encoding="utf-8"))
    cols = cf.get("columns", {}) or {}
    df = pq.read_table(parquet).replace_schema_metadata(None).to_pandas()
    df["_amt"] = pd.to_numeric(df[fuzzy(df, cols.get("amount"))], errors="coerce").fillna(0)
    df["_ng"] = derive_ng(df, cf, cols)
    df["_v"] = df.get("vertical_label", df.get("vertical_id", "")).astype(str)
    df["_h"] = df.get("horizontal_label", df.get("horizontal_id", "")).astype(str)
    df["_hid"] = df.get("horizontal_id", "").astype(str)
    df["_vid"] = df.get("vertical_id", "").astype(str)
    ad = fuzzy(df, cols.get("account_desc")); pj = fuzzy(df, cols.get("project"))
    df["_sig"] = (df[ad].astype(str) if ad else "") + " / " + (df[pj].astype(str).str.slice(0, 26) if pj else "")
    yc = next((c for c in ("report_period", "report_year") if c in df.columns), None)
    yr = df[yc].astype(str) if yc else pd.Series("", index=df.index)
    flags = []
    by_year = {}
    print(f"\n##################  {ent.upper()}  ##################")
    for tag in ("25", "24", "23"):
        sub = df[yr.str.startswith(tag)]
        tot = float(sub["_amt"].sum())
        if sub.empty or abs(tot) < 1:
            continue
        by_year[tag] = tot
        print(f"\n=== {ent} {tag}:  Σ={tot:,.0f}   rows={len(sub):,} ===")
        ng = sub.groupby("_ng")["_amt"].sum().reset_index(); ng["k"] = ng["_ng"].map(ngnum)
        ng = ng.sort_values("k")
        print("  NG: " + "  ".join(f"{r['_ng']}={r['_amt']/tot*100:.0f}%" for _, r in ng.iterrows()))
        uncl = float(sub.loc[sub["_ng"].eq("(未分類)"), "_amt"].sum())
        hoth = float(sub.loc[sub["_hid"].eq("H_OTHER"), "_amt"].sum())
        voth = float(sub.loc[sub["_vid"].eq("V_OTHER"), "_amt"].sum())
        for label, val in [("未分類NG", uncl), ("H_OTHER", hoth), ("V_OTHER", voth)]:
            pct = val / tot * 100 if tot else 0
            thr = {"未分類NG": 3, "H_OTHER": 5, "V_OTHER": 25}[label]
            mark = " ⚠" if pct > thr else ""
            print(f"  {label:9s} {val:>15,.0f}  {pct:4.1f}%{mark}")
            if pct > thr: flags.append(f"{ent} {tag}: {label} {pct:.0f}% (>{thr}%)")
        # double-count radar — biggest single sigs
        big = sub.groupby("_sig")["_amt"].sum().reindex(
            sub.groupby("_sig")["_amt"].sum().abs().sort_values(ascending=False).index).head(3)
        print("  biggest sigs:")
        for s, v in big.items():
            pct = v / tot * 100 if tot else 0
            mark = " ⚠DOUBLE-COUNT?" if abs(pct) > 5 else ""
            print(f"     {str(s)[:50]:52s} {v:>15,.0f}  {pct:4.1f}%{mark}")
            if abs(pct) > 5: flags.append(f"{ent} {tag}: single sig {str(s)[:30]} = {pct:.0f}% of year")
    if "24" in by_year and "25" in by_year and by_year["25"]:
        ratio = by_year["24"] / by_year["25"]
        if ratio > 1.6 or ratio < 0.6:
            print(f"\n  ⚠ {ent}: 24/25 amount ratio = {ratio:.2f}x (24={by_year['24']:,.0f} vs 25={by_year['25']:,.0f})")
            flags.append(f"{ent}: 24 is {ratio:.1f}x of 25")
    return flags


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--entity", nargs="+", choices=sorted(ENT))
    p.add_argument("--all", action="store_true")
    a = p.parse_args()
    ents = sorted(ENT) if a.all else (a.entity or [])
    if not ents:
        print("specify --all or --entity ..."); return
    allflags = []
    for e in ents:
        allflags += audit(e)
    print("\n\n################  ⚠ FLAGS SUMMARY  ################")
    if not allflags:
        print("  (none over threshold)")
    for f in allflags:
        print("  ⚠ " + f)


if __name__ == "__main__":
    main()
