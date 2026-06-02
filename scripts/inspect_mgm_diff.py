"""Locate the MGM 25 取數 difference vs golden — straight from the raw, no fuzzy name match.

Pipeline 投資方向 25 = ~1.867B, golden (非博彩項目 投資計劃) = Σ. They differ because the pipeline
also carries rows that map to NO golden 項目序號: gaming, GL-pseudo "projects" (人工成本– / 餐飲收支–
pivot rows = the step2 GL-into-project bug), and events not in golden. This breaks the gap down.

Joins raw.project_id (zfill 3) directly to golden 項目序號 (Master col B). wd → capex (CAPEX) /
payroll (WD3) / opex (rest). Writes results/mgm_diff_25.xlsx:
  0_summary       raw total vs matched vs unmatched vs golden
  1_per_項目       golden vs raw (matched by project_id) + Δ
  2_unmatched     raw rows with no golden 序號, grouped + classified (gaming / GL-pseudo / event)
  3_golden_missing golden 項目 with raw=0

Run:
  python scripts/inspect_mgm_diff.py
"""
from __future__ import annotations
import re, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
OPEX_WD = {"Gaming_OPEX", "WD1", "WD2", "WD4", "WD5_Patron"}


def _num(x):
    x = str(x).replace(",", "").strip()
    try:
        return float(x) * 10000.0          # golden 萬元
    except Exception:
        return 0.0


def parse_golden(p):
    out = {}
    for line in p.read_text(encoding="utf-8-sig").splitlines():
        c = [x.strip() for x in line.split("\t")]
        if len(c) < 6 or not re.fullmatch(r"\d+", c[0]):
            continue
        out[c[0].zfill(3)] = {"name": c[1], "payroll": _num(c[2]), "capex": _num(c[3]),
                              "opex": _num(c[4]), "total": _num(c[5])}
    return out


def _classify(name):
    s = str(name)
    if s.startswith("人工成本"):
        return "GL-pseudo 人工成本"
    if s.startswith("餐飲收支"):
        return "GL-pseudo 餐飲收支"
    if any(k in s for k in ("博彩", "娛樂場", "角子機", "智慧娛樂場", "Gaming", "Chips", "DECK MATE", "TABLE INSERT")):
        return "gaming"
    return "event/other (not in golden)"


def main():
    gp = ROOT / "results" / "mgm_golden_25.tsv"
    if not gp.exists():
        print(f"X {gp} missing"); return
    golden = parse_golden(gp)
    g_tot = sum(v["total"] for v in golden.values())

    pq = ROOT / "data/mgm/interim/company_6_raw.parquet"
    if not pq.exists():
        pq = ROOT / "data/mgm/output/company_6_kpi_report.parquet"
    df = pd.read_parquet(pq)
    for c in ("project_id", "project_name", "wd", "amount_mop"):
        if c not in df.columns:
            print(f"X raw missing col {c!r}. cols={list(df.columns)}"); return
    if "report_period" in df.columns:
        df = df[df["report_period"].astype(str).str.startswith("25")].copy()

    # serial = project_id (if it IS a golden 序號) else golden-名 substring match (same as prebuild).
    name2sn = sorted(((g["name"], sn) for sn, g in golden.items() if g.get("name") and len(g["name"]) >= 4),
                     key=lambda x: -len(x[0]))
    _cache = {}

    def _resolve(pid, nm):
        p = str(pid).strip()
        if re.fullmatch(r"0*\d+", p) and p.zfill(3) in golden:
            return p.zfill(3)
        k = str(nm)
        if k in _cache:
            return _cache[k]
        sn = ""
        for gname, gsn in name2sn:
            if gname in k:
                sn = gsn; break
        _cache[k] = sn
        return sn

    df["_sn"] = [_resolve(a, b) for a, b in zip(df["project_id"], df["project_name"])]
    df["_amt"] = pd.to_numeric(df["amount_mop"], errors="coerce").fillna(0)
    wd = df["wd"].astype(str).str.strip()
    df["_capex"] = df["_amt"].where(wd.eq("CAPEX"), 0)
    df["_payroll"] = df["_amt"].where(wd.eq("WD3"), 0)
    df["_opex"] = df["_amt"].where(wd.isin(OPEX_WD), 0)
    df["_matched"] = df["_sn"].isin(golden.keys())

    raw_tot = df["_amt"].sum()
    matched_tot = df[df["_matched"]]["_amt"].sum()
    unmatched_tot = df[~df["_matched"]]["_amt"].sum()

    # 1_per_項目
    gp_rows = df[df["_matched"]].groupby("_sn")[["_amt", "_capex", "_opex", "_payroll"]].sum()
    per = []
    for sn, gd in sorted(golden.items()):
        pi = gp_rows.loc[sn] if sn in gp_rows.index else pd.Series({"_amt": 0, "_capex": 0, "_opex": 0, "_payroll": 0})
        per.append([sn, gd["name"][:34], round(gd["total"]), round(pi["_amt"]), round(pi["_amt"] - gd["total"]),
                    round(gd["capex"]), round(pi["_capex"]), round(gd["opex"]), round(pi["_opex"]),
                    round(gd["payroll"]), round(pi["_payroll"])])
    per_df = pd.DataFrame(per, columns=["序號", "項目", "G_total", "P_total", "Δ", "G_capex", "P_capex",
                                        "G_opex", "P_opex", "G_payroll", "P_payroll"])
    per_df = per_df.reindex(per_df["Δ"].abs().sort_values(ascending=False).index).reset_index(drop=True)

    # 2_unmatched — classify + group by project
    um = df[~df["_matched"]].copy()
    um["_cat"] = um["project_name"].apply(_classify)
    um_proj = (um.groupby(["_cat", "project_name"])["_amt"].agg(["sum", "size"]).reset_index()
               .rename(columns={"sum": "amount", "size": "rows"}))
    um_proj = um_proj.reindex(um_proj["amount"].abs().sort_values(ascending=False).index).reset_index(drop=True)
    um_cat = um.groupby("_cat")["_amt"].agg(["sum", "size"]).reset_index().rename(columns={"sum": "amount", "size": "rows"})

    # 3_golden_missing
    miss = per_df[per_df["P_total"] == 0][["序號", "項目", "G_total"]]

    summ = pd.DataFrame({
        "metric": ["raw_total (pipeline 25)", "matched_to_golden", "UNMATCHED (the diff)",
                   "golden_total", "matched − golden"],
        "amount": [round(raw_tot), round(matched_tot), round(unmatched_tot), round(g_tot), round(matched_tot - g_tot)],
    })
    cat_lines = "\n".join(f"   {r['_cat']:<28} {r['amount']:>16,.0f}  ({int(r['rows'])} rows)"
                          for _, r in um_cat.sort_values("amount", key=abs, ascending=False).iterrows())

    out = ROOT / "results" / "mgm_diff_25.xlsx"
    with pd.ExcelWriter(out, engine="xlsxwriter") as w:
        summ.to_excel(w, sheet_name="0_summary", index=False)
        um_cat.to_excel(w, sheet_name="0_unmatched_cat", index=False)
        per_df.to_excel(w, sheet_name="1_per_項目", index=False)
        um_proj.to_excel(w, sheet_name="2_unmatched", index=False)
        miss.to_excel(w, sheet_name="3_golden_missing", index=False)

    print(f"raw 25 total = {raw_tot:,.0f}")
    print(f"  matched to golden 序號 = {matched_tot:,.0f}")
    print(f"  UNMATCHED (= 你睇唔到嘅 diff) = {unmatched_tot:,.0f}")
    print(f"  golden total = {g_tot:,.0f}   (matched − golden = {matched_tot - g_tot:,.0f})")
    print("\nUNMATCHED 拆分:")
    print(cat_lines)
    print(f"\n→ {out}  (0_summary / 1_per_項目 / 2_unmatched / 3_golden_missing)")


if __name__ == "__main__":
    main()
