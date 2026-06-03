"""Per-account_desc H audit — the RELIABLE 'is this account mixed?' check.

Keyword-scanning desc_norm over-flags (comp accounts carry guest-name noise; broad
patterns like 'po to proper' fire spuriously). The ground truth for mixing is: what H
did the ALREADY-classified rows (mostly 24/25 via 項目組 分類1) actually get for this
account? If an account_desc carries >1 substantial H there, it's genuinely mixed — and
that distribution is also the 項目組 baseline to compare our independent 2023 H against.

Reads tagged_rows.parquet (all years), and per account_desc reports:
  n_all / n_23 / n_classified ; the H distribution among classified rows (H:pct|H:pct…);
  dominant proj_H + %; n_distinct_H(>=10%) ; MIXED flag.

  python scripts/audit_account_h.py --entity vml

Output: results/{ent}_account_h_audit.tsv + .jsonl (paste the .jsonl) + console table.
"""
from __future__ import annotations
import argparse, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd
import pyarrow.parquet as pq
import yaml

ROOT = Path(__file__).resolve().parent.parent
ENTITIES = {"galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
            "vml": "company_4", "melco": "company_5", "mgm": "company_6"}


def fuzzy(df, name):
    if not name:
        return None
    if name in df.columns:
        return name
    for c in df.columns:
        if str(c).strip() == str(name).strip():
            return c
    return None


def audit(ent, year):
    com = ENTITIES[ent]
    p = ROOT / "data" / ent / "interim" / f"{com}_tagged_rows.parquet"
    if not p.exists():
        print(f"[{ent}] X missing {p.relative_to(ROOT)}"); return
    cf = yaml.safe_load((ROOT / "conf" / com / "parameters.yml").open(encoding="utf-8"))
    cols = cf.get("columns", {}) or {}
    df = pq.read_table(p).replace_schema_metadata(None).to_pandas()
    ad = fuzzy(df, cols.get("account_desc"))
    if ad is None or "horizontal_id" not in df.columns or "report_period" not in df.columns:
        print(f"[{ent}] X need account_desc + horizontal_id + report_period"); return
    df["_ad"] = df[ad].astype(str).fillna("")
    df["_H"] = df["horizontal_id"].astype(str)
    yr = df["report_period"].astype(str)
    is_y = yr.str.startswith(str(year))
    classified = ~df["_H"].isin(["H_OTHER", "", "nan", "None"])

    rows = []
    for adv, g in df.groupby("_ad"):
        cg = g[classified.loc[g.index]]
        vc = cg["_H"].value_counts()
        tot = int(vc.sum())
        dist = {h: round(n / tot * 100) for h, n in vc.items()} if tot else {}
        big = {h: pct for h, pct in dist.items() if pct >= 10}
        dom = max(big, key=big.get) if big else ""
        rows.append({
            "account_desc": adv,
            "n_all": len(g),
            "n_23": int(is_y.loc[g.index].sum()),
            "n_classified": tot,
            "proj_H": dom,
            "dom_pct": big.get(dom, 0),
            "n_distinct": len(big),
            "mixed": "MIXED" if len(big) >= 2 else ("" if tot else "NO-BASELINE(23-only)"),
            "breakdown": "|".join(f"{h}:{p}" for h, p in sorted(dist.items(), key=lambda x: -x[1])[:5]),
        })
    out = pd.DataFrame(rows).sort_values(["mixed", "n_all"], ascending=[True, False])
    res = ROOT / "results"; res.mkdir(exist_ok=True)
    out.to_csv(res / f"{ent}_account_h_audit.tsv", sep="\t", index=False, encoding="utf-8-sig")
    out.to_json(res / f"{ent}_account_h_audit.jsonl", orient="records", lines=True, force_ascii=False)

    mixed = out[out["mixed"] == "MIXED"]
    nobase = out[out["mixed"].str.startswith("NO-BASELINE")]
    print(f"\n[{ent}] {len(out)} account_descs | MIXED(項目組用咗>1 H): {len(mixed)} | "
          f"23-only(冇項目組 baseline,要我獨立分): {len(nobase)}")
    print(f"  ► paste: results/{ent}_account_h_audit.jsonl")
    print(f"\n=== MIXED accounts (項目組 themselves split these) ===")
    for _, r in mixed.iterrows():
        print(f"  {r['n_all']:5d} (23={r['n_23']:4d})  {r['account_desc'][:30]:32s} {r['breakdown']}")
    print(f"\n=== 23-only accounts with NO 項目組 baseline (top 15 by n_23) ===")
    for _, r in nobase.sort_values("n_23", ascending=False).head(15).iterrows():
        print(f"  {r['n_23']:5d}  {r['account_desc'][:40]}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--entity", nargs="+", required=True, choices=sorted(ENTITIES))
    ap.add_argument("--year", default="23")
    a = ap.parse_args()
    for e in a.entity:
        audit(e, a.year)


if __name__ == "__main__":
    main()
