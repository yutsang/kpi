"""Apply the BLIND H classification to the 8,574 VML-2023 sigs, then compare vs 項目組.

Strict order (per user): blind classify (done by workflow) → apply → compare → override.
This script does APPLY + COMPARE only — no reconciliation baked in, so the diff is honest.

Inputs:
  results/vml23_blind_map.jsonl   — {account_desc, default_H, is_mixed, desc_rules:[{keywords,H}], ...}
  results/_vml23_parsed.tsv       — 8,574 sigs (account_code, account_desc, desc_norm)
  results/vml_proj_baseline.jsonl — {account_desc, proj_H, breakdown, mixed}  (項目組, compare only)

Outputs (console + results/vml23_compare.tsv):
  per account: blind_dom_H + blind split  vs  proj_H + proj split ; AGREE / DISAGREE / NO-BASELINE
"""
from __future__ import annotations
import json, re
from pathlib import Path
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
R = ROOT / "results"


def load_jsonl(p):
    return [json.loads(l) for l in open(p, encoding="utf-8") if l.strip()]


def main():
    blind = {a["account_desc"]: a for a in load_jsonl(R / "vml23_blind_map.jsonl")}
    proj = {a["account_desc"]: a for a in load_jsonl(R / "vml_proj_baseline.jsonl")}
    df = pd.read_csv(R / "_vml23_parsed.tsv", sep="\t", dtype=str).fillna("")

    def apply_H(ad, desc):
        a = blind.get(ad)
        if not a:
            return "H_OTHER"
        d = str(desc).lower()
        for rule in (a.get("desc_rules") or []):
            kws = [str(k).lower() for k in (rule.get("keywords") or []) if str(k).strip()]
            if kws and any(k in d for k in kws):
                return rule["H"]
        return a["default_H"]

    df["blind_H"] = [apply_H(ad, dn) for ad, dn in zip(df["account_desc"], df["desc_norm"])]

    rows, agree, disagree, nobase = [], 0, 0, 0
    n_sig_agree = n_sig_total = 0
    for ad, g in df.groupby("account_desc"):
        vc = (g["blind_H"].value_counts(normalize=True) * 100).round()
        blind_dom = vc.index[0]
        blind_bd = "|".join(f"{h}:{int(p)}" for h, p in vc.items() if p >= 5)
        p = proj.get(ad, {})
        proj_H, proj_bd = p.get("proj_H", ""), p.get("breakdown", "")
        if not proj_H:
            status = "NO-BASELINE"; nobase += 1
        elif blind_dom == proj_H:
            status = "AGREE"; agree += 1; n_sig_agree += len(g)
        else:
            status = "DISAGREE"; disagree += 1
        if proj_H:
            n_sig_total += len(g)
        rows.append({"account_desc": ad, "n_sigs": len(g), "blind_dom": blind_dom,
                     "blind_split": blind_bd, "proj_H": proj_H, "proj_split": proj_bd,
                     "status": status, "is_mixed": blind.get(ad, {}).get("is_mixed", "")})
    out = pd.DataFrame(rows).sort_values(["status", "n_sigs"], ascending=[True, False])
    out.to_csv(R / "vml23_compare.tsv", sep="\t", index=False, encoding="utf-8-sig")

    print(f"accounts: AGREE {agree} | DISAGREE {disagree} | NO-BASELINE {nobase}")
    print(f"sig-weighted account-dominant agreement (vs baseline accts): "
          f"{n_sig_agree:,}/{n_sig_total:,} = {n_sig_agree/n_sig_total*100:.1f}%")
    print("\n=== DISAGREEMENTS (blind dominant ≠ 項目組) — these are the meaningful overrides ===")
    for _, r in out[out.status == "DISAGREE"].iterrows():
        print(f"  {r['n_sigs']:5d}  {r['account_desc'][:30]:32s} blind={r['blind_dom']:16s} proj={r['proj_H']:16s}"
              f"  [blind {r['blind_split']}] vs [proj {r['proj_split']}]")
    print("\n=== blind H distribution across 8,574 sigs ===")
    print((df["blind_H"].value_counts()).to_string())


if __name__ == "__main__":
    main()
