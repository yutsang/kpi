"""Audit ONLY the account_code horizontal rules — measure their accuracy by
checking whether each rule's matched signatures are homogeneous (one GL code /
prefix → one consistent account_desc → one H is trustworthy; diverse desc = suspect).

縱向 view: one row per account_code rule, showing target H + evidence + a flag,
so you can eyeball whether the code→H mapping is right BEFORE deciding to keep
the account_code rulebase.

For each account_code rule (account_code_equals / account_code_prefix) in
predominant_rules.horizontal, matched against {com}_unique_signatures.xlsx:

  entity | code | match | target_H | n_sigs | n_rows | amount
        | n_distinct_desc | top_account_desc | sample_desc | flag

flag:
  ⚠DIVERSE   — equals-rule whose sigs have >2 distinct account_desc (a single GL
               code should be ~1 desc; many = possible data issue / wrong code)
  ⚠H_HINT    — a matched sig's account_desc/desc keyword suggests a DIFFERENT H
               than assigned (e.g. desc says 'construction' but rule → H_PROFESSIONAL)
  (blank)    — homogeneous, high confidence

Also lists desc_contains horizontal rules separately (the fragile kind) so you
can see how much of H still depends on keyword rules.

Run (Windows, needs data):
  python scripts/audit_acctcode_rules.py --entity galaxy
  python scripts/audit_acctcode_rules.py --all
"""
from __future__ import annotations
import argparse, sys, csv, re
from pathlib import Path
from collections import Counter
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
import pandas as pd
import yaml
from kpi.lib.rules import _predominant_match_one

ENTITIES = {"galaxy":"company_1","sjm":"company_2","wynn":"company_3",
            "vml":"company_4","melco":"company_5","mgm":"company_6"}

# keyword → H hint, to flag account_code rules whose sigs look like a different H
H_HINT = [
    ("H_CONSTRUCTION", ["construction","building","fit-out","fit out","leasehold improv","civil works","建築","裝修","工程"]),
    ("H_LABOR", ["payroll","salary","salaries","staff cost","人工","薪","職工"]),
    ("H_FNB", ["cost of sales, food","cost of sales, beverage","food & beverage","餐飲","食品"]),
    ("H_PERFORMER", ["show production","performer","artist fee","演藝","表演製作"]),
    ("H_MAINTENANCE", ["repairs and maintenance","repair & maintenance","維護","維修"]),
    ("H_LEASE", ["rent expense","lease","租賃","租金"]),
    ("H_PROFESSIONAL", ["professional fee","consultancy","legal fee","agency fee","專業服務"]),
    ("H_ADVERTISING", ["promotional","advertising","media","廣告","推廣"]),
]


def _is_acctcode_rule(rule):
    cond = rule.get("if") or {}
    return any(k in cond for k in ("account_code_equals","account_code_prefix"))


def _is_desc_rule(rule):
    cond = rule.get("if") or {}
    return any(k in cond for k in ("desc_contains","account_desc_contains","account_desc_equals","account_desc_prefix"))


def _code_of(rule):
    cond = rule.get("if") or {}
    if "account_code_equals" in cond:
        v = cond["account_code_equals"]; return ("equals", ",".join(v) if isinstance(v,list) else str(v))
    if "account_code_prefix" in cond:
        v = cond["account_code_prefix"]; return ("prefix", ",".join(v) if isinstance(v,list) else str(v))
    return ("?", "?")


def _hint_flag(target, descs):
    blob = " ".join(descs).lower()
    for h, kws in H_HINT:
        if h == target: continue
        if any(k in blob for k in kws):
            # only flag if the assigned target's own keywords are ABSENT
            tgt_kws = next((kk for hh,kk in H_HINT if hh==target), [])
            if not any(k in blob for k in tgt_kws):
                return f"⚠H_HINT:{h}"
    return ""


def audit(ent, com, out_rows, desc_summary):
    cfg = yaml.safe_load(Path(f"conf/{com}/parameters.yml").read_text(encoding="utf-8"))
    rules = ((cfg.get("predominant_rules") or {}).get("horizontal")) or []
    acct_rules = [r for r in rules if _is_acctcode_rule(r)]
    desc_rules = [r for r in rules if _is_desc_rule(r) and not _is_acctcode_rule(r)]

    sig_path = Path(f"data/{ent}/interim/{com}_unique_signatures.xlsx")
    if not sig_path.exists():
        print(f"[{ent}] {sig_path} missing — skip"); return
    sig = pd.read_excel(sig_path)
    ac = next((c for c in ("account_code",) if c in sig.columns), None)
    ad = next((c for c in ("account_desc",) if c in sig.columns), None)
    ds = next((c for c in ("desc_samples","desc_norm") if c in sig.columns), None)
    rc = next((c for c in ("row_count","n_rows") if c in sig.columns), None)
    amt = next((c for c in ("total_amount","amount") if c in sig.columns), None)

    items = []
    for _, r in sig.iterrows():
        items.append({
            "account_code": str(r[ac]) if ac else "",
            "account_desc": str(r[ad]) if ad else "",
            "desc_samples": str(r[ds]) if ds else "",
            "_amt": float(pd.to_numeric(r[amt], errors="coerce") or 0) if amt else 0.0,
            "_rows": int(r[rc]) if rc and pd.notna(r[rc]) else 0,
        })

    for rule in acct_rules:
        mt, code = _code_of(rule)
        target = str(rule.get("then") or rule.get("then_scope") or "?")
        matched = [it for it in items if _predominant_match_one(rule, it)]
        if not matched:
            out_rows.append([ent, code, mt, target, 0, 0, 0.0, 0, "", "", "⚠DEAD"]); continue
        descs = [m["account_desc"] for m in matched if m["account_desc"]]
        dd = Counter(descs)
        n_distinct = len(dd)
        top_desc = dd.most_common(1)[0][0] if dd else ""
        sample_desc = next((m["desc_samples"][:50] for m in matched if m["desc_samples"]), "")
        amount = sum(m["_amt"] for m in matched)
        nrows = sum(m["_rows"] for m in matched)
        flag = ""
        if mt == "equals" and n_distinct > 2:
            flag = "⚠DIVERSE"
        if not flag:
            flag = _hint_flag(target, descs[:20] + [sample_desc])
        out_rows.append([ent, code, mt, target, len(matched), nrows, round(amount,2),
                         n_distinct, top_desc[:40], sample_desc, flag])

    # ── Coverage: classify EACH sig by which rule-type resolves it first ──
    # (mirrors pipeline order: pattern rules in list order, acct_code vs desc).
    tot_amt = sum(abs(x["_amt"]) for x in items)
    cov = {"acct": 0.0, "desc": 0.0, "uncovered": 0.0}
    cov_n = {"acct": 0, "desc": 0, "uncovered": 0}
    for it in items:
        a = abs(it["_amt"])
        hit_acct = any(_predominant_match_one(r, it) for r in acct_rules)
        hit_desc = any(_predominant_match_one(r, it) for r in desc_rules)
        if hit_acct:   cov["acct"] += a; cov_n["acct"] += 1
        elif hit_desc: cov["desc"] += a; cov_n["desc"] += 1
        else:          cov["uncovered"] += a; cov_n["uncovered"] += 1
    desc_summary.append([ent, len(acct_rules), len(desc_rules), len(items),
                         round(tot_amt,0),
                         round(100*cov["acct"]/tot_amt,1) if tot_amt else 0,
                         round(100*cov["desc"]/tot_amt,1) if tot_amt else 0,
                         round(100*cov["uncovered"]/tot_amt,1) if tot_amt else 0,
                         cov_n["uncovered"]])


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--entity", choices=list(ENTITIES), default=None)
    p.add_argument("--all", action="store_true")
    args = p.parse_args()
    targets = list(ENTITIES.items()) if args.all else (
        [(args.entity, ENTITIES[args.entity])] if args.entity else [])
    if not targets: print("Specify --entity or --all"); sys.exit(1)

    out = Path("results"); out.mkdir(exist_ok=True)
    rows = []; desc_summary = []
    for ent, com in targets:
        audit(ent, com, rows, desc_summary)

    rows.sort(key=lambda r: -abs(r[6]))
    tsv = out / "acctcode_rule_audit.tsv"
    with tsv.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter="\t")
        w.writerow(["entity","code","match","target_H","n_sigs","n_rows","amount",
                    "n_distinct_desc","top_account_desc","sample_desc","flag"])
        w.writerows(rows)

    n_flag = sum(1 for r in rows if r[10])
    print(f"\n=== {len(rows)} account_code rules audited → {tsv} ===")
    print(f"  flagged (⚠DIVERSE/H_HINT/DEAD): {n_flag}")
    print(f"\n  COVERAGE — % of total sig |amount| resolved by each rule-type")
    print(f"  (acct% high = keep rules saves work; uncovered% = must classify regardless)")
    print(f"  {'entity':<8} {'acct_r':>6} {'desc_r':>6} {'sigs':>7} {'tot_amt':>16} {'acct%':>6} {'desc%':>6} {'uncov%':>7} {'uncov_sigs':>10}")
    cov_tsv = out / "acctcode_coverage.tsv"
    with cov_tsv.open("w", encoding="utf-8-sig", newline="") as f:
        cw = csv.writer(f, delimiter="\t")
        cw.writerow(["entity","acct_rules","desc_rules","n_sigs","total_amount","acct_pct","desc_pct","uncovered_pct","uncovered_sigs"])
        for d in desc_summary:
            cw.writerow(d)
            print(f"  {d[0]:<8} {d[1]:>6} {d[2]:>6} {d[3]:>7} {d[4]:>16,.0f} {d[5]:>6} {d[6]:>6} {d[7]:>7} {d[8]:>10}")
    print(f"  → {cov_tsv}")
    print(f"\n  --- flagged account_code rules (eyeball these) ---")
    for r in [r for r in rows if r[10]][:30]:
        print(f"  {r[0]:<7} {r[1]:<14} {r[3]:<16} amt={r[6]:>15,.0f} distinct_desc={r[7]:<3} {r[10]}  ({r[8]})")


if __name__ == "__main__":
    main()
