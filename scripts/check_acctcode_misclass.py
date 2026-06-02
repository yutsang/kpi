"""Check whether account_code H rules MISCLASSIFY any signatures.

Unlike audit_acctcode_rules (which tested each rule in isolation), this REPLAYS
the horizontal rules in list order (first-match-wins, exactly like the pipeline),
assigns each signature to the rule that ACTUALLY catches it, then flags sigs whose
DESCRIPTION strongly hints a different H than the assigned one. Those are the real
misclassification candidates — sorted by amount so the biggest-$ ones surface first.

For each sig: winning rule → assigned_H. Compute desc_hint_H from account_desc +
description keywords. If desc_hint_H is confident AND ≠ assigned_H → FLAG.

Output:
  results/acctcode_misclass.tsv   — flagged sigs: entity|code|assigned_H|hint_H|amount|account_desc|sample_desc
  console summary: per entity, total acct-rule amount vs flagged amount (%)

Run (Windows, needs data):
  python scripts/check_acctcode_misclass.py --all
  python scripts/check_acctcode_misclass.py --entity melco
"""
from __future__ import annotations
import argparse, sys, csv, re
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
import pandas as pd
import yaml
from kpi.lib.rules import _predominant_match_one

# strong description → H hints (account_desc + desc). Conservative: only confident cues.
H_HINT = [
 ("H_LABOR", ["payroll","salary","salaries","basic salary","bonus","staff cost","wages",
    "人工","薪","職工薪","工資","outsource casual"]),
 ("H_PERFORMER", ["show production","performer","performers & contractor","artist fee",
    "residency artist","演藝","表演製作","水舞間製作","contract performer","contract entertainment"]),
 ("H_CONSTRUCTION", ["construction in progress","cip","building & improv","building improvement",
    "leasehold improvement","fit-out","fit out","main contract","civil works","建築","裝修工程","判頭"]),
 ("H_FNB", ["cost of sales, food","cost of sales, beverage","food & beverage supplies","china & glassware",
    "pantry supplies","餐飲","食品飲料","食品成本"]),
 ("H_MAINTENANCE", ["repairs and maintenance","repairs & maintenance","maintenance contract",
    "preventive maintenance","r&m","維護","維修"]),
 ("H_LEASE", ["rent expense","equipment rental","lease","operating lease","租賃","租金"]),
 ("H_PROFESSIONAL", ["professional fee","professional service","consultancy","legal fee","audit fee",
    "agency fee","commission","專業服務","顧問費","律師費"]),
 ("H_ADVERTISING", ["promotional expense","advertising","media buy","media-other","digital ad",
    "production","廣告","推廣費","宣傳"]),
 ("H_SPONSORSHIP", ["sponsorship","community events and sponsor","贊助"]),
 ("H_OTHER", ["electricity","water charge","utilities","gas & oil","水電","公用事業"]),
 ("H_LICENSE", ["royalt","license fee","licence fee","ip licensing","授權費","版稅"]),
 ("H_HOTEL_ROOM", ["room revenue","hotel room","lodging","客房","房費"]),
]


def _acct_or_desc_rule(rule):
    cond = rule.get("if") or {}
    return any(k.startswith("account_code") or k.startswith("account_desc") or k=="desc_contains"
               for k in cond)


def desc_hint(account_desc, desc):
    blob = f"{account_desc} {desc}".lower()
    for h, kws in H_HINT:
        for kw in kws:
            if kw in blob:
                return h
    return ""


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--entity", default=None)
    p.add_argument("--all", action="store_true")
    args = p.parse_args()
    ENT = {"galaxy":"company_1","sjm":"company_2","wynn":"company_3",
           "vml":"company_4","melco":"company_5","mgm":"company_6"}
    targets = list(ENT.items()) if args.all else ([(args.entity, ENT[args.entity])] if args.entity else [])
    if not targets: print("Specify --entity or --all"); sys.exit(1)

    out = Path("results"); out.mkdir(exist_ok=True)
    flagged = []; summary = []
    for ent, com in targets:
        cfg = yaml.safe_load(Path(f"conf/{com}/parameters.yml").read_text(encoding="utf-8"))
        rules = ((cfg.get("predominant_rules") or {}).get("horizontal")) or []
        sig_path = Path(f"data/{ent}/interim/{com}_unique_signatures.xlsx")
        if not sig_path.exists():
            print(f"[{ent}] {sig_path} missing — skip"); continue
        sig = pd.read_excel(sig_path)
        ac = "account_code" if "account_code" in sig.columns else None
        ad = "account_desc" if "account_desc" in sig.columns else None
        ds = "desc_samples" if "desc_samples" in sig.columns else ("desc_norm" if "desc_norm" in sig.columns else None)
        amt = "total_amount" if "total_amount" in sig.columns else ("amount" if "amount" in sig.columns else None)
        rc = "row_count" if "row_count" in sig.columns else None

        acct_amt = 0.0; flag_amt = 0.0; n_acct = 0; n_flag = 0
        for _, r in sig.iterrows():
            item = {"account_code": str(r[ac]) if ac else "",
                    "account_desc": str(r[ad]) if ad else "",
                    "desc": str(r[ds]) if ds else "",
                    "desc_samples": str(r[ds]) if ds else "",
                    "description": str(r[ds]) if ds else ""}
            a = abs(float(pd.to_numeric(r[amt], errors="coerce") or 0)) if amt else 0.0
            # winning rule (first match)
            win = None
            for rule in rules:
                if _predominant_match_one(rule, item):
                    win = rule; break
            if win is None:
                continue
            cond = win.get("if") or {}
            is_acct = any(k.startswith("account_code") for k in cond)
            if not is_acct:
                continue  # only checking account_code-resolved sigs
            assigned = str(win.get("then") or win.get("then_scope") or "?")
            acct_amt += a; n_acct += 1
            # TIER 1 (REAL): the GL account NAME itself contradicts the assigned H.
            # TIER 2 (AMBIG): account name is silent but the transaction description
            #                 suggests a different H (generic account, context matters).
            hint_name = desc_hint(item["account_desc"], "")      # account_desc only — authoritative
            hint_desc = desc_hint("", item["desc"])              # description only
            if not assigned.startswith("H_"):
                continue
            tier = ""
            hint = ""
            if hint_name and hint_name != assigned:
                tier = "REAL"; hint = hint_name
            elif hint_desc and hint_desc != assigned:
                tier = "AMBIG"; hint = hint_desc
            if tier:
                if tier == "REAL":
                    flag_amt += a; n_flag += 1
                flagged.append([ent, str(r[ac]) if ac else "", assigned, hint, round(a,0), tier,
                                str(r[ad])[:34] if ad else "", str(r[ds])[:46] if ds else ""])
        summary.append([ent, n_acct, round(acct_amt,0), n_flag, round(flag_amt,0),
                        round(100*flag_amt/acct_amt,1) if acct_amt else 0])

    # REAL first (account-name contradiction), then by amount
    flagged.sort(key=lambda x: (0 if x[5]=="REAL" else 1, -abs(x[4])))
    with (out/"acctcode_misclass.tsv").open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter="\t")
        w.writerow(["entity","code","assigned_H","hint_H","amount","tier","account_desc","sample_desc"])
        w.writerows(flagged)

    real = [r for r in flagged if r[5]=="REAL"]
    print(f"\n=== account_code misclassification check (account-NAME contradiction = REAL) ===")
    print(f"  {'entity':<8} {'acct_sigs':>9} {'acct_amt':>16} {'REAL_sigs':>9} {'REAL_amt':>16} {'REAL%':>6}")
    for s in summary:
        print(f"  {s[0]:<8} {s[1]:>9} {s[2]:>16,.0f} {s[3]:>9} {s[4]:>16,.0f} {s[5]:>6}")
    print(f"\n  REAL (GL account name contradicts assigned H — these are rule bugs): {len(real)} sigs")
    print(f"  AMBIG (generic account, description hints other — context cases): {len(flagged)-len(real)} sigs")
    print(f"  → results/acctcode_misclass.tsv (REAL first)")
    print(f"\n  --- REAL misclassifications by amount (fix these codes) ---")
    seen=set()
    for r in real:
        k=(r[0],r[1],r[2],r[3])
        if k in seen: continue
        seen.add(k)
        print(f"  {r[0]:<7} code={r[1]:<11} {r[2]:<15}→ {r[3]:<15} amt~{r[4]:>13,.0f}  acct_desc={r[6][:32]}")
        if len(seen)>=25: break


if __name__ == "__main__":
    main()
