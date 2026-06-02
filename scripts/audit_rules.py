"""Rule linter — audit predominant_rules against the ACTUAL unique items.

Uses the SAME matcher as the pipeline (kpi.lib.rules._predominant_match_one) so
results are faithful to what kedro does.

For --kind horizontal: matches each conf predominant_rules.horizontal rule against
  data/{ent}/interim/{com}_unique_signatures.xlsx (account_code/desc/desc_samples/...).
For --kind vertical:   matches predominant_rules.vertical against
  data/{ent}/interim/{com}_unique_projects.xlsx (project names — the 838 round_pasted rules).

Reports (to console + TSV in data/review/_dump/):
  A) PER-RULE stats   — idx, target, n_items matched, n_rows, Σamount, sample
  B) DEAD rules       — match 0 items (delete candidates)
  C) COLLISIONS       — one item matched by >1 rule with DIFFERENT targets
                        (engine takes FIRST by order → later ones shadowed;
                         if they disagree it's a latent bug, e.g. Golf→V_CONCERT
                         vs a V_SPORT rule, or 'comp room for staff' keyword traps)
  D) DANGER keywords  — desc_contains/project_contains rules whose matched items
                        contain a collision word (comp/staff/payroll/guest)

Run (on Windows where data/ exists):
  python scripts/audit_rules.py --entity mgm --kind horizontal
  python scripts/audit_rules.py --entity galaxy --kind vertical
  python scripts/audit_rules.py --all --kind horizontal
"""
from __future__ import annotations
import argparse, sys, json
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
import pandas as pd
import yaml
from kpi.lib.rules import _predominant_match_one  # faithful matcher

ENTITIES = {"galaxy":"company_1","sjm":"company_2","wynn":"company_3",
            "vml":"company_4","melco":"company_5","mgm":"company_6"}
COLLISION_WORDS = ("comp", "staff", "payroll", "salary", "guest", "patron",
                   "贈", "員工", "薪", "招待")


def _col(df, *cands):
    for c in cands:
        if c in df.columns: return c
    low = {str(c).lower(): c for c in df.columns}
    for c in cands:
        if c.lower() in low: return low[c.lower()]
    return None


def _rule_target(rule: dict) -> str:
    return str(rule.get("then") or rule.get("then_scope") or rule.get("then_any_of") or "?")


def _rule_desc(rule: dict) -> str:
    return json.dumps(rule.get("if") or {}, ensure_ascii=False)


def _build_items(df, kind, cfg):
    """Build list of (item_dict, amount, n_rows, label_str) per row of the xlsx."""
    cols = cfg.get("columns", {}) or {}
    items = []
    if kind == "horizontal":
        ac = _col(df, "account_code"); ad = _col(df, "account_desc")
        dn = _col(df, "desc_norm"); ds = _col(df, "desc_samples")
        ps = _col(df, "project_samples"); jc = _col(df, "job_code_samples")
        co = _col(df, "capex_opex_dominant", "capex_opex")
        rc = _col(df, "row_count", "n_rows"); amt = _col(df, "total_amount", "amount")
        for _, r in df.iterrows():
            item = {
                "account_code": str(r[ac]) if ac else "",
                "account_desc": str(r[ad]) if ad else "",
                "desc_norm": str(r[dn]) if dn else "",
                "desc_samples": str(r[ds]) if ds else "",
                "description": str(r[ds]) if ds else "",
                "project_samples": str(r[ps]) if ps else "",
                "job_code_samples": str(r[jc]) if jc else "",
                "capex_opex": str(r[co]) if co else "",
            }
            label = f"{item['account_code']}|{item['account_desc'][:30]}|{item['desc_samples'][:40]}"
            items.append((item, float(pd.to_numeric(r[amt], errors="coerce") or 0) if amt else 0.0,
                          int(r[rc]) if rc and pd.notna(r[rc]) else 0, label))
    else:  # vertical
        pc = _col(df, cols.get("project", ""), "project", "Project", "project_name",
                  "Name of Investment Project", "SubProject_Name", "Project Name")
        rc = _col(df, "row_count", "n_rows"); amt = _col(df, "total_amount", "amount")
        for _, r in df.iterrows():
            pname = str(r[pc]) if pc else ""
            item = {"project": pname, "project_samples": pname}
            items.append((item, float(pd.to_numeric(r[amt], errors="coerce") or 0) if amt else 0.0,
                          int(r[rc]) if rc and pd.notna(r[rc]) else 0, pname[:70]))
    return items


def audit(ent, com, kind):
    cfg = yaml.safe_load(Path(f"conf/{com}/parameters.yml").read_text(encoding="utf-8"))
    rules = ((cfg.get("predominant_rules") or {}).get(kind)) or []
    if not rules:
        print(f"[{ent}] no predominant_rules.{kind} — skip", flush=True); return

    src = (Path(f"data/{ent}/interim/{com}_unique_signatures.xlsx") if kind == "horizontal"
           else Path(f"data/{ent}/interim/{com}_unique_projects.xlsx"))
    if not src.exists():
        print(f"[{ent}] {src} missing — skip", flush=True); return
    df = pd.read_excel(src)
    items = _build_items(df, kind, cfg)
    print(f"\n===== {ent} {kind}: {len(rules)} rules vs {len(items):,} items =====", flush=True)

    # Match every item against every rule (capture ALL matches for collision detection)
    per_rule = [{"idx": i, "target": _rule_target(r), "if": _rule_desc(r),
                 "n_items": 0, "n_rows": 0, "amount": 0.0, "sample": "", "danger": ""}
                for i, r in enumerate(rules)]
    collisions = []
    for item, amt, nrows, label in items:
        matched = []
        for i, rule in enumerate(rules):
            try:
                if _predominant_match_one(rule, item):
                    matched.append(i)
            except Exception:
                pass
        for i in matched:
            pr = per_rule[i]
            pr["n_items"] += 1; pr["n_rows"] += nrows; pr["amount"] += amt
            if not pr["sample"]:
                pr["sample"] = label
            blob = (item.get("desc_samples","") + " " + item.get("project_samples","") + " " + item.get("account_desc","")).lower()
            if any(w in blob for w in COLLISION_WORDS):
                pr["danger"] = "⚠kw"
        # Collision: >1 rule, differing targets
        if len(matched) > 1:
            targets = {per_rule[i]["target"] for i in matched}
            if len(targets) > 1:
                collisions.append({
                    "item": label, "amount": amt,
                    "winner_idx": matched[0], "winner": per_rule[matched[0]]["target"],
                    "shadowed": "; ".join(f"#{i}->{per_rule[i]['target']}" for i in matched[1:]),
                })

    out_dir = Path("data/review/_dump"); out_dir.mkdir(parents=True, exist_ok=True)
    rep = pd.DataFrame(per_rule).sort_values("amount", ascending=False)
    rep_path = out_dir / f"{ent}_{kind}_rule_audit.tsv"
    rep.to_csv(rep_path, sep="\t", index=False, encoding="utf-8-sig")

    dead = rep[rep["n_items"] == 0]
    danger = rep[(rep["danger"] != "") & (rep["n_items"] > 0)]
    print(f"  rules matching ≥1 item : {(rep['n_items']>0).sum()}", flush=True)
    print(f"  DEAD rules (0 match)   : {len(dead)}", flush=True)
    print(f"  DANGER-keyword rules   : {len(danger)}", flush=True)
    print(f"  COLLISIONS (diff target): {len(collisions)}", flush=True)
    print(f"  → per-rule audit: {rep_path}", flush=True)

    if len(collisions):
        col_df = pd.DataFrame(collisions).reindex(
            pd.DataFrame(collisions)["amount"].abs().sort_values(ascending=False).index)
        col_path = out_dir / f"{ent}_{kind}_rule_collisions.tsv"
        col_df.to_csv(col_path, sep="\t", index=False, encoding="utf-8-sig")
        print(f"  → collisions: {col_path}", flush=True)
        print(f"\n  --- top 15 collisions by |amount| ---", flush=True)
        print(col_df.head(15).to_csv(sep="\t", index=False), flush=True)

    if len(dead):
        print(f"\n  --- DEAD rules (delete candidates) ---", flush=True)
        print(dead[["idx", "target", "if"]].head(40).to_csv(sep="\t", index=False), flush=True)

    if len(danger):
        print(f"\n  --- DANGER-keyword rules (review for collision like 'comp room for staff') ---", flush=True)
        print(danger[["idx", "target", "if", "n_items", "amount", "sample"]].head(40).to_csv(sep="\t", index=False), flush=True)


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--entity", choices=list(ENTITIES), default=None)
    p.add_argument("--all", action="store_true")
    p.add_argument("--kind", choices=["horizontal", "vertical"], default="horizontal")
    args = p.parse_args()
    targets = list(ENTITIES.items()) if args.all else (
        [(args.entity, ENTITIES[args.entity])] if args.entity else [])
    if not targets:
        print("Specify --entity ENT or --all"); sys.exit(1)
    for ent, com in targets:
        audit(ent, com, args.kind)
    print("\n✓ audit TSVs in data/review/_dump/", flush=True)


if __name__ == "__main__":
    main()
