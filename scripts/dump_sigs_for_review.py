"""Dump unique signatures (橫向/H-axis review unit) in a compact pasteable form.

Reads data/{ent}/interim/{com}_unique_signatures.xlsx and emits, per entity, the
ONE thing needed to fix a horizontal misclassification:

  account_code | account_desc | desc_sample | current_H | source | n_rows | amount | projects | flag

- current_H  = manual_horizontal if set else llm_horizontal (matches step4 precedence)
- source     = tag_source (rule:* / acct_code:* / llm / manual / rerank)
- flag       = ⚠RULE if a rule (not LLM) assigned it AND the desc looks ambiguous
               (heuristic: desc_samples contains a 'comp/staff/payroll' collision word
                while H is the opposite family) — surfaces "comp room for staff" bugs

Sorted by |amount| desc so the biggest-impact sigs come first.

Output:
  data/review/_dump/{ent}_sigs.tsv          (full, open + paste chunks)
  + prints top-N to console for direct paste

Run:
  python scripts/dump_sigs_for_review.py --entity mgm --top 80
  python scripts/dump_sigs_for_review.py --all --top 50
  python scripts/dump_sigs_for_review.py --entity galaxy --rules-only   # only rule-tagged sigs
"""
from __future__ import annotations
import argparse, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ENTITIES = {"galaxy":"company_1","sjm":"company_2","wynn":"company_3",
            "vml":"company_4","melco":"company_5","mgm":"company_6"}

# Collision heuristic: words that, when present in description, often indicate the
# row is NOT what a naive keyword rule would assume.
COMP_WORDS = ("comp", "guest", "patron", "client", "贈", "招待", "免費")
LABOR_WORDS = ("staff", "payroll", "salary", "employee", "員工", "薪", "人工")


def _col(df, *cands):
    for c in cands:
        if c in df.columns: return c
    low = {str(c).lower(): c for c in df.columns}
    for c in cands:
        if c.lower() in low: return low[c.lower()]
    return None


def _flag(h, desc):
    """Heuristic collision flag for rule-assigned sigs."""
    d = str(desc).lower()
    h = str(h)
    has_comp = any(w in d for w in COMP_WORDS)
    has_labor = any(w in d for w in LABOR_WORDS)
    # H_LABOR but description screams comp/guest → likely "comp ... for staff" trap
    if h == "H_LABOR" and has_comp and not has_labor:
        return "⚠COMP→LABOR?"
    # Comp family but description has no comp signal
    if h in ("H_HOTEL_ROOM","H_FNB","H_COMP_TICKET","H_COMP_OTHER","H_VENUE") and has_labor and not has_comp:
        return "⚠LABOR→COMP?"
    return ""


def dump_one(ent, com, top, rules_only, batch):
    sig_path = Path(f"data/{ent}/interim/{com}_unique_signatures.xlsx")
    if not sig_path.exists():
        print(f"[{ent}] {sig_path} missing — skip", flush=True); return
    df = pd.read_excel(sig_path)

    ac = _col(df, "account_code")
    ad = _col(df, "account_desc")
    ds = _col(df, "desc_samples", "desc_norm", "description")
    ml = _col(df, "manual_horizontal")
    ll = _col(df, "llm_horizontal", "horizontal_id")
    ts = _col(df, "tag_source", "horizontal_source")
    rc = _col(df, "row_count", "n_rows")
    amt = _col(df, "total_amount", "amount", "amount_mop")
    ps = _col(df, "project_samples", "project")
    sg = _col(df, "signature", "sig")          # the step3 override KEY (maps OUR H back)
    dn = _col(df, "desc_norm")

    out = pd.DataFrame()
    out["signature"] = df[sg].astype(str) if sg else ""   # FIRST col = override key
    out["desc_norm"] = df[dn].astype(str).str.slice(0, 80) if dn else ""
    out["account_code"] = df[ac].astype(str) if ac else ""
    out["account_desc"] = df[ad].astype(str) if ad else ""
    out["desc_sample"] = df[ds].astype(str).str.slice(0, 80) if ds else ""
    # current_H = manual else llm
    man = df[ml].astype(str).str.strip() if ml else pd.Series([""]*len(df))
    llm = df[ll].astype(str).str.strip() if ll else pd.Series([""]*len(df))
    out["current_H"] = man.where(man.ne("") & man.ne("nan"), llm)
    out["source"] = df[ts].astype(str) if ts else ""
    out["n_rows"] = df[rc] if rc else 0
    out["amount"] = pd.to_numeric(df[amt], errors="coerce").fillna(0) if amt else 0
    out["projects"] = df[ps].astype(str).str.slice(0, 50) if ps else ""
    out["flag"] = [_flag(h, d) for h, d in zip(out["current_H"], out["desc_sample"])]

    if rules_only:
        out = out[out["source"].str.startswith(("rule", "acct_code"), na=False)]

    out = out.reindex(out["amount"].abs().sort_values(ascending=False).index)

    out_dir = Path("data/review/_dump")
    out_dir.mkdir(parents=True, exist_ok=True)
    tsv = out_dir / f"{ent}_sigs.tsv"
    out.to_csv(tsv, sep="\t", index=False, encoding="utf-8-sig")

    # Batch files (≤batch rows each) for pasting as text, sorted big-amount first.
    n = len(out)
    n_batches = (n + batch - 1) // batch if n else 0
    for b in range(n_batches):
        chunk = out.iloc[b*batch:(b+1)*batch]
        bp = out_dir / f"{ent}_sigs_batch{b+1:02d}of{n_batches:02d}.tsv"
        chunk.to_csv(bp, sep="\t", index=False, encoding="utf-8-sig")

    n_flag = (out["flag"] != "").sum()
    print(f"\n===== {ent}: {len(out):,} sigs ({n_flag} collision-flagged) =====", flush=True)
    print(f"  full: {tsv}", flush=True)
    print(f"  {n_batches} paste-batch files (≤{batch} rows each): {ent}_sigs_batch01of{n_batches:02d}.tsv ...", flush=True)
    print(f"----- top {min(top,len(out))} by |amount| (preview) -----", flush=True)
    print(out.head(top).to_csv(sep="\t", index=False), flush=True)


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--entity", choices=list(ENTITIES), default=None)
    p.add_argument("--all", action="store_true")
    p.add_argument("--top", type=int, default=60)
    p.add_argument("--batch", type=int, default=3000, help="rows per paste-batch TSV file")
    p.add_argument("--rules-only", action="store_true", help="only rule/acct_code-tagged sigs (rule audit)")
    args = p.parse_args()

    targets = list(ENTITIES.items()) if args.all else (
        [(args.entity, ENTITIES[args.entity])] if args.entity else [])
    if not targets:
        print("Specify --entity ENT or --all"); sys.exit(1)
    for ent, com in targets:
        dump_one(ent, com, args.top, args.rules_only, args.batch)
    print("\n✓ TSV in data/review/_dump/ — paste chunks for H-rule review.", flush=True)


if __name__ == "__main__":
    main()
