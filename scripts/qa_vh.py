"""QA the V and H LABELS by surfacing things that look WRONG (not by re-stating the map).

Two checks that actually catch mis-tagging:

  A) V×H suspicious cells — combinations that should almost never co-occur, e.g. a
     gaming vertical (博彩) carrying comp hotel-room / F&B spend, or a facility-upgrade
     vertical carrying performer fees. These distort the report at a glance. Each flagged
     cell drills down to its top account_desc + the rule/source that tagged it.

  B) H-bucket composition — for every horizontal, the top account_code/account_desc
     feeding it AND which source/rule fired. This is how you catch an over-matching rule
     (e.g. an account_desc="Sponsorship" filter sweeping in advertising / hotel / prof rows).

Optionally restrict to one vertical or horizontal to zoom in.

Run:
  python scripts/qa_vh.py --entity sjm --year 25
  python scripts/qa_vh.py --entity sjm --year 25 --h H_SPONSORSHIP   # zoom one H bucket
  python scripts/qa_vh.py --entity galaxy --year 25 --min 1000000    # flag cells > 1M
Outputs results/{ent}_vh_matrix_{year}.tsv  and  results/{ent}_h_audit_{year}.tsv
"""
from __future__ import annotations
import argparse, sys, csv
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd
import yaml

ENT = {"galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
       "vml": "company_4", "melco": "company_5", "mgm": "company_6"}

GAMING_V = {"V_GAMING_VENUE", "V_GAMING_EQUIP"}
COMP_H = {"H_HOTEL_ROOM", "H_FNB", "H_VENUE", "H_COMP_TICKET", "H_COMP_OTHER"}
PERF_V = {"V_CONCERT", "V_SPORT_EVENT", "V_VENUE_PERF_SPORT_MICE", "V_THEME_PARK",
          "V_FOOD_EVENT", "V_ART_EXHIBITION", "V_MICE", "V_COMMUNITY"}


def suspicious(v_id: str, h_id: str) -> str | None:
    """Return a reason string if this V×H combo is implausible, else None."""
    if v_id in GAMING_V and h_id in COMP_H:
        return "gaming vertical with comp/hospitality spend"
    if v_id in GAMING_V and h_id in {"H_PERFORMER", "H_SPONSORSHIP", "H_ADVERTISING", "H_COMP_TICKET"}:
        return "gaming vertical with show/marketing spend"
    if h_id == "H_PERFORMER" and v_id not in PERF_V:
        return "performer fee on a non-entertainment vertical"
    if h_id == "H_HOTEL_ROOM" and v_id in {"V_RESTAURANT"}:
        return "hotel-room comp on a restaurant vertical"
    return None


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--entity", required=True, choices=list(ENT))
    p.add_argument("--year", default="25")
    p.add_argument("--min", type=float, default=200000, help="amount threshold to flag a suspicious V×H cell")
    p.add_argument("--split_min", type=float, default=100000, help="min leaked amount to report an account split across H")
    p.add_argument("--h", default=None, help="zoom: only audit this horizontal_id")
    p.add_argument("--v", default=None, help="zoom: only this vertical_id")
    p.add_argument("--top", type=int, default=12, help="top account_desc per bucket / cell")
    p.add_argument("--detail", action="store_true",
                   help="full enumeration writes project×subproject×vendor grain (big); default collapses to V×H×account (~1-3k rows)")
    args = p.parse_args()
    ent, com = args.entity, ENT[args.entity]

    cfg = yaml.safe_load(Path(f"conf/{com}/parameters.yml").read_text(encoding="utf-8"))
    cols = cfg.get("columns", {}) or {}
    pq = Path(f"data/{ent}/interim/{com}_tagged_rows.parquet")
    if not pq.exists():
        print(f"X {pq} missing — run kedro first"); sys.exit(1)
    df = pd.read_parquet(pq)

    rp = next((c for c in ("report_period", "report_year") if c in df.columns), None)
    if rp:
        df = df[df[rp].astype(str).str.strip().str.startswith(args.year)].copy()
    if args.v: df = df[df.get("vertical_id", "").astype("string") == args.v]
    if args.h: df = df[df.get("horizontal_id", "").astype("string") == args.h]

    amt_col = next((c for c in (cols.get("amount"), "amount_mop", "Reported Amount(MOP)", "amount")
                    if c and c in df.columns), None)
    desc_col = next((c for c in (cols.get("account_desc"), "account_desc", "Account Description",
                                 "Cost element descr.", "Nature of Expenses") if c and c in df.columns), None)
    code_col = next((c for c in (cols.get("account_code"), "account_code", "Account Code", "Account",
                                 "Cost Element") if c and c in df.columns), None)
    proj_col = next((c for c in (cols.get("project"), "project_name", "Project", "Project Name",
                                 "Name of Investment Project") if c and c in df.columns), None)
    sub_col = next((c for c in (cols.get("project_name_cols") or []) if c in df.columns and c != proj_col), None)
    vendor_col = next((c for c in (cols.get("vendor"), "vendor", "Vendor Name", "Name of Supplier",
                                   "AC1 Vendor Name", "supplier") if c and c in df.columns), None)
    descr_col = next((c for c in (cols.get("description"), "description", "Description", "line_memo",
                                  "Expense Description") if c and c in df.columns), None)

    def col_s(name):
        return df[name].astype("string").fillna("") if name and name in df.columns else pd.Series("", index=df.index)

    amt = pd.to_numeric(df[amt_col], errors="coerce").fillna(0) if amt_col else pd.Series(0.0, index=df.index)
    vid = df.get("vertical_id", pd.Series("", index=df.index)).astype("string").fillna("")
    vlab = df.get("vertical_label", vid).astype("string").fillna("")
    vsrc0 = df.get("vertical_source", pd.Series("", index=df.index)).astype("string").fillna("")
    hid = df.get("horizontal_id", pd.Series("", index=df.index)).astype("string").fillna("")
    hlab = df.get("horizontal_label", hid).astype("string").fillna("")
    hsrc = df.get("horizontal_source", pd.Series("", index=df.index)).astype("string").fillna("")
    desc = df[desc_col].astype("string").fillna("") if desc_col else pd.Series("", index=df.index)
    code = df[code_col].astype("string").fillna("") if code_col else pd.Series("", index=df.index)
    proj = col_s(proj_col); sub = col_s(sub_col); vendor = col_s(vendor_col); descr = col_s(descr_col)
    total = amt.sum()
    print(f"[{ent}] year={args.year}  rows={len(df):,}  total={total:,.0f}  (amt={amt_col})")

    base = pd.DataFrame({"vid": vid, "vlab": vlab, "vsrc": vsrc0, "hid": hid, "hlab": hlab,
                         "hsrc": hsrc, "proj": proj, "sub": sub, "vendor": vendor,
                         "code": code, "desc": desc, "descr": descr, "amt": amt})

    # Everything goes to BOTH console and one consolidated per-entity .md (easy to paste).
    L: list[str] = []
    def emit(s: str = "") -> None:
        print(s); L.append(s)

    emit(f"# QA {ent}  year={args.year}")
    emit(f"rows={len(df):,}  total={total:,.0f}  amt_col={amt_col}")
    emit("source legend:  rule = our YAML predominant rule (the kind that can over-match)  |  "
         "rule:row = row override (项目组 column_map OR our targeted when/set)  |  "
         "llm  |  manual = feedback / project broadcast")
    emit("→ a flag on a `rule`/`llm` cell is more likely OUR bug; a flag on `rule:row` "
         "(项目组's own column) is more likely THEIR classification — judge before changing.")

    # ---- A) suspicious V×H cells (the 縱向×橫向 contradictions) ----
    cell = base.groupby(["vid", "hid"]).agg(n=("amt", "size"), amount=("amt", "sum")).reset_index()
    flags = []
    for _, r in cell.iterrows():
        why = suspicious(r["vid"], r["hid"])
        if why and abs(r["amount"]) >= args.min:
            flags.append((r["vid"], r["hid"], int(r["n"]), float(r["amount"]), why))
    flags.sort(key=lambda x: -abs(x[3]))
    emit(f"\n## A) suspicious V×H cells (|amount| >= {args.min:,.0f}) : {len(flags)}")
    for vid_, hid_, n, a, why in flags:
        emit(f"- {vid_} x {hid_}   {a:,.0f}  ({n} rows) — {why}")
        sub = base[(base["vid"] == vid_) & (base["hid"] == hid_)]
        topd = sub.groupby(["code", "desc", "hsrc", "vsrc"]).agg(amount=("amt", "sum"), n=("amt", "size")) \
                  .reset_index().sort_values("amount", key=abs, ascending=False).head(args.top)
        for _, d in topd.iterrows():
            emit(f"    {str(d['code'])[:12]:<12} {str(d['desc'])[:38]:<38} {d['amount']:>13,.0f}  [V:{d['vsrc']} H:{d['hsrc']}]")
    if not flags:
        emit("  (none above threshold)")

    # ---- B) 橫向 composition: top account_desc + source per H (catch over-matching rules) ----
    emit(f"\n## B) 橫向 H composition (top {args.top} account_desc per H — watch misfits)")
    hb = base.groupby("hid").agg(amount=("amt", "sum"), n=("amt", "size")).reset_index() \
             .sort_values("amount", key=abs, ascending=False)
    h_rows = []
    for _, hr in hb.iterrows():
        h = str(hr["hid"]).strip()
        if not h: continue
        emit(f"\n### {h}  total={hr['amount']:,.0f}  ({int(hr['n'])} rows)")
        sub = base[base["hid"] == h]
        topd = sub.groupby(["code", "desc", "hsrc"]).agg(amount=("amt", "sum"), n=("amt", "size")) \
                  .reset_index().sort_values("amount", key=abs, ascending=False).head(args.top)
        for _, d in topd.iterrows():
            emit(f"    {str(d['code'])[:12]:<12} {str(d['desc'])[:40]:<40} {d['amount']:>13,.0f} n={int(d['n']):>5} [{d['hsrc']}]")
            h_rows.append([h, d["code"], d["desc"], d["hsrc"], int(d["n"]), round(float(d["amount"]), 0)])

    # ---- C) 縱向 composition: top account_desc + source per V ----
    emit(f"\n## C) 縱向 V composition (top {args.top} account_desc per V)")
    vb = base.groupby("vid").agg(amount=("amt", "sum"), n=("amt", "size")).reset_index() \
             .sort_values("amount", key=abs, ascending=False)
    for _, vr in vb.iterrows():
        v = str(vr["vid"]).strip()
        if not v: continue
        emit(f"\n### {v}  total={vr['amount']:,.0f}  ({int(vr['n'])} rows)")
        sub = base[base["vid"] == v]
        topd = sub.groupby(["desc", "vsrc"]).agg(amount=("amt", "sum"), n=("amt", "size")) \
                  .reset_index().sort_values("amount", key=abs, ascending=False).head(args.top)
        for _, d in topd.iterrows():
            emit(f"    {str(d['desc'])[:46]:<46} {d['amount']:>13,.0f} n={int(d['n']):>5} [{d['vsrc']}]")

    # ---- D) account split across H buckets (EXHAUSTIVE — not top-N sampled) ----
    # An account_code that lands in >1 H bucket almost always means one of them is wrong.
    # "leaked" = total amount of the account that sits OUTSIDE its dominant H bucket.
    base["key"] = base["code"].where(base["code"].astype("string").str.strip().ne(""), base["desc"])
    ah = base.groupby(["key", "hid"]).agg(amount=("amt", "sum"), n=("amt", "size"),
                                          src=("hsrc", lambda s: ",".join(sorted(set(s))[:3]))).reset_index()
    split_rows = []
    for key, grp in ah.groupby("key"):
        if not str(key).strip():
            continue
        g = grp.reindex(grp["amount"].abs().sort_values(ascending=False).index)
        if g["hid"].nunique() < 2:
            continue
        dom = g.iloc[0]
        leaked = float(g["amount"].iloc[1:].abs().sum())
        if leaked < args.split_min:
            continue
        others = "; ".join(f"{r.hid}:{r.amount:,.0f}[{r.src}]" for r in g.iloc[1:].itertuples())
        split_rows.append((str(key), str(dom["hid"]), float(dom["amount"]), leaked,
                           int(g["hid"].nunique()), others))
    split_rows.sort(key=lambda x: -x[3])
    emit(f"\n## D) account split across >1 H  (leaked >= {args.split_min:,.0f}) : {len(split_rows)} accounts")
    emit("   one account in many H = one is likely mis-tagged; leaked = $ outside its dominant H. "
         "full list in _acct_split_ tsv.")
    for key, domh, doma, leaked, nh, others in split_rows[:40]:
        emit(f"- {str(key)[:44]:<44} dom={domh}({doma:,.0f}) leaked={leaked:,.0f} /{nh}H")
        emit(f"      {others[:220]}")

    out = Path("results"); out.mkdir(exist_ok=True)
    mat = base.pivot_table(index="vlab", columns="hlab", values="amt", aggfunc="sum", fill_value=0)
    mat.to_csv(out / f"{ent}_vh_matrix_{args.year}.tsv", sep="\t", encoding="utf-8-sig")
    # FULL H composition — every account×desc×H (not just the top shown above) for offline filtering
    full = base.groupby(["hid", "code", "desc", "hsrc"]).agg(amount=("amt", "sum"), n=("amt", "size")).reset_index()
    full = full.reindex(full["amount"].abs().sort_values(ascending=False).index)
    full.to_csv(out / f"{ent}_h_audit_{args.year}.tsv", sep="\t", index=False, encoding="utf-8-sig")
    with (out / f"{ent}_acct_split_{args.year}.tsv").open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter="\t")
        w.writerow(["account", "dominant_H", "dominant_amt", "leaked_amt", "n_H", "other_H_breakdown"])
        for r in split_rows:
            w.writerow([r[0], r[1], round(r[2], 0), round(r[3], 0), r[4], r[5]])
    # Enumeration for eyeball review. Default grain = V × H × account_code × label
    # (label = account_desc, or vendor when desc is blank), collapsing across projects so the
    # file is ~1-3k rows instead of 萬行. n_proj = how many distinct projects feed the cell.
    # Use --detail to expand one bucket (with --h/--v) down to project×subproject×vendor.
    base["lab"] = base["desc"].where(base["desc"].astype("string").str.strip().ne(""), base["vendor"])
    if args.detail:
        rich = base.groupby(["vid", "hid", "proj", "sub", "code", "lab", "vsrc", "hsrc"]) \
                   .agg(amount=("amt", "sum"), n=("amt", "size")).reset_index()
        cols_out = ["vid", "hid", "proj", "sub", "code", "lab", "vsrc", "hsrc", "n", "amount"]
    else:
        rich = base.groupby(["vid", "hid", "code", "lab"]).agg(
            amount=("amt", "sum"), n=("amt", "size"), n_proj=("proj", "nunique"),
            vsrc=("vsrc", lambda s: ",".join(sorted({x for x in s if x})[:2])),
            hsrc=("hsrc", lambda s: ",".join(sorted({x for x in s if x})[:2])),
        ).reset_index()
        cols_out = ["vid", "hid", "code", "lab", "n_proj", "n", "vsrc", "hsrc", "amount"]
    rich["_abs"] = rich["amount"].abs()
    rich = rich.sort_values(["vid", "hid", "_abs"], ascending=[True, True, False]).drop(columns="_abs")
    rich = rich[cols_out]
    rich.to_csv(out / f"{ent}_full_{args.year}.tsv", sep="\t", index=False, encoding="utf-8-sig")
    (out / f"{ent}_qa_{args.year}.md").write_text("\n".join(L), encoding="utf-8")
    print(f"\n→ results/{ent}_full_{args.year}.tsv  = {len(rich):,} rows "
          f"({'project×vendor detail' if args.detail else 'V×H×account, collapsed'})")
    print(f"  (+ {ent}_qa_{args.year}.md headline issues, _h_audit_ , _acct_split_ , _vh_matrix_ tsv)")


if __name__ == "__main__":
    main()
