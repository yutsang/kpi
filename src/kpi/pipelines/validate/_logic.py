"""Validate pipeline — structural cross-validation of LLM tagging quality.

Different from audit (which lists rows worth human spot-check):
this pipeline asks "are LLM outputs internally and externally consistent?"

Checks:
  1. Self-consistency:        same signature -> same horizontal (sanity, should be 100%)
  2. Account coherence:       same account_code -> dominant horizontal % (low % = LLM unstable)
  3. Keyword consistency:     desc contains '客房' -> H_HOTEL_ROOM expected, etc.
  4. NG x horizontal:         per NG category, top horizontals (sanity vs business expectation)
  5. V x H amount matrix:     full crosstab + flagged suspicious cells
  6. Confidence x amount:     bucketed risk view
  7. V_OTHER / H_OTHER share: amount + count fallback rate
  8. Random sample:           50 rows for human spot-check

Console: prints paste-friendly summary tables for each check.
File:    <entity>_validation_report.xlsx with one sheet per check.
"""
from __future__ import annotations

import pandas as pd

from kpi.lib.conf import load_categories, load_config, path
from kpi.lib.io_setup import force_unbuffered_io

force_unbuffered_io()

# desc keyword -> expected horizontal_id. Tweak as needed per entity.
KEYWORD_TESTS = {
    "客房": "H_HOTEL_ROOM",
    "酒店房": "H_HOTEL_ROOM",
    "餐飲": "H_FNB",
    "飲食": "H_FNB",
    "餐廳": "H_FNB",
    "薪資": "H_LABOR",
    "工資": "H_LABOR",
    "薪金": "H_LABOR",
    "顧問": "H_PROFESSIONAL",
    "consulting": "H_PROFESSIONAL",
    "贊助": "H_SPONSORSHIP",
    "sponsor": "H_SPONSORSHIP",
    "建設": "H_CONSTRUCTION",
    "裝修": "H_CONSTRUCTION",
    "施工": "H_CONSTRUCTION",
    "場地": "H_VENUE",
    "venue": "H_VENUE",
}


def _print_header(title: str):
    print(f"\n{'=' * 70}\n  {title}\n{'=' * 70}", flush=True)


def main():
    cfg = load_config()
    cats = load_categories()
    company = cfg["company"]["code"]
    interim = path(cfg, "interim_dir")
    out_dir = path(cfg, "output_dir")
    cols = cfg["columns"]

    parquet = interim / f"{company}_tagged_rows.parquet"
    sig_xlsx = interim / f"{company}_unique_signatures.xlsx"
    if not parquet.exists():
        raise FileNotFoundError(f"Run step 4 first. Missing: {parquet}")
    if not sig_xlsx.exists():
        raise FileNotFoundError(f"Run step 3 first. Missing: {sig_xlsx}")

    print(f"Loading {parquet.name} ...", flush=True)
    df = pd.read_parquet(parquet)
    sigs = pd.read_excel(sig_xlsx)
    n = len(df)
    amt = pd.to_numeric(df[cols["amount"]], errors="coerce")
    df["_abs_amt"] = amt.abs()
    total_amt = df["_abs_amt"].sum()
    print(f"  rows={n:,}  signatures={len(sigs):,}  total |amount|={total_amt:,.0f}", flush=True)

    out_path = out_dir / f"{company}_validation_report.xlsx"
    writer = pd.ExcelWriter(out_path, engine="xlsxwriter")

    # ============ 1. Self-consistency ============
    _print_header("1. Self-consistency (same signature -> same horizontal_id)")
    by_sig = df.groupby("signature")["horizontal_id"].nunique()
    n_inconsistent = (by_sig > 1).sum()
    print(f"  signatures with >1 distinct horizontal_id: {n_inconsistent} (expected 0)", flush=True)
    if n_inconsistent > 0:
        bad = by_sig[by_sig > 1].head(20)
        print(bad.to_string())
        bad.to_excel(writer, sheet_name="1_self_consistency_violations")

    # ============ 2. Account-code coherence ============
    _print_header("2. Account-code coherence (low % = LLM inconsistent for that account)")
    acc_col = cols.get("account_code", "")
    desc_col = cols.get("account_desc", "")
    rows = []
    if acc_col and acc_col in df.columns:
        top_accs = df.groupby(acc_col)["_abs_amt"].sum().nlargest(25).index.tolist()
        for acc in top_accs:
            sub = df[df[acc_col].eq(acc)]
            adesc = sub[desc_col].iloc[0] if desc_col in sub.columns else ""
            # Row-count weighting (how often each horizontal is picked)
            h_cnt = sub["horizontal_label"].value_counts(normalize=True)
            top_h_cnt = h_cnt.head(3)
            # Amount weighting (a 50M misclass matters more than 1000 x 100 misclass)
            h_amt = sub.groupby("horizontal_label")["_abs_amt"].sum()
            h_amt_pct = (h_amt / max(h_amt.sum(), 1)).sort_values(ascending=False)
            top_h_amt = h_amt_pct.head(3)
            rows.append({
                "account_code": acc,
                "account_desc": str(adesc)[:40],
                "rows": len(sub),
                "amount": int(sub["_abs_amt"].sum()),
                "top_h_by_count":  top_h_cnt.index[0] if len(top_h_cnt) > 0 else "",
                "count_pct":       round(top_h_cnt.iloc[0] * 100, 0) if len(top_h_cnt) > 0 else 0,
                "top_h_by_amount": top_h_amt.index[0] if len(top_h_amt) > 0 else "",
                "amount_pct":      round(top_h_amt.iloc[0] * 100, 0) if len(top_h_amt) > 0 else 0,
                "h_count_dist":  ", ".join(f"{lbl}({pct*100:.0f}%)" for lbl, pct in top_h_cnt.items()),
                "h_amount_dist": ", ".join(f"{lbl}({pct*100:.0f}%)" for lbl, pct in top_h_amt.items()),
            })
        acc_df = pd.DataFrame(rows)
        # Flags: LOW = top horizontal owns < 80% (LLM unstable for that account)
        # MISMATCH = count-majority horizontal != amount-majority (small rows tagged differently from large rows)
        acc_df["flag_count"]  = acc_df["count_pct"].apply(lambda p: "LOW" if p < 80 else "ok")
        acc_df["flag_amount"] = acc_df["amount_pct"].apply(lambda p: "LOW" if p < 80 else "ok")
        acc_df["flag_count_vs_amt"] = (
            acc_df["top_h_by_count"].ne(acc_df["top_h_by_amount"])
        ).map({True: "MISMATCH", False: "ok"})
        cols_show = ["account_code", "account_desc", "rows", "amount",
                     "top_h_by_count", "count_pct", "top_h_by_amount", "amount_pct",
                     "flag_count", "flag_amount", "flag_count_vs_amt"]
        print(acc_df[cols_show].to_string(index=False))
        acc_df.to_excel(writer, sheet_name="2_account_coherence", index=False)

    # ============ 3. Keyword consistency ============
    _print_header("3. Keyword vs horizontal consistency")
    desc_field = cols.get("description", "")
    kw_rows = []
    if desc_field and desc_field in df.columns:
        desc = df[desc_field].astype("string").fillna("")
        for kw, expected in KEYWORD_TESTS.items():
            mask = desc.str.contains(kw, case=False, na=False, regex=False)
            n_hit = int(mask.sum())
            if n_hit == 0:
                continue
            h_id = df.loc[mask, "horizontal_id"]
            pct_match = (h_id == expected).sum() / n_hit * 100
            top3 = h_id.value_counts().head(3)
            kw_rows.append({
                "keyword": kw,
                "expected": expected,
                "n_rows": n_hit,
                "amount": int(df.loc[mask, "_abs_amt"].sum()),
                "pct_correct": round(pct_match, 1),
                "top_actual": ", ".join(f"{l}({c})" for l, c in top3.items()),
                "flag": "LOW" if pct_match < 60 else "ok",
            })
        kw_df = pd.DataFrame(kw_rows)
        if len(kw_df):
            print(kw_df.to_string(index=False))
            kw_df.to_excel(writer, sheet_name="3_keyword_consistency", index=False)

    # ============ 4. NG-category x horizontal ============
    _print_header("4. NG-category x horizontal (top 3 per NG, by amount)")
    ng_col = cols.get("ng11_category", "")
    if ng_col and ng_col in df.columns:
        ng_h = pd.crosstab(df[ng_col], df["horizontal_label"], values=df["_abs_amt"], aggfunc="sum").fillna(0)
        ng_h_int = ng_h.astype("int64")
        ng_h_int.to_excel(writer, sheet_name="4_NG_x_horizontal_amount")
        for ng in sorted(ng_h.index):
            row = ng_h.loc[ng]
            top = row[row > 0].nlargest(3)
            tot = row.sum()
            if tot == 0:
                continue
            line = f"  {ng:5}  total={int(tot):>15,}  | "
            line += ", ".join(f"{lbl}({pct*100:.0f}%)" for lbl, pct in (top / tot).items())
            print(line)

    # ============ 5. V x H matrix + suspicious cells ============
    _print_header("5. V x H amount matrix saved to xlsx (sheet '5_VxH_amount')")
    vh = pd.crosstab(df["vertical_label"], df["horizontal_label"],
                     values=df["_abs_amt"], aggfunc="sum").fillna(0)
    vh.astype("int64").to_excel(writer, sheet_name="5_VxH_amount")
    # flag suspicious cells: V_GAMING_* x non-construction-related horizontal
    susp = []
    gaming_v = [v["label"] for v in cats["verticals"] if v.get("scope") == "NG0"]
    construction_h = {"建設與設施支出"}
    for v_lbl in gaming_v:
        if v_lbl not in vh.index:
            continue
        for h_lbl, amt_val in vh.loc[v_lbl].items():
            if amt_val > 0 and h_lbl not in construction_h and h_lbl != "":
                susp.append({"vertical": v_lbl, "horizontal": h_lbl, "amount": int(amt_val)})
    if susp:
        sdf = pd.DataFrame(susp).sort_values("amount", ascending=False)
        print("  Suspicious gaming-vertical x non-construction-horizontal cells (>0 amount):")
        print(sdf.to_string(index=False))
        sdf.to_excel(writer, sheet_name="5b_suspicious_VxH", index=False)

    # ============ 5c. V x H concentration: each vertical's top H share ============
    _print_header("5c. Vertical -> horizontal concentration "
                  "(top horizontal share of vertical's amount; <70% = LLM diffuse / unstable)")
    conc = []
    for v_lbl in vh.index:
        row = vh.loc[v_lbl]
        tot = row.sum()
        if tot == 0:
            continue
        pos = row[row > 0].sort_values(ascending=False)
        if len(pos) == 0:
            continue
        conc.append({
            "vertical": v_lbl,
            "total_amount":     int(tot),
            "top_horizontal":   pos.index[0],
            "top_pct":          round(pos.iloc[0] / tot * 100, 1),
            "second_horizontal": pos.index[1] if len(pos) > 1 else "",
            "second_pct":       round(pos.iloc[1] / tot * 100, 1) if len(pos) > 1 else 0.0,
            "n_horizontals":    int((row > 0).sum()),
            "flag":             "DIFFUSE" if pos.iloc[0] / tot < 0.7 else "ok",
        })
    if conc:
        conc_df = pd.DataFrame(conc).sort_values("total_amount", ascending=False)
        print(conc_df.to_string(index=False))
        conc_df.to_excel(writer, sheet_name="5c_VxH_concentration", index=False)

    # ============ 6. Confidence bucket x amount ============
    _print_header("6. Confidence bucket x signature count + amount")
    if "llm_confidence" in sigs.columns and "total_amount" in sigs.columns:
        conf = pd.to_numeric(sigs["llm_confidence"], errors="coerce")
        s_amt = pd.to_numeric(sigs["total_amount"], errors="coerce").abs()
        buckets = pd.cut(conf.fillna(-1), bins=[-1.01, -0.001, 0.5, 0.7, 0.9, 1.01],
                         labels=["null", "<0.5", "0.5-0.7", "0.7-0.9", ">=0.9"])
        stat = pd.DataFrame({
            "n_signatures": buckets.value_counts().sort_index(),
            "total_amount": s_amt.groupby(buckets).sum().round(0).astype("int64"),
        })
        stat["amt_pct"] = (stat["total_amount"] / max(s_amt.sum(), 1) * 100).round(1)
        print(stat.to_string())
        stat.to_excel(writer, sheet_name="6_confidence_buckets")

    # ============ 7. V_OTHER / H_OTHER share ============
    _print_header("7. '其他' fallback rate")
    v_other_amt = df.loc[df["vertical_id"].eq("V_OTHER"), "_abs_amt"].sum()
    h_other_amt = df.loc[df["horizontal_id"].eq("H_OTHER"), "_abs_amt"].sum()
    v_other_n = int(df["vertical_id"].eq("V_OTHER").sum())
    h_other_n = int(df["horizontal_id"].eq("H_OTHER").sum())
    print(f"  V_OTHER: rows={v_other_n:>8,}  amount={int(v_other_amt):>15,} ({v_other_amt/total_amt*100:.1f}% of total)")
    print(f"  H_OTHER: rows={h_other_n:>8,}  amount={int(h_other_amt):>15,} ({h_other_amt/total_amt*100:.1f}% of total)")
    if v_other_amt / total_amt > 0.05:
        print("  ⚠  V_OTHER >5%: classification fallback may be over-firing")
    if h_other_amt / total_amt > 0.05:
        print("  ⚠  H_OTHER >5%: classification fallback may be over-firing")

    # ============ 9. Capex/Opex x horizontal sanity ============
    _print_header("9. Capex/Opex x horizontal sanity "
                  "(each horizontal has an expected dominant capex_opex)")
    EXPECTED_DOMINANT = {  # horizontal_label -> "Capex" or "Opex"
        "建設與設施支出": "Capex",
        "人工成本":        "Opex",
        "酒店客房":        "Opex",
        "餐飲":            "Opex",
        "活動場地":        "Opex",
        "贊助費":          "Opex",
        "專業服務費":       "Opex",
    }
    ch = pd.crosstab(df["horizontal_label"], df["final_capex_opex"],
                     values=df["_abs_amt"], aggfunc="sum").fillna(0)
    ch_rows = []
    for h_lbl, expected in EXPECTED_DOMINANT.items():
        if h_lbl not in ch.index:
            continue
        row = ch.loc[h_lbl]
        cap = float(row.get("Capex", 0))
        opex = float(row.get("Opex", 0))
        tot = cap + opex
        if tot == 0:
            continue
        actual = "Capex" if cap > opex else "Opex"
        cap_pct = round(cap / tot * 100, 1)
        ch_rows.append({
            "horizontal":        h_lbl,
            "expected_dominant": expected,
            "actual_dominant":   actual,
            "capex_pct":         cap_pct,
            "opex_pct":          round(100 - cap_pct, 1),
            "capex_amount":      int(cap),
            "opex_amount":       int(opex),
            "flag":              "ok" if actual == expected else "MISMATCH",
        })
    if ch_rows:
        ch_df = pd.DataFrame(ch_rows)
        print(ch_df.to_string(index=False))
        ch_df.to_excel(writer, sheet_name="9_capex_horiz_sanity", index=False)

    # ============ 10. NG <-> vertical eligible reconciliation ============
    _print_header("10. NG <-> vertical eligible reconciliation "
                  "(every row's V should be in NG.eligible_verticals)")
    ng_to_eligible: dict[str, set[str]] = {}
    for ng, info in (cats.get("ng_categories") or {}).items():
        elig = set(info.get("eligible_verticals") or []) | {"V_OTHER"}
        ng_to_eligible[ng] = elig
        # Also key by Chinese label and by any "starts with 博彩" gaming variant
        if info.get("label"):
            ng_to_eligible[info["label"]] = elig
    # Catch gaming-vertical labels used as NG (e.g. VML/SJM: "博彩設施及設備的優化")
    _gaming_elig = ng_to_eligible.get("NG0", set())
    _all_keys = list(ng_to_eligible.keys())
    for _k in _all_keys:
        if str(_k).startswith("博彩") and _k not in ng_to_eligible:
            ng_to_eligible[_k] = _gaming_elig
    if ng_col and ng_col in df.columns:
        ng_str_series = df[ng_col].astype("string").fillna("")
        v_id_series   = df["vertical_id"].astype("string").fillna("")
        # Build (ng, v) eligibility lookup once over unique pairs (fast for 624k rows)
        pair_df = pd.DataFrame({"ng": ng_str_series.values, "v": v_id_series.values})
        unique_pairs = pair_df.drop_duplicates().copy()
        unique_pairs["eligible"] = [
            (not v) or (v in ng_to_eligible.get(ng, set()))
            for ng, v in zip(unique_pairs["ng"], unique_pairs["v"])
        ]
        elig_lookup = dict(zip(zip(unique_pairs["ng"], unique_pairs["v"]),
                                unique_pairs["eligible"]))
        is_elig = pd.Series(
            [elig_lookup[(ng, v)] for ng, v in zip(pair_df["ng"], pair_df["v"])],
            index=df.index,
        )
        n_violations = int((~is_elig).sum())
        amt_violations = float(df.loc[~is_elig, "_abs_amt"].sum())
        pct_v = (n_violations / len(df)) * 100 if len(df) else 0
        pct_a = (amt_violations / total_amt) * 100 if total_amt > 0 else 0
        print(f"  rows where vertical is NOT in NG.eligible_verticals: "
              f"{n_violations:,}/{len(df):,} ({pct_v:.2f}%)  "
              f"amount: {int(amt_violations):,} ({pct_a:.2f}% of total)  (expected ~0)")
        if n_violations > 0:
            bad = df.loc[~is_elig, [ng_col, "vertical_id", "vertical_label"]].copy()
            bad["_abs_amt"] = df.loc[~is_elig, "_abs_amt"].values
            bad_sum = bad.groupby(
                [ng_col, "vertical_id", "vertical_label"]
            ).agg(n_rows=("_abs_amt", "size"),
                  amount=("_abs_amt", "sum")).reset_index()
            bad_sum = bad_sum.sort_values("amount", ascending=False)
            bad_sum["amount"] = bad_sum["amount"].astype("int64")
            print("  Top violation (NG, vertical) pairs by amount:")
            print(bad_sum.head(15).to_string(index=False))
            bad_sum.to_excel(writer, sheet_name="10_ng_vertical_violations", index=False)

    # ============ 8. Random sample ============
    _print_header("8. Random 50-row sample written (sheet '8_random_sample')")
    sample_n = min(50, len(df))
    sample = df.sample(n=sample_n, random_state=42)
    keep_cols = [c for c in [
        cols.get("account_code", ""), cols.get("account_desc", ""),
        cols.get("description", ""), cols.get("amount", ""),
        cols.get("ng11_category", ""), cols.get("capex_opex", ""),
        "vertical_label", "horizontal_label", "vertical_source", "horizontal_source",
    ] if c and c in sample.columns]
    sample[keep_cols].to_excel(writer, sheet_name="8_random_sample", index=False)

    writer.close()
    print(f"\n  Wrote {out_path}", flush=True)
    print(f"  Total |amount| reconciled: {total_amt:,.0f}", flush=True)
