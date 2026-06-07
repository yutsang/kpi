"""MGM 投資方向 — GOLDEN-DRIVEN build (ties to project-team golden by construction).

Why this exists (vs the normal build_master_audit path):
  MGM's f1 CAPEX raw mixes gaming + non-gaming + CIP clearing (−258M) + multi-year Budget
  Sources, and ~1.13B of capex has no Project Plan Task — so reproducing the golden
  non-博彩 capex (~1.097B) bottom-up from raw is a rabbit hole. The project-team golden
  (results/mgm_golden_25.tsv) already states the AUTHORITATIVE capex / opex / payroll per
  項目. So we INVERT: take the golden amounts as truth (total ties exactly), and use the
  pipeline ONLY for the two things the golden does not give us:
      1. each 項目's V (→ NG bucket)         — amount-weighted vertical_id of its pipeline rows
      2. how to SPLIT capex & opex across H  — the pipeline's H proportions for that 項目
  Payroll → H_LABOR (人工) wholesale.

  Result: grand total + per-項目 capex/opex/payroll == golden (100% tie), WITH an H breakdown
  and a V/NG bucket per 項目. Where a 項目 has no pipeline rows, its H falls back (all-rows
  proportion, else H_OTHER) and is flagged in the recon sheet so it can be eyeballed.

Golden TSV layout (tab-separated, no header; 序號 rows are all-digits):
    序號 | 名稱 | payroll | capex | opex | total       (amounts in 萬元 → ×10000 = MOP)

Run ON THE WINDOWS BOX (where data/mgm/ + results/mgm_golden_25.tsv live):
  python scripts/build_mgm_golden.py
  python scripts/build_mgm_golden.py --golden results/mgm_golden_25.tsv --year 25
"""
from __future__ import annotations
import argparse, re, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd
import yaml

ROOT = Path(__file__).resolve().parent.parent

OPEX_WD = {"Gaming_OPEX", "WD1", "WD2", "WD4", "WD5_Patron"}   # WD3 = payroll, CAPEX = capex

V_TO_NG = {
    "V_GAMING_VENUE": ("NG0", "博彩項目"), "V_GAMING_EQUIP": ("NG0", "博彩項目"),
    "V_OVERSEAS_OFFICE": ("NG1", "吸引外國客源"), "V_OVERSEAS_WEB_SEO": ("NG1", "吸引外國客源"),
    "V_OVERSEAS_ROADSHOW": ("NG1", "吸引外國客源"), "V_INVITE_GUEST": ("NG1", "吸引外國客源"),
    "V_INVITE_AGENCY": ("NG1", "吸引外國客源"), "V_REGIONAL_TEAM": ("NG1", "吸引外國客源"),
    "V_REGIONAL_SALES": ("NG1", "吸引外國客源"), "V_PROMO_VIDEO": ("NG1", "吸引外國客源"),
    "V_MICE": ("NG2", "會議展覽"), "V_CONCERT": ("NG3", "娛樂表演"),
    "V_SPORT_EVENT": ("NG4", "體育盛事"), "V_VENUE_PERF_SPORT_MICE": ("NG4", "體育盛事"),
    "V_ART_EXHIBITION": ("NG5", "文化藝術"), "V_MUSEUM": ("NG5", "文化藝術"),
    "V_WELLNESS": ("NG6", "健康養生"), "V_THEME_PARK": ("NG7", "主題遊樂"),
    "V_RESTAURANT": ("NG8", "美食之都"), "V_FOOD_EVENT": ("NG8", "美食之都"),
    "V_COMMUNITY": ("NG9", "社區旅遊"), "V_MARITIME": ("NG10", "海上旅遊"),
    "V_PROPERTY_UPGRADE": ("NG11", "其他"), "V_OTHER": ("NG11", "其他"),
}

# Fallback V by project-name keyword when a 項目 has no usable pipeline V.
NAME2V = [
    ("美食", "V_RESTAURANT"), ("餐飲", "V_RESTAURANT"), ("食肆", "V_RESTAURANT"),
    ("藝術", "V_ART_EXHIBITION"), ("美術", "V_ART_EXHIBITION"), ("展覽", "V_MICE"),
    ("會議", "V_MICE"), ("會展", "V_MICE"), ("表演", "V_CONCERT"), ("演出", "V_CONCERT"),
    ("演唱", "V_CONCERT"), ("體育", "V_SPORT_EVENT"), ("賽事", "V_SPORT_EVENT"),
    ("主題", "V_THEME_PARK"), ("樂園", "V_THEME_PARK"), ("遊樂", "V_THEME_PARK"),
    ("養生", "V_WELLNESS"), ("健康", "V_WELLNESS"), ("社區", "V_COMMUNITY"),
    ("海上", "V_MARITIME"), ("遊艇", "V_MARITIME"),
]


def _num(x):
    x = str(x).replace(",", "").strip()
    try:
        return float(x) * 10000.0           # golden in 萬元
    except Exception:
        return 0.0


def parse_golden(p: Path):
    out = {}
    for line in p.read_text(encoding="utf-8-sig").splitlines():
        parts = [c.strip() for c in line.split("\t")]
        if len(parts) < 6:
            continue
        sn = parts[0].strip()
        if not re.fullmatch(r"\d+", sn):
            continue
        out[sn.zfill(3)] = {"name": parts[1], "payroll": _num(parts[2]),
                            "capex": _num(parts[3]), "opex": _num(parts[4]), "total": _num(parts[5]),
                            "theme": parts[6].strip() if len(parts) >= 7 else ""}
    return out


# 項目性質 (project-team theme, golden TSV col 7) → our V. The Leadsheet states each 序號's theme
# directly, so V/NG come from it — NOT from the pipeline's amount-weighted vertical_id, which
# collapses to V_OTHER whenever the row-level NG is blank (mgm CAPEX/payroll have no NG). Same
# 13-value DICJ taxonomy as wynn 項目性質 / 項目分類2.
THEME2V = {
    "博彩娛樂場場地的優化": "V_GAMING_VENUE", "博彩設施及設備的優化": "V_GAMING_EQUIP",
    "博彩娛樂": "V_GAMING_VENUE", "博彩": "V_GAMING_VENUE", "博彩項目": "V_GAMING_VENUE",
    "吸引外國客源": "V_INVITE_GUEST", "會議展覽": "V_MICE", "娛樂表演": "V_CONCERT",
    "體育盛事": "V_SPORT_EVENT", "體育賽事": "V_SPORT_EVENT",
    "文化藝術": "V_ART_EXHIBITION", "健康養生": "V_WELLNESS", "主題遊樂": "V_THEME_PARK",
    "美食之都": "V_RESTAURANT", "社區旅遊": "V_COMMUNITY", "海上旅遊": "V_MARITIME", "其他": "V_OTHER",
}


def _serial_token(pid, nm):
    m = re.fullmatch(r"0*(\d+)", str(pid).strip())
    if m:
        return m.group(1).zfill(3)
    m = re.search(r"項目\s*0*(\d+)", str(nm))
    if m:
        return m.group(1).zfill(3)
    return ""


def _norm_nm(s):
    s = re.sub(r"項目\s*\d+\s*", "", str(s))
    s = re.sub(r"[－—–−]", "-", s)
    return s.replace(" ", "").replace("　", "")


def make_resolver(golden):
    name2sn = sorted(((_norm_nm(g["name"]), sn) for sn, g in golden.items()
                      if g.get("name") and len(g["name"]) >= 4), key=lambda x: -len(x[0]))
    cache = {}

    def resolve(pid, nm):
        s = _serial_token(pid, nm)
        if s and s in golden:
            return s
        key = _norm_nm(nm)
        if key in cache:
            return cache[key]
        sn = ""
        for gname, gsn in name2sn:
            if gname and gname in key:
                sn = gsn; break
        cache[key] = sn
        return sn
    return resolve


def _name_v(nm):
    for kw, v in NAME2V:
        if kw in str(nm):
            return v
    return "V_OTHER"


def _h_props(sub_amt_by_h):
    """{H: |amount|} → {H: proportion} (drop blanks/zeros). Empty in → {}."""
    tot = sum(v for v in sub_amt_by_h.values() if v > 0)
    if tot <= 0:
        return {}
    return {h: v / tot for h, v in sub_amt_by_h.items() if v > 0}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--golden", default="results/mgm_golden_25.tsv")
    ap.add_argument("--year", default="25")
    args = ap.parse_args()

    gp = ROOT / args.golden
    if not gp.exists():
        print(f"X golden {gp} missing"); return
    golden = parse_golden(gp)
    g_cap = sum(g["capex"] for g in golden.values())
    g_opx = sum(g["opex"] for g in golden.values())
    g_pay = sum(g["payroll"] for g in golden.values())
    print(f"golden: {len(golden)} 項目  capex={g_cap:,.0f}  opex={g_opx:,.0f}  payroll={g_pay:,.0f}  "
          f"total={g_cap+g_opx+g_pay:,.0f}")

    cats = yaml.safe_load((ROOT / "conf/base/categories.yml").read_text(encoding="utf-8"))
    hlab = {h["id"]: h.get("label", h["id"]) for h in cats.get("horizontals", [])}
    vlab = {v["id"]: v.get("label", v["id"]) for v in cats.get("verticals", [])}
    h_order = [h["label"] for h in cats.get("horizontals", []) if h.get("id") != "H_COUNT" and h.get("label")]

    pq = ROOT / "data/mgm/output/company_6_kpi_report.parquet"
    have_pipe = pq.exists()
    by_sn = {}            # sn -> {"vid_amt": {vid: amt}, "cap_h": {h: amt}, "opx_h": {h: amt}}
    p_sums = {}           # sn -> {"cap","opx","pay"} pipeline SIGNED sums (for golden-vs-pipeline Δ)
    je_detail = None      # row-level pipeline rows (for tracing WHY pipeline ≠ golden)
    if have_pipe:
        cfg = yaml.safe_load((ROOT / "conf/company_6/parameters.yml").read_text(encoding="utf-8"))
        cols = cfg.get("columns", {})
        df = pd.read_parquet(pq)
        yc = next((c for c in ("report_period", "report_year", "years") if c in df.columns), None)
        if yc:
            df = df[df[yc].astype(str).str.startswith(args.year)].copy()
        wd_col = next((c for c in (cols.get("capex_opex"), "wd", "final_capex_opex") if c and c in df.columns), None)
        amt_col = cols.get("amount", "amount_mop")
        proj_col = cols.get("project", "project_name")
        resolve = make_resolver(golden)
        pid = df["project_id"].astype(str) if "project_id" in df.columns else pd.Series([""] * len(df), index=df.index)
        nm = df[proj_col].astype(str) if proj_col in df.columns else pd.Series([""] * len(df), index=df.index)
        df["_sn"] = [resolve(a, b) for a, b in zip(pid, nm)]
        df["_wd"] = df[wd_col].astype(str).str.strip() if wd_col else ""
        _amt_signed = pd.to_numeric(df[amt_col], errors="coerce").fillna(0)
        df["_amt_signed"] = _amt_signed
        df["_amt"] = _amt_signed.abs()
        df["_vid"] = df["vertical_id"].astype(str) if "vertical_id" in df.columns else "V_OTHER"
        df["_hid"] = df["horizontal_id"].astype(str) if "horizontal_id" in df.columns else "H_OTHER"
        # per-row bucket (Capex / Opex / Payroll / Other) from the wd column
        _bk = pd.Series("Other", index=df.index)
        _bk = _bk.mask(df["_wd"].isin(OPEX_WD), "Opex")
        _bk = _bk.mask(df["_wd"].eq("WD3"), "Payroll")
        _bk = _bk.mask(df["_wd"].eq("CAPEX"), "Capex")
        df["_bucket"] = _bk

        matched = df[df["_sn"].astype(bool)]                  # rows that mapped to a 序號

        def _nest(gb):
            out = {}
            for (sn, k), a in gb.items():
                out.setdefault(sn, {})[k] = float(a)
            return out
        v_nest = _nest(matched.groupby(["_sn", "_vid"])["_amt"].sum())
        cap_nest = _nest(matched[matched["_wd"].eq("CAPEX")].groupby(["_sn", "_hid"])["_amt"].sum())
        opx_nest = _nest(matched[matched["_wd"].isin(OPEX_WD)].groupby(["_sn", "_hid"])["_amt"].sum())
        for sn in set(v_nest) | set(cap_nest) | set(opx_nest):
            by_sn[sn] = {"vid_amt": v_nest.get(sn, {}),
                         "cap_h": cap_nest.get(sn, {}), "opx_h": opx_nest.get(sn, {})}

        # pipeline SIGNED sums per 序號 (golden-vs-pipeline Δ shows where取數 over/under)
        _cap_s = matched.loc[matched["_wd"].eq("CAPEX")].groupby("_sn")["_amt_signed"].sum()
        _opx_s = matched.loc[matched["_wd"].isin(OPEX_WD)].groupby("_sn")["_amt_signed"].sum()
        _pay_s = matched.loc[matched["_wd"].eq("WD3")].groupby("_sn")["_amt_signed"].sum()
        for sn in set(_cap_s.index) | set(_opx_s.index) | set(_pay_s.index):
            p_sums[sn] = {"cap": float(_cap_s.get(sn, 0.0)), "opx": float(_opx_s.get(sn, 0.0)),
                          "pay": float(_pay_s.get(sn, 0.0))}

        # row-level JE detail (ALL rows incl 未對應) — trace the差異 down to the ledger line
        def _cv(name):
            return df[name].astype(str).values if name and name in df.columns else [""] * len(df)
        je_detail = pd.DataFrame({
            "序號": df["_sn"].replace("", "(未對應)").values,
            "項目": _cv(proj_col),
            "bucket": df["_bucket"].values,
            "wd": df["_wd"].values,
            "account_code": _cv(cols.get("account_code", "")),
            "account_desc": _cv(cols.get("account_desc", "")),
            "description": _cv(cols.get("description", "")),
            "vertical_id": df["_vid"].values,
            "horizontal_id": df["_hid"].values,
            "amount_mop": df["_amt_signed"].round().values,
        })
        je_detail["_abs"] = je_detail["amount_mop"].abs()
        je_detail = je_detail.sort_values(["序號", "bucket", "_abs"],
                                          ascending=[True, True, False]).drop(columns="_abs")
        _nmatch = int(df["_sn"].astype(bool).sum())
        print(f"pipeline: matched {len(by_sn)}/{len(golden)} 項目; JE rows={len(je_detail):,} "
              f"(對應={_nmatch:,}, 未對應={len(je_detail)-_nmatch:,})")
    else:
        print(f"! {pq} missing — V falls back to project-name keyword, H to H_OTHER (run kedro mgm for finer split)")

    # ── Allocate golden amounts using pipeline V + H proportions ──────────────────
    detail = []     # (序號, 項目, NG, V, V_label, bucket, H, H_label, amount, h_source)
    recon = []      # per 項目 tie + how V/H were sourced
    _unmapped_themes = {}   # theme not in THEME2V → {theme: count} (so we catch byte-mismatches)
    for sn, g in sorted(golden.items()):
        d = by_sn.get(sn, {})
        theme = str(g.get("theme", "")).strip()
        vid_amt = d.get("vid_amt", {})
        # V (→ NG): PREFER the Leadsheet 項目性質 theme (authoritative per 序號); only fall back to the
        # pipeline's amount-weighted vertical_id when there is no usable theme. The pipeline V collapses
        # to V_OTHER for mgm CAPEX/payroll (blank row-level NG), so theme-first avoids that corruption.
        if theme in THEME2V:
            v = THEME2V[theme]; vsrc = "leadsheet-theme"
        elif vid_amt:
            if theme:
                _unmapped_themes[theme] = _unmapped_themes.get(theme, 0) + 1
            non_g = {v2: a for v2, a in vid_amt.items() if V_TO_NG.get(v2, ("NG11",))[0] != "NG0"}
            pick = max((non_g or vid_amt).items(), key=lambda kv: kv[1])[0]
            v = pick if pick and pick != "nan" else _name_v(g["name"])
            vsrc = "pipeline" if non_g else "pipeline(gaming-only→kept)"
        else:
            if theme:
                _unmapped_themes[theme] = _unmapped_themes.get(theme, 0) + 1
            v = _name_v(g["name"]); vsrc = "name-keyword" if v != "V_OTHER" else "default"
        ng, ng_lab = V_TO_NG.get(v, ("NG11", "其他"))

        cap_props = _h_props(d.get("cap_h", {}))
        opx_props = _h_props(d.get("opx_h", {}))
        # capex fallback chain: capex-row H prop → all-row H prop → H_CONSTRUCTION
        cap_src = "capex-rows"
        if not cap_props:
            allh = {}
            for src in ("cap_h", "opx_h"):
                for h, a in d.get(src, {}).items():
                    allh[h] = allh.get(h, 0.0) + a
            cap_props = _h_props(allh)
            cap_src = "all-rows" if cap_props else "default→建設"
            if not cap_props:
                cap_props = {"H_CONSTRUCTION": 1.0}
        opx_src = "opex-rows"
        if not opx_props:
            opx_props = {"H_OTHER": 1.0}; opx_src = "default→其他"

        def _emit(bucket, amount, props):
            for h, p in props.items():
                detail.append([sn, g["name"], ng, ng_lab, v, vlab.get(v, v),
                               bucket, h, hlab.get(h, h), round(amount * p)])

        if g["capex"]:
            _emit("Capex", g["capex"], cap_props)
        if g["opex"]:
            _emit("Opex", g["opex"], opx_props)
        if g["payroll"]:
            detail.append([sn, g["name"], ng, ng_lab, v, vlab.get(v, v),
                           "Payroll", "H_LABOR", hlab.get("H_LABOR", "人工"), round(g["payroll"])])

        ps = p_sums.get(sn, {})
        p_cap, p_opx, p_pay = ps.get("cap", 0.0), ps.get("opx", 0.0), ps.get("pay", 0.0)
        recon.append([sn, g["name"][:34], ng, vlab.get(v, v), vsrc,
                      round(g["capex"]), round(p_cap), round(p_cap - g["capex"]),
                      round(g["opex"]), round(p_opx), round(p_opx - g["opex"]),
                      round(g["payroll"]), round(p_pay), round(p_pay - g["payroll"]),
                      round(g["total"]), cap_src, opx_src])

    det = pd.DataFrame(detail, columns=["序號", "項目", "ng_code", "ng_label", "vertical_id",
                                        "vertical_label", "bucket", "horizontal_id",
                                        "horizontal_label", "amount_mop"])
    # G_* = golden (truth, used for the deliverable). P_* = pipeline bottom-up sum. Δ = P − G:
    # negative = pipeline 取數不足 for that 項目 (the gap to drill via 7_JE明細).
    rec = pd.DataFrame(recon, columns=["序號", "項目", "ng_code", "vertical_label", "V來源",
                                       "G_capex", "P_capex", "Δcapex",
                                       "G_opex", "P_opex", "Δopex",
                                       "G_payroll", "P_payroll", "Δpayroll",
                                       "G_total", "capex_H來源", "opex_H來源"])

    # pivot V × H (NG header) over the allocated amounts
    pv = det.pivot_table(index=["ng_code", "ng_label", "vertical_id", "vertical_label"],
                         columns="horizontal_label", values="amount_mop",
                         aggfunc="sum", fill_value=0, margins=True, margins_name="總計")
    extra = [c for c in pv.columns if c not in h_order and c != "總計"]
    pv = pv.reindex(columns=h_order + extra + (["總計"] if "總計" in pv.columns else []), fill_value=0)

    h_by = det.groupby(["bucket", "horizontal_label"])["amount_mop"].sum().reset_index() \
              .sort_values(["bucket", "amount_mop"], ascending=[True, False])
    v_by = det.groupby(["ng_code", "ng_label", "vertical_label", "bucket"])["amount_mop"].sum() \
              .reset_index().sort_values(["ng_code", "vertical_label", "bucket"])

    a_cap = det.loc[det.bucket.eq("Capex"), "amount_mop"].sum()
    a_opx = det.loc[det.bucket.eq("Opex"), "amount_mop"].sum()
    a_pay = det.loc[det.bucket.eq("Payroll"), "amount_mop"].sum()
    summ = pd.DataFrame({
        "metric": ["golden_capex", "alloc_capex", "golden_opex", "alloc_opex",
                   "golden_payroll", "alloc_payroll", "golden_total", "alloc_total"],
        "amount": [round(g_cap), round(a_cap), round(g_opx), round(a_opx),
                   round(g_pay), round(a_pay), round(g_cap + g_opx + g_pay),
                   round(a_cap + a_opx + a_pay)],
    })

    out_dir = ROOT / "data/review"; out_dir.mkdir(parents=True, exist_ok=True)
    out = out_dir / f"mgm_投資方向_{args.year}_golden.xlsx"
    with pd.ExcelWriter(out, engine="xlsxwriter") as w:
        pd.DataFrame([
            ("0_index", "this map"),
            ("1_投資方向pivot", "V × H cross-tab of GOLDEN amounts, NG header + 總計 (capex+opex+payroll)"),
            ("2_recon_by_項目", "per 項目: G_*(golden, 用嚟出數) vs P_*(pipeline bottom-up sum) + Δ=P−G. Δ<0 = 取數不足 → 去 7_JE明細 drill"),
            ("3_橫向", "(bucket, H) × Σ golden amount"),
            ("4_縱向", "(NG, V, bucket) × Σ golden amount"),
            ("5_大表", "allocation detail: (序號, 項目, NG, V, bucket, H) × amount"),
            ("6_tie", "golden vs allocated grand totals — must match"),
            ("7_JE明細", "row-level pipeline JE rows per 序號 (incl 未對應) — WHY pipeline≠golden; 診斷用,唔係 tie 數字"),
        ], columns=["sheet", "contents"]).to_excel(w, sheet_name="0_index", index=False)
        pv.to_excel(w, sheet_name="1_投資方向pivot")
        rec.to_excel(w, sheet_name="2_recon_by_項目", index=False)
        h_by.to_excel(w, sheet_name="3_橫向", index=False)
        v_by.to_excel(w, sheet_name="4_縱向", index=False)
        det.to_excel(w, sheet_name="5_大表", index=False)
        summ.to_excel(w, sheet_name="6_tie", index=False)
        if je_detail is not None and len(je_detail):
            _XL = 1_048_574
            if len(je_detail) <= _XL:
                je_detail.to_excel(w, sheet_name="7_JE明細", index=False)
            else:
                _n = (len(je_detail) + _XL - 1) // _XL
                for _i in range(_n):
                    je_detail.iloc[_i*_XL:(_i+1)*_XL].to_excel(
                        w, sheet_name=f"7_JE明細_p{_i+1}of{_n}", index=False)

    print(f"\n=== tie: allocated (deliverable) must match golden ===")
    print(f"  capex   golden={g_cap:>15,.0f}   alloc={a_cap:>15,.0f}   Δ={a_cap-g_cap:>12,.0f}")
    print(f"  opex    golden={g_opx:>15,.0f}   alloc={a_opx:>15,.0f}   Δ={a_opx-g_opx:>12,.0f}")
    print(f"  payroll golden={g_pay:>15,.0f}   alloc={a_pay:>15,.0f}   Δ={a_pay-g_pay:>12,.0f}")
    nfb = (rec["capex_H來源"] != "capex-rows").sum()
    print(f"  {len(rec)} 項目 | capex H fallback (no pipeline capex rows): {nfb}")
    print("  V來源 breakdown:")
    for s, n in rec["V來源"].value_counts().items():
        print(f"    {s:28s} {n:>3} 項目")
    if _unmapped_themes:
        print(f"  ⚠️ {sum(_unmapped_themes.values())} 項目 theme NOT in THEME2V (fell back) — add byte-exact keys:")
        for t, n in sorted(_unmapped_themes.items(), key=lambda x: -x[1]):
            print(f"      {t!r}: {n}")

    # diagnostic: where the pipeline bottom-up UNDER-counts vs golden (drill via 7_JE明細)
    if p_sums:
        P_cap, P_opx, P_pay = rec["P_capex"].sum(), rec["P_opex"].sum(), rec["P_payroll"].sum()
        print(f"\n=== pipeline bottom-up vs golden (診斷 — 解釋差異) ===")
        print(f"  capex   pipeline={P_cap:>15,.0f}  golden={g_cap:>15,.0f}  Δ={P_cap-g_cap:>13,.0f}")
        print(f"  opex    pipeline={P_opx:>15,.0f}  golden={g_opx:>15,.0f}  Δ={P_opx-g_opx:>13,.0f}")
        print(f"  payroll pipeline={P_pay:>15,.0f}  golden={g_pay:>15,.0f}  Δ={P_pay-g_pay:>13,.0f}")
        rec["_dt"] = rec["P_capex"] + rec["P_opex"] + rec["P_payroll"] - rec["G_total"]
        worst = rec.reindex(rec["_dt"].abs().sort_values(ascending=False).index).head(8)
        print(f"\n  最大差異 8 項 (P_total − G_total):")
        for _, r in worst.iterrows():
            print(f"    {r['序號']}  Δ={r['_dt']:>14,.0f}  G={r['G_total']:>13,.0f}  {str(r['項目'])[:26]}")
    print(f"\n→ {out}  (0_index/1_pivot/2_recon(G/P/Δ)/3_橫向/4_縱向/5_大表/6_tie/7_JE明細)")


if __name__ == "__main__":
    main()
