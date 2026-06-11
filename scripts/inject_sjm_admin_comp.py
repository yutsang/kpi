"""Build SJM internal-resource (admin) comp rows for injection into 1_投資方向pivot.

SJM rows flagged 是否計入內部資源=Y (AN) are internal comp — excluded from the main JE
tagging and instead reconstructed here from 'Admin Comp summary v2.xlsx':

  • BKD_補充           — direct: V=項目類型, H=H_ADVERTISING (推廣費), subproject=CO object name,
                          amount=Val/COArea Crcy.  (first-batch comp not captured in opera)
  • combined admin comp — filter 是否包括在25年投资金額(KPMG)=Y ; H from 'Type_for report (KPMG)'
                          (F&B→H_FNB, Venue→H_VENUE, Hotel→H_HOTEL_ROOM, others→H_COMP_OTHER;
                           blank→TRX_DESC); amount from 'Amount'; V keyword-mapped from the
                          event/outlet (GUEST_FULL_NAME / Comp Type); subproject=event/outlet.

NOT proportional —每筆一行. Header is on row 2 of both sheets (header=1).
Outputs results/sjm_admincomp_rows.tsv (vertical_id, horizontal_id, subproject, amount, source)
and PRINTS combined events whose V could not be mapped (fill those, then we finalize).

Run:
  python scripts/inject_sjm_admin_comp.py --file "data/sjm/raw/Admin Comp summary v2.xlsx"
"""
from __future__ import annotations
import argparse, sys, csv, re
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

# project-team 項目類型 (13 themes) -> our V  (same map as SJM row_vertical, used for BKD)
THEME2V = {"社區旅遊": "V_COMMUNITY", "娛樂表演": "V_CONCERT", "會議展覽": "V_MICE",
           "博彩設施及設備的優化": "V_GAMING_EQUIP", "美食之都": "V_RESTAURANT",
           "博彩娛樂場場地的優化": "V_GAMING_VENUE", "體育盛事": "V_SPORT_EVENT",
           "海上旅遊": "V_MARITIME", "吸引外國客源": "V_OVERSEAS_ROADSHOW",
           "文化藝術": "V_ART_EXHIBITION", "主題遊樂": "V_THEME_PARK",
           "其他": "V_OTHER", "健康養生": "V_WELLNESS"}

# Type_for report (KPMG) -> H
TYPE2H = {"f&b": "H_FNB", "venue": "H_VENUE", "hotel": "H_HOTEL_ROOM",
          "room": "H_HOTEL_ROOM", "others": "H_COMP_OTHER", "other": "H_COMP_OTHER"}

# fallback when Type_for report blank: TRX_DESC keyword -> H
TRX2H = [("room", "H_HOTEL_ROOM"), ("food", "H_FNB"), ("beverage", "H_FNB"),
         ("service charge", "H_FNB"), ("tax", "H_HOTEL_ROOM"), ("rental", "H_VENUE"),
         ("ticket", "H_COMP_TICKET")]

# combined comp event/outlet keyword -> V  (seed; unmapped are printed for review)
EVENT2V = [("music", "V_CONCERT"), ("show", "V_CONCERT"), ("concert", "V_CONCERT"),
           ("演唱", "V_CONCERT"), ("tvb", "V_CONCERT"), ("award", "V_CONCERT"),
           ("buffet", "V_RESTAURANT"), ("restaurant", "V_RESTAURANT"), ("dining", "V_RESTAURANT"),
           ("aji", "V_RESTAURANT"), ("bar", "V_RESTAURANT"), ("chef", "V_RESTAURANT"),
           ("spa", "V_WELLNESS"), ("wellness", "V_WELLNESS"),
           ("sport", "V_SPORT_EVENT"), ("art", "V_ART_EXHIBITION"), ("mice", "V_MICE"),
           ("exhibition", "V_ART_EXHIBITION"), ("gourmet", "V_RESTAURANT")]

# SJM combined-comp guest-name rules (project-team-confirmed event types) — checked AFTER the
# 表1 Description 1-to-1 match but BEFORE the generic keyword map. Ordered, first-match-wins;
# specific phrases first so e.g. "Chefs' Table" → food (not the generic "chef" → restaurant).
SJM_GUEST_RULES = [
    # food events
    ("chefs’ table", "V_FOOD_EVENT"), ("chefs' table", "V_FOOD_EVENT"),
    ("chef’s table", "V_FOOD_EVENT"), ("chef's table", "V_FOOD_EVENT"),
    ("dim sum", "V_FOOD_EVENT"), ("gala dinner", "V_FOOD_EVENT"),
    ("whisky", "V_FOOD_EVENT"), ("wine & dine", "V_FOOD_EVENT"),
    # sport
    ("dragon boat", "V_SPORT_EVENT"), ("grand prix", "V_SPORT_EVENT"),
    ("china tennis", "V_SPORT_EVENT"), ("tennis", "V_SPORT_EVENT"),
    ("macao open", "V_SPORT_EVENT"), ("macau open", "V_SPORT_EVENT"),
    ("national game", "V_SPORT_EVENT"),
    ("cta 2024", "V_SPORT_EVENT"), ("cta -", "V_SPORT_EVENT"), ("cta-", "V_SPORT_EVENT"),
    # concert / performance
    ("music show", "V_CONCERT"), ("music concert", "V_CONCERT"),
    ("music award", "V_CONCERT"), ("the fact music", "V_CONCERT"), ("tma 2025", "V_CONCERT"),
    ("tvb", "V_CONCERT"), ("anniversary award", "V_CONCERT"),
    # culture / art
    ("culture city", "V_ART_EXHIBITION"), ("exhibition", "V_ART_EXHIBITION"),
    ("museum", "V_ART_EXHIBITION"), ("biennale", "V_ART_EXHIBITION"),
    # MICE / summit / conference (project team 77.x)
    ("sommelier", "V_MICE"), ("summit", "V_MICE"), ("conference", "V_MICE"),
    ("council meeting", "V_MICE"), ("forum", "V_MICE"), ("ectaa", "V_MICE"),
    # invite guest (media / KOL stays) — before trade/roadshow
    ("media stay", "V_INVITE_GUEST"), ("family media", "V_INVITE_GUEST"),
    ("media fam", "V_INVITE_GUEST"), ("kol", "V_INVITE_GUEST"),
    ("tailor made drama", "V_INVITE_GUEST"),
    # overseas roadshow / trade fam
    ("trade fam", "V_OVERSEAS_ROADSHOW"), ("trade fair", "V_OVERSEAS_ROADSHOW"),
    ("jata", "V_OVERSEAS_ROADSHOW"), ("matta", "V_OVERSEAS_ROADSHOW"),
    ("roadshow", "V_OVERSEAS_ROADSHOW"), ("experience tour", "V_OVERSEAS_ROADSHOW"),
    # overseas media / web
    ("creator week", "V_OVERSEAS_WEB_SEO"), ("robb report", "V_OVERSEAS_WEB_SEO"),
    ("media investment", "V_OVERSEAS_WEB_SEO"),
]


def _col(df, *needles):
    for c in df.columns:
        cl = str(c).lower()
        if all(n.lower() in cl for n in needles):
            return c
    return None


def _h_from_type(tfr, trx):
    h = TYPE2H.get(str(tfr).strip().lower())
    if h:
        return h
    s = str(trx).lower()
    for kw, hh in TRX2H:
        if kw in s:
            return hh
    return "H_COMP_OTHER"


def _v_from_event(*texts):
    s = " ".join(str(t) for t in texts).lower()
    for kw, v in EVENT2V:
        if kw in s:
            return v
    return None


def _v_from_sjm_guest(name):
    """SJM combined-comp guest-name → V via project-team-confirmed event rules (SJM_GUEST_RULES)."""
    s = str(name).lower()
    for kw, v in SJM_GUEST_RULES:
        if kw in s:
            return v
    return None


def build_guest_v_map(comb_df, c_guest, je_df, je_desc_col, je_vid_col="vertical_id"):
    """GUEST_FULL_NAME → vertical_id by substring-matching the guest inside 表1 (JE) Description.

    Project team's own method: each comp guest/event appears in 表1 Description, whose row carries
    the project's (LLM/manual) vertical_id. Adopt that V ONLY when the guest's matches resolve to
    exactly ONE distinct V (unambiguous); ambiguous / no-match guests are left for keyword / V_OTHER.
    Returns {guest: vid}."""
    if je_df is None or not je_desc_col or je_desc_col not in je_df.columns or je_vid_col not in je_df.columns:
        return {}
    jd = je_df[[je_desc_col, je_vid_col]].copy()
    jd[je_desc_col] = jd[je_desc_col].astype(str)
    jd = jd[jd[je_desc_col].str.strip().ne("") & jd[je_desc_col].str.lower().ne("nan")].drop_duplicates()
    if jd.empty:
        print(f"  [guest→V] 表1 description col {je_desc_col!r} is empty — skip match")
        return {}
    guests = sorted({str(g).strip() for g in comb_df[c_guest].dropna()
                     if str(g).strip() and str(g).strip().lower() != "nan"}, key=len, reverse=True)
    gv = {}
    for g in guests:
        if len(g) < 3:            # too short → unreliable substring
            continue
        try:
            hit = jd[jd[je_desc_col].str.contains(re.escape(g), case=False, na=False)]
        except Exception:
            continue
        vids = [v for v in hit[je_vid_col].astype(str).unique() if v and v != "nan"]
        if len(vids) == 1:
            gv[g] = vids[0]
    print(f"  [guest→V] {len(gv)}/{len(guests)} combined guests matched to a single V via 表1 Description")
    return gv


def build_rows(admin_comp_path="data/sjm/raw/Admin Comp summary v2.xlsx",
               je_df=None, je_desc_col=None, je_vid_col="vertical_id"):
    """Return (rows, unmapped). rows = [vertical_id, horizontal_id, subproject, amount, source].
    Importable by build_master_audit so the inject runs INSIDE the pipeline (step6 / generate),
    not just as a standalone script — so it's never forgotten after a kedro run.

    je_df (optional) = 表1 (JE tagged rows). When given, combined comp V is resolved via
    GUEST_FULL_NAME → Description substring match (project team ground truth) before keyword."""
    fp = Path(admin_comp_path)
    if not fp.exists():
        print(f"X {fp} not found — admin-comp inject skipped"); return [], {}
    xl = pd.ExcelFile(fp)
    rows = []
    unmapped = {}

    # ---- BKD_補充 ----  (header row varies — auto-detect by scanning for 項目類型; theme/amount value-fallback)
    bkd_name = next((s for s in xl.sheet_names if "bkd" in s.lower() or "補充" in s or "补充" in s), None)
    n_bkd = 0
    if not bkd_name:
        print(f"BKD_補充: sheet NOT found — sheets = {xl.sheet_names}")
    else:
        _probe = xl.parse(bkd_name, header=None, dtype=str, nrows=8)
        _hdr = 0
        for _i in range(len(_probe)):
            if any(("項目類型" in str(x) or "項目性質" in str(x)) for x in _probe.iloc[_i].tolist()):
                _hdr = _i; break
        b = xl.parse(bkd_name, header=_hdr, dtype=str)
        c_theme = _col(b, "項目類型") or _col(b, "项目类型") or _col(b, "項目性質")
        if not c_theme:                       # value-based: the column whose values best match THEME2V
            _best = 0
            for _c in b.columns:
                _n = b[_c].astype(str).str.strip().isin(THEME2V).sum()
                if _n > _best:
                    _best, c_theme = _n, _c
        c_co = _col(b, "co object"); c_wbs = _col(b, "wbs")
        c_amt = _col(b, "val", "crcy") or _col(b, "val") or _col(b, "amount")
        if not c_amt:                          # value-based: most-numeric, non-zero-sum column
            _best = 0
            for _c in b.columns:
                _num = pd.to_numeric(b[_c], errors="coerce")
                if _num.notna().sum() > _best and float(_num.abs().sum() or 0) > 0:
                    _best, c_amt = int(_num.notna().sum()), _c
        for _, r in b.iterrows():
            v = THEME2V.get(str(r.get(c_theme, "")).strip())
            if not v:
                continue
            amt = pd.to_numeric(r.get(c_amt), errors="coerce")
            if pd.isna(amt) or amt == 0:
                continue
            sub = " | ".join(x for x in [str(r.get(c_wbs, "")).strip(), str(r.get(c_co, "")).strip()] if x and x != "nan")
            rows.append([v, "H_ADVERTISING", sub, round(float(amt), 2), "BKD_補充"])
            n_bkd += 1
        _bkd_sum = sum(r[3] for r in rows if r[4] == "BKD_補充")
        print(f"BKD_補充: {n_bkd} rows  Σ={_bkd_sum:,.0f} (target ≈1,091,000)  "
              f"(sheet={bkd_name!r}, header row={_hdr}, 項目類型={c_theme!r}, amt={c_amt!r}, raw {len(b)} rows)")
        if n_bkd == 0:
            print(f"  [debug] BKD columns = {list(b.columns)[:25]}")
            if c_theme:
                vc = b[c_theme].astype(str).str.strip().value_counts().head(20)
                print("  [debug] distinct theme values (top 20) — must match THEME2V keys:")
                for _k, _n in vc.items():
                    print(f"     {_k!r:42} x{_n:<5} {'OK' if _k in THEME2V else 'NOT IN MAP'}")

    # ---- combined admin comp ----
    comb_name = next((s for s in xl.sheet_names if "combined" in s.lower()), None)
    if comb_name:
        d = xl.parse(comb_name, header=1, dtype=str)
        c_flag = _col(d, "包括在25") or _col(d, "25年")
        c_type = _col(d, "type_for report") or _col(d, "type_for")
        c_amt = _col(d, "amount")
        c_guest = _col(d, "guest_full") or _col(d, "guest")
        c_comptype = _col(d, "comp type")
        c_trx = _col(d, "trx_desc")
        # take filter = sheet column H (8th col) == 'Y' (project team 2026-06-13).
        # The old 是否包括在25年 flag over-took ~2× (69.4M vs 取數 34,354k).
        if len(d.columns) >= 8:
            c_h = d.columns[7]
            d = d[d[c_h].astype(str).str.strip().str.upper().eq("Y")]
            print(f"combined: column-H take filter {str(c_h)[:30]!r}=Y → {len(d):,} rows "
                  f"(target Σ≈34,354k)")
        elif c_flag:
            d = d[d[c_flag].astype(str).str.strip().str.upper().eq("Y")]

        # project-team ground truth: GUEST_FULL_NAME → 表1 Description → that row's vertical_id.
        guest_v = build_guest_v_map(d, c_guest, je_df, je_desc_col, je_vid_col) if c_guest else {}

        n_comb = n_je = n_rule = n_kw = 0
        for _, r in d.iterrows():
            amt = pd.to_numeric(r.get(c_amt), errors="coerce")
            if pd.isna(amt) or amt == 0:
                continue
            h = _h_from_type(r.get(c_type), r.get(c_trx))
            guest = str(r.get(c_guest, "")).strip()
            v = guest_v.get(guest)                       # 1) 表1 Description 1-to-1 (project-team truth)
            src = "combined:je"
            if v:
                n_je += 1
            else:
                v = _v_from_sjm_guest(guest)             # 2) SJM guest-name rules
                if v:
                    src = "combined:rule"; n_rule += 1
                else:
                    v = _v_from_event(guest, r.get(c_comptype, ""), r.get(c_type, ""))  # 3) generic keyword
                    src = "combined:kw"
                    if v:
                        n_kw += 1
            if not v:
                key = guest or str(r.get(c_comptype, ""))
                unmapped[key] = unmapped.get(key, 0) + float(amt)
                v = "V_OTHER"; src = "combined:other"
            rows.append([v, h, guest[:60], round(float(amt), 2), src])
            n_comb += 1
        _comb_sum = sum(r[3] for r in rows if str(r[4]).startswith("combined"))
        print(f"combined (col-H=Y): {n_comb} rows  Σ={_comb_sum:,.0f} (target ≈34,354,000)  "
              f"(V via 表1 match: {n_je}, SJM rule: {n_rule}, keyword: {n_kw}, V_OTHER: {n_comb - n_je - n_rule - n_kw})")

    return rows, unmapped


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--file", default="data/sjm/raw/Admin Comp summary v2.xlsx")
    p.add_argument("--parquet", default="data/sjm/output/company_2_kpi_report.parquet",
                   help="表1 (JE tagged rows) for GUEST_FULL_NAME → Description → V match")
    args = p.parse_args()
    je_df = je_desc = None
    if Path(args.parquet).exists():
        je_df = pd.read_parquet(args.parquet)
        je_desc = next((c for c in je_df.columns if str(c).strip().lower() == "description"), None) \
            or next((c for c in je_df.columns if "description" in str(c).lower() or "摘要" in str(c)), None)
        print(f"表1 loaded: {len(je_df):,} rows, description col = {je_desc!r}")
    rows, unmapped = build_rows(args.file, je_df=je_df, je_desc_col=je_desc, je_vid_col="vertical_id")
    out = Path("results"); out.mkdir(exist_ok=True)
    with (out / "sjm_admincomp_rows.tsv").open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter="\t")
        w.writerow(["vertical_id", "horizontal_id", "subproject", "amount", "source"])
        w.writerows(rows)

    tot = sum(r[3] for r in rows)
    byH = {}
    for r in rows:
        byH[r[1]] = byH.get(r[1], 0) + r[3]
    print(f"\nTotal admin-comp rows: {len(rows):,}  amount: {tot:,.0f}")
    print("By H:")
    for h, a in sorted(byH.items(), key=lambda kv: -abs(kv[1])):
        print(f"   {h:<16} {a:>16,.0f}")
    if unmapped:
        print(f"\n⚠ {len(unmapped)} combined events with UNMAPPED V (→V_OTHER for now) — paste these so I refine EVENT2V:")
        for ev, a in sorted(unmapped.items(), key=lambda kv: -abs(kv[1]))[:40]:
            print(f"   {str(ev)[:50]:<50} {a:>14,.0f}")
    print(f"\n→ results/sjm_admincomp_rows.tsv")


if __name__ == "__main__":
    main()
