r"""comp H 行嘅來源拆解（galaxy/wynn/mgm）：
  galaxy：pt_class_H (項目組H 含 Comp) vs 非 pt_class_H (desc/comp_type)
  wynn  ：有 comp_type 嘅行 vs 冇 comp_type 但有 comp-keyword 嘅行（找漏標）
  mgm   ：有 comp_type / 有 PT source comp 行拆解（找 over 來源）

Run（Windows）:
  python scripts\inspect_comp_source.py galaxy wynn mgm sjm vml
出 results\inspect_comp_source_<ent>.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd, re

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
ENTS = sys.argv[1:] or ["galaxy", "wynn", "mgm"]
COMP_H = {"H_HOTEL_ROOM", "H_VENUE", "H_FNB", "H_COMP_TICKET", "H_COMP_OTHER"}
COMP_LABEL = {"H_HOTEL_ROOM": "Comp房間", "H_VENUE": "Comp活動場地",
              "H_FNB": "Comp餐飲", "H_COMP_TICKET": "Comp贈票", "H_COMP_OTHER": "Comp其他"}
COMP_KW = re.compile(
    r"comp|complimentary|贈|客房|Ritz|Marriott|Okura|贈房|贈票|Ticket|Venue|Hall|Ballroom",
    re.IGNORECASE)

HQ_COMP = {
    "galaxy": {"23": 66407, "24": 55309, "25": 91878},
    "wynn":   {"23": 25772, "24": 24459, "25": 21489},
    "vml":    {"23": 6338,  "24": 6327,  "25": 12802},
    "melco":  {"23": 3833,  "24": 6223,  "25": 8424},
    "mgm":    {"23": 10189, "24": 10549, "25": 7914},
    "sjm":    {"23": 2970,  "24": 6553,  "25": 3575},
}


def _s(x): return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def _m_exact(e, yr):
    post = pd.to_numeric(e.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    pre  = pd.to_numeric(e.get("調整前_萬",  0), errors="coerce").fillna(0.0)
    yb   = e["year_bucket"].astype(str)
    return post.where(yb.isin(["23", "24"]), pre)


def run_galaxy(e, L):
    L.append("\n## galaxy — pt_class_H vs 非 pt_class_H comp 分解")
    _ph = _s(e.get("項目組H", pd.Series("", index=e.index)) if "項目組H" in e.columns
             else pd.Series("", index=e.index))
    _from_pt = _ph.str.split("|").str[0].str.strip().eq("Comp")

    for yr in ("23", "24", "25"):
        if yr == "23":
            sub = e[e["year_bucket"].astype(str) == "23"].copy()
        else:
            sub = e[e["year_bucket"].astype(str).str[:2] == yr].copy()
        sub["m"] = _m_exact(sub, yr)
        sub_comp = sub[sub["horizontal_id"].isin(COMP_H)]
        total = sub_comp["m"].sum()
        hq = HQ_COMP.get("galaxy", {}).get(yr, 0)
        L.append(f"\n  === galaxy bucket {yr}  comp合計={total:,.0f}萬  HQ={hq:,.0f}萬  Δ={total-hq:,.0f}萬 ===")

        # A: pt_class_H = Comp
        pt_mask = _from_pt.reindex(sub_comp.index, fill_value=False)
        pt = sub_comp[pt_mask]
        npt = sub_comp[~pt_mask]
        L.append(f"    A) 來自 pt_class_H(項目組H=Comp)  rows={len(pt):,}  Σ={pt['m'].sum():,.0f}萬")
        L.append(f"    B) 非 pt_class_H (desc/comp_type)  rows={len(npt):,}  Σ={npt['m'].sum():,.0f}萬")
        if len(npt):
            g = npt.groupby(["horizontal_id", "account_desc"])["m"].sum().reset_index()
            g = g[g["m"].abs() >= 50].sort_values("m", key=lambda s: s.abs(), ascending=False)
            L.append("    B breakdown (top 15 by |amt|):")
            for _, r in g.head(15).iterrows():
                L.append(f"      H={COMP_LABEL.get(r['horizontal_id'], r['horizontal_id']):<14} {str(r['account_desc'])[:30]:<30} {r['m']:>8,.0f}萬")


def run_wynn(e, L):
    L.append("\n## wynn — comp_type 有/無 + 漏標 keyword scan")
    for yr in ("23", "24", "25"):
        if yr == "23":
            sub = e[e["year_bucket"].astype(str) == "23"].copy()
        else:
            sub = e[e["year_bucket"].astype(str).str[:2] == yr].copy()
        sub["m"] = _m_exact(sub, yr)
        hid   = _s(sub["horizontal_id"]) if "horizontal_id" in sub.columns else pd.Series("", index=sub.index)
        ct    = _s(sub["comp_type"])     if "comp_type"     in sub.columns else pd.Series("", index=sub.index)
        ad    = sub["account_desc"].astype(str) if "account_desc" in sub.columns else pd.Series("", index=sub.index)
        desc  = sub["description"].astype(str) if "description" in sub.columns else pd.Series("", index=sub.index)

        # 現 comp H
        comp_mask = hid.isin(COMP_H)
        sub_comp = sub[comp_mask]
        total = sub_comp["m"].sum()
        hq = HQ_COMP.get("wynn", {}).get(yr, 0)
        L.append(f"\n  === wynn bucket {yr}  comp合計={total:,.0f}萬  HQ={hq:,.0f}萬  Δ={total-hq:,.0f}萬 ===")

        # 有 comp_type 的 comp 行
        has_ct = comp_mask & ct.ne("")
        L.append(f"    a) 現 comp 有 comp_type  rows={int(has_ct.sum()):,}  Σ={sub.loc[has_ct,'m'].sum():,.0f}萬")
        L.append(f"    b) 現 comp 冇 comp_type  rows={int((comp_mask & ct.eq('')).sum()):,}  Σ={sub.loc[comp_mask & ct.eq(''),'m'].sum():,.0f}萬")

        # comp_type breakdown
        if int(has_ct.sum()):
            g = sub[has_ct].groupby("comp_type")["m"].sum().reset_index().sort_values("m", key=lambda s: s.abs(), ascending=False)
            L.append("    comp_type breakdown:")
            for _, r in g.iterrows():
                L.append(f"      ct={str(r['comp_type']):<20}  Σ={r['m']:>8,.0f}萬")

        # 非 comp H 但有 comp keyword → 可能漏標
        blob = (ad + " " + desc)
        maybe = sub[~comp_mask & blob.str.contains(COMP_KW, na=False)]
        if len(maybe):
            g2 = maybe.groupby(["horizontal_id", "account_desc"])["m"].sum().reset_index()
            g2 = g2[g2["m"].abs() >= 100].sort_values("m", key=lambda s: s.abs(), ascending=False)
            L.append(f"    c) 非comp但有comp keyword rows={len(maybe):,} Σ={maybe['m'].sum():,.0f}萬 (可能漏標 top15):")
            for _, r in g2.head(15).iterrows():
                L.append(f"      H={r['horizontal_id']:<14} {str(r['account_desc'])[:30]:<30} {r['m']:>8,.0f}萬")


def run_mgm(e, L):
    L.append("\n## mgm — comp H 分解（找 over 來源）")
    for yr in ("23", "24", "25"):
        if yr == "23":
            sub = e[e["year_bucket"].astype(str) == "23"].copy()
        else:
            sub = e[e["year_bucket"].astype(str).str[:2] == yr].copy()
        sub["m"] = _m_exact(sub, yr)
        hid = _s(sub["horizontal_id"]) if "horizontal_id" in sub.columns else pd.Series("", index=sub.index)
        ct  = _s(sub["comp_type"])     if "comp_type"     in sub.columns else pd.Series("", index=sub.index)
        ad  = sub["account_desc"].astype(str) if "account_desc" in sub.columns else pd.Series("", index=sub.index)
        pt  = _s(sub["項目組H"])           if "項目組H"        in sub.columns else pd.Series("", index=sub.index)
        src = sub["source"].astype(str)    if "source"        in sub.columns else pd.Series("", index=sub.index)

        sub_comp = sub[hid.isin(COMP_H)]
        total = sub_comp["m"].sum()
        hq = HQ_COMP.get("mgm", {}).get(yr, 0)
        L.append(f"\n  === mgm bucket {yr}  comp合計={total:,.0f}萬  HQ={hq:,.0f}萬  Δ={total-hq:,.0f}萬 ===")

        # breakdown by H sub-type
        for hid_v, lab in COMP_LABEL.items():
            hrows = sub_comp[hid.reindex(sub_comp.index, fill_value="").eq(hid_v)]
            if abs(hrows["m"].sum()) < 1: continue
            L.append(f"    {lab:<16} Σ={hrows['m'].sum():>8,.0f}萬  rows={len(hrows):,}")
            g = hrows.groupby("account_desc")["m"].sum().reset_index()
            g = g[g["m"].abs() >= 50].sort_values("m", key=lambda s: s.abs(), ascending=False)
            for _, r in g.head(8).iterrows():
                L.append(f"      {str(r['account_desc'])[:36]:<36} {r['m']:>8,.0f}萬")

        # comp_type breakdown
        ct_sub = _s(sub_comp["comp_type"]) if "comp_type" in sub_comp.columns else pd.Series("", index=sub_comp.index)
        g2 = sub_comp.groupby(ct_sub.values if len(ct_sub) else [""])["m"].sum()
        if len(g2):
            L.append("    comp_type breakdown:")
            for k, v in g2.sort_values(key=lambda s: s.abs(), ascending=False).items():
                if abs(v) >= 50:
                    L.append(f"      ct={str(k):<20}  Σ={v:>8,.0f}萬")


def run_generic(ent, e, L):
    L.append(f"\n## {ent} — comp H breakdown")
    for yr in ("23", "24", "25"):
        if yr == "23":
            sub = e[e["year_bucket"].astype(str) == "23"].copy()
        else:
            sub = e[e["year_bucket"].astype(str).str[:2] == yr].copy()
        sub["m"] = _m_exact(sub, yr)
        hid = _s(sub["horizontal_id"]) if "horizontal_id" in sub.columns else pd.Series("", index=sub.index)
        sub_comp = sub[hid.isin(COMP_H)]
        total = sub_comp["m"].sum()
        hq = HQ_COMP.get(ent, {}).get(yr, 0)
        L.append(f"\n  === {ent} bucket {yr}  comp合計={total:,.0f}萬  HQ={hq:,.0f}萬  Δ={total-hq:,.0f}萬 ===")
        g = sub_comp.groupby(["horizontal_id", "account_desc"])["m"].sum().reset_index()
        g = g[g["m"].abs() >= 100].sort_values("m", key=lambda s: s.abs(), ascending=False)
        for _, r in g.head(12).iterrows():
            L.append(f"    {COMP_LABEL.get(r['horizontal_id'], r['horizontal_id']):<14} {str(r['account_desc'])[:32]:<32} {r['m']:>8,.0f}萬")


def main():
    if not CSV.exists():
        print("!! tableau_combined_25.csv 揾唔到"); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    for ent in ENTS:
        e = df[df["entity"].astype(str) == ent].copy()
        L = [f"# inspect_comp_source — {ent}"]
        if ent == "galaxy":
            run_galaxy(e, L)
        elif ent == "wynn":
            run_wynn(e, L)
        elif ent == "mgm":
            run_mgm(e, L)
        else:
            run_generic(ent, e, L)
        p = ROOT / "results" / f"inspect_comp_source_{ent}.txt"
        p.parent.mkdir(exist_ok=True)
        p.write_text("\n".join(L), encoding="utf-8")
        print("\n".join(L))
        print(f"\nwrote {p.relative_to(ROOT)}\n")


if __name__ == "__main__":
    main()
