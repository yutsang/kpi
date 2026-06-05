"""audit_vproj.py — AUTO-FLAG likely vertical (V) errors by matching the PROJECT NAME against
our_V. High-precision keyword rules only (so the flagged list is short + worth eyeballing).
NG is FIXED (from 投資領域) and never touched — this only questions the V within that NG.

Each rule = a name pattern + the V it implies; we flag a project when its name matches but our_V
is a different (confusable) category. Output → paste the SHORT flagged list → Claude confirms →
per-entity row_vertical_overrides.

  python scripts/audit_vproj.py --entity melco --year 25
  python scripts/audit_vproj.py --entity galaxy sjm wynn melco --year 25
  python scripts/audit_vproj.py --all --year 25        # every entity
  python scripts/audit_vproj.py --entity melco --year 25 --min 100000   # only Σ|amt|>=100k
"""
from __future__ import annotations
import argparse, re, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd
import pyarrow.parquet as pq
import yaml

ROOT = Path(__file__).resolve().parent.parent
ENT = {"galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
       "vml": "company_4", "melco": "company_5", "mgm": "company_6"}

# rule = (id, name_regex, implied_V_keyword, flag_when_our_V_contains[list or None])
#   flag a project if name matches name_regex, our_V does NOT contain implied_V_keyword,
#   and (flag_when is None OR our_V contains one of flag_when). Keep rules HIGH-PRECISION.
RULES = [
    ("venue_as_equip",  r"gaming venue|娛樂場(?!.*設施)|mass floor|casino floor|賭場.*優化",
        "娛樂場", ["設施設備"]),
    ("equip_as_venue",  r"gaming equipment|博彩設施|slot|角子機|EGM|smart table|surveillance|監察|數字化|digital",
        "設施設備", ["娛樂場優化"]),
    ("music_as_sport",  r"guzheng|古箏|piano|鋼琴|orchestra|樂團|ballet|芭蕾|dance|舞蹈|art|藝術|music|音樂|film|電影|opera|歌劇",
        "文藝", ["體育"]),
    ("museum_as_sport", r"museum|博物館|exhibition|展覽",
        "文藝", ["體育"]),
    ("restaurant_else", r"restaurant|steakhouse|brasserie|餐廳|eatery|café|cafe\b",
        "餐廳", ["體育", "娛樂表演", "邀請外國", "其他", "社區", "會議"]),
    ("concert_else",    r"concert|演唱會|live in macau|world tour|fan ?meeting|fancon|fan con|巡迴演唱",
        "娛樂表演", ["體育", "餐廳", "主題遊樂", "其他", "社區"]),
    ("mice_as_else",    r"conference|congress|seminar|會議|論壇|峰會|summit|\bforum\b|博覽|會展|tradeshow",
        "會議", ["美食", "餐廳", "參與海外", "邀請外國", "體育"]),
    ("sport_as_else",   r"grand prix|格蘭披治|\bgolf\b|tennis|marathon|龍舟|dragon boat|錦標賽|championship|tournament|\bUFC\b|\bWTT\b|\bITTF\b|全運會",
        "體育", ["餐廳", "美食", "主題遊樂", "其他", "邀請外國", "文藝"]),
    ("food_as_roadshow", r"美食|gastronomy|culinary|wine dinner|food fair|美酒",
        "美食", ["參與海外", "邀請外國", "社區"]),
    ("culture_as_other", r"culture|文化|art exhibition|藝術|文藝",
        "文藝", ["其他"]),
]


def numify(s):
    return pd.to_numeric(s.astype(str).str.replace(",", "", regex=False).str.replace(r"^\s*-\s*$", "0", regex=True),
                         errors="coerce").fillna(0)


def find(df, *subs, exact=None):
    if exact and exact in df.columns:
        return exact
    for c in df.columns:
        if any(s in str(c) for s in subs):
            return c
    return None


def mode(s):
    s = s.astype(str).str.strip()
    s = s[s.ne("") & s.ne("nan")]
    return s.mode().iloc[0] if len(s.mode()) else ""


def _cn_kw(s) -> str:
    s = str(s)
    if "非博彩" in s: return ""
    for kws, ng in [(["博彩", "gaming"], "NG0"), (["海上"], "NG10"), (["外國", "客源", "國際客"], "NG1"),
                    (["會議", "會展", "mice"], "NG2"), (["娛樂", "演唱", "表演"], "NG3"), (["體育", "賽事"], "NG4"),
                    (["文化", "藝術", "文藝"], "NG5"), (["健康", "養生"], "NG6"), (["主題", "遊樂"], "NG7"),
                    (["美食", "餐飲"], "NG8"), (["社區"], "NG9"), (["其他"], "NG11")]:
        if any(k in s or k in s.lower() for k in kws): return ng
    return ""


def ng_cols_of(df, cf, cols):
    names = [cols.get("ng11_category", "")] + [
        (ys.get("columns_override") or {}).get("ng11_category") for ys in (cf.get("yearly_sources") or [])]
    out = []
    for nm in names:
        fc = find(df, exact=nm)
        if fc and fc not in out: out.append(fc)
    for c in df.columns:
        if c not in out and any(k in str(c) for k in ("項目性質", "項目類型", "項目分類", "範疇", "投資領域", "NG11 Category", "Section")):
            out.append(c)
    return out


def derive_ng(df, ngcols):
    try:
        sys.path.insert(0, str(ROOT / "src"))
        from kpi.lib.conf import load_categories
        from kpi.pipelines.step2_tag_projects._logic import normalize_ng_code
        cats = load_categories()
        def res(x):
            for c in (x, x.upper().replace(" ", "")):
                r = normalize_ng_code(c, cats) or ""
                if r[:2] == "NG" and r[2:].isdigit(): return r
            return _cn_kw(x)
    except Exception:
        res = _cn_kw
    out = pd.Series("", index=df.index, dtype="object")
    for fc in ngcols:
        m = {x: res(x) for x in set(df[fc].astype(str).unique())}
        r = df[fc].astype(str).map(m).fillna("")
        r = r.where(r.str.fullmatch(r"NG\d+").fillna(False), "")
        out = out.mask(out.eq(""), r)
    return out.replace("", "(未分類)")


def audit(ent, year, minamt):
    com = ENT[ent]
    src = ROOT / "data" / ent / "interim" / f"{com}_tagged_rows.parquet"
    if not src.exists():
        print(f"\n#### {ent}: X missing tagged_rows — run kedro"); return
    cf = yaml.safe_load((ROOT / "conf" / com / "parameters.yml").open(encoding="utf-8"))
    cols = cf.get("columns", {}) or {}
    df = pq.read_table(src).replace_schema_metadata(None).to_pandas()
    if "report_period" in df.columns:
        df = df[df["report_period"].astype(str).str.startswith(year)].copy()
    if df.empty:
        print(f"\n#### {ent} {year}: no rows"); return
    amt = next((c for c in [cols.get("amount"), "MOP Amt", "調整後金額", "Reported Amount(MOP)"]
                if c and c in df.columns and numify(df[c]).abs().sum() > 0), None)
    df["_amt"] = numify(df[amt]) if amt else 0.0
    proj = find(df, exact=cols.get("project")) or find(df, "Project Name", "SubProject_Name", "Project name", "Project", "項目")
    ourv = "vertical_label" if "vertical_label" in df.columns else "vertical_id"
    df["_ng"] = derive_ng(df, ng_cols_of(df, cf, cols))
    if not proj:
        print(f"\n#### {ent} {year}: no project column"); return
    df["_p"] = df[proj].astype(str).str.strip()
    g = df.groupby("_p").agg(_v=(ourv, mode), _ng=("_ng", mode), _amt=("_amt", "sum")).reset_index()

    flags = []
    for _, r in g.iterrows():
        nm = r["_p"]; v = str(r["_v"])
        for rid, name_re, implied, flag_when in RULES:
            if not re.search(name_re, nm, re.I):
                continue
            if implied in v:
                continue
            if flag_when and not any(w in v for w in flag_when):
                continue
            flags.append({"rule": rid, "project": nm, "ng": r["_ng"], "our_V": v,
                          "implies": implied, "amt": r["_amt"]})
            break  # one flag per project (first matching rule)
    flags = [f for f in flags if abs(f["amt"]) >= minamt]
    flags.sort(key=lambda x: -abs(x["amt"]))
    print(f"\n#### {ent} {year}: {len(g)} projects, {len(flags)} flagged (|amt|>={minamt:,.0f}) "
          f"[proj={proj!r} V={ourv!r} amt={amt!r}]")
    if not flags:
        print("   (no high-confidence name↔V mismatches)"); return
    print("   rule | project | NG | our_V | name-implies | Σamt")
    for f in flags:
        print(f"   [{f['rule']:15s}] {str(f['project'])[:40]:42s} {f['ng']:<6} "
              f"{str(f['our_V'])[:13]:15s} ~{f['implies']:<6} {f['amt']:>14,.0f}")


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--entity", nargs="+", choices=sorted(ENT))
    p.add_argument("--all", action="store_true")
    p.add_argument("--year", required=True)
    p.add_argument("--min", type=float, default=0.0, help="only flag projects with |Σamt| >= this")
    a = p.parse_args()
    ents = sorted(ENT) if a.all else (a.entity or [])
    if not ents:
        print("specify --all or --entity ..."); return
    for e in ents:
        audit(e, a.year, a.min)


if __name__ == "__main__":
    main()
