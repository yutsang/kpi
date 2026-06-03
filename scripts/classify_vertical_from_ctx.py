"""Classify each project's Vertical from its step2 context dump (results/ctx_{ent}.jsonl) —
the 'LLM here' step of the bypass. NG is taken from the project team's 項目性質 (the record's
`NG`, already normalized); V is chosen WITHIN that NG's candidate verticals by keyword rules
over the project name + name_hint + top accounts/descriptions. Output → data/_overrides/{ent}_vertical.tsv
('V_CODE<tab>project') for inject_manual_vertical.py.

V is constrained to candidates_for_ng(NG) MINUS V_OTHER for NG0-NG10 (so V_TO_NG(V) == 項目性質's
NG — picking V_OTHER would silently re-bucket the project to NG11). NG11 keeps V_OTHER/V_PUBLIC_FACILITY.

  python scripts/classify_vertical_from_ctx.py --entity melco            # review only
  python scripts/classify_vertical_from_ctx.py --entity melco --write    # write the TSV
  python scripts/classify_vertical_from_ctx.py --entity melco --show NG1 # eyeball one NG
"""
from __future__ import annotations
import argparse, json, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))

# Per-NG default V (must be a non-V_OTHER candidate of that NG).
DEFAULT = {
    "NG0": "V_GAMING_VENUE", "NG1": "V_INVITE_GUEST", "NG2": "V_MICE", "NG3": "V_CONCERT",
    "NG4": "V_SPORT_EVENT", "NG5": "V_ART_EXHIBITION", "NG6": "V_WELLNESS", "NG7": "V_THEME_PARK",
    "NG8": "V_RESTAURANT", "NG9": "V_COMMUNITY", "NG10": "V_MARITIME", "NG11": "V_OTHER",
}
# Per-NG keyword→V (first match wins; keywords matched lowercased as substrings, Chinese as-is).
RULES = {
    "NG0": [
        (["in016", "gaming equipment", "electronic gaming machine", "gaming machine", "egm",
          "slot", "角子機", "conversion kit", "chip counting", "smart table", "博彩機", "博彩設備",
          "gaming facility and equipment", "設施及設備", "perfect pay table", "kiosk"], "V_GAMING_EQUIP"),
        (["in015", "gaming venue", "gaming area", "gaming areas", "博彩區", "casino floor",
          "qi long", "皇璽會", "娛樂場場地", "娛樂場", "pit ", "section ", "layout change",
          "reconfiguration", "back wall"], "V_GAMING_VENUE"),
        (["boh", "back of house", "賬房", "cage", "lobby", "atrium", "it equipment", "network",
          "server", "infrastructure", "reception", "機房", "data centre", "data center"], "V_PROPERTY_UPGRADE"),
    ],
    "NG1": [
        (["roadshow", "路演"], "V_OVERSEAS_ROADSHOW"),
        (["overseas office", "海外辦事處", "representative office"], "V_OVERSEAS_OFFICE"),
        (["regional team", "區域代表", "國際代表", "代表團隊", "international representative"], "V_REGIONAL_TEAM"),
        (["travel agency", "旅行社", "tour operator", "travel trade", "bonvoy", "loyalty program",
          "world of hyatt", "agoda", "trip.com", "ctrip"], "V_INVITE_AGENCY"),
        (["group sales", "區域銷售", "sales rep", "sales representative", "集團銷售", "trade sales",
          "commercial fee"], "V_REGIONAL_SALES"),
        (["website", "seo", "search engine", "digital marketing", "online promotion", "海外網站",
          "media and online", "advertisment"], "V_OVERSEAS_WEB_SEO"),
        (["video", "視頻", "videography"], "V_PROMO_VIDEO"),
    ],
    "NG2": [
        (["event centre", "event center", "convention center", "convention centre", "會展中心",
          "演藝", "arena", "stadium", "館", "hall", "活動中心", "盛事及活動"], "V_VENUE_PERF_SPORT_MICE"),
        (["boh", "renovation", "facility upgrade", "infrastructure"], "V_PROPERTY_UPGRADE"),
    ],
    "NG3": [
        (["spectacle theatre", "spectacle theater", "international spectacle", "劇場", "theatre", "theater",
          "演藝中心", "concert hall", "arena", "館 ", "venue"], "V_VENUE_PERF_SPORT_MICE"),
        (["boh", "renovation", "facility upgrade", "infrastructure"], "V_PROPERTY_UPGRADE"),
    ],
    "NG4": [
        (["stadium", "arena", "體育館", "venue", "球場"], "V_VENUE_PERF_SPORT_MICE"),
        (["boh", "renovation", "facility upgrade", "infrastructure"], "V_PROPERTY_UPGRADE"),
    ],
    "NG5": [
        # V_MUSEUM intentionally NOT keyword-matched — '博物館' fires off minor sub-projects in the
        # aggregated name_hint (mis-tagged an art festival as museum). NG5 museums fall to
        # V_ART_EXHIBITION (same NG); add a targeted override if a genuine museum project appears.
        (["concert", "演唱", "live show", "music live", "performance", "show"], "V_CONCERT"),
        (["sport", "賽", "marathon", "tournament", "golf"], "V_SPORT_EVENT"),
        (["food", "美食", "dining", "culinary", "wine", "餐"], "V_FOOD_EVENT"),
        (["community", "社區", "donation", "charity", "公益", "youth"], "V_COMMUNITY"),
        (["mice", "conference", "convention", "會議", "會展"], "V_MICE"),
        (["venue", "arena", "館", "劇場"], "V_VENUE_PERF_SPORT_MICE"),
        (["boh", "renovation", "facility upgrade", "infrastructure"], "V_PROPERTY_UPGRADE"),
    ],
    "NG6": [
        (["boh", "renovation", "facility upgrade", "infrastructure", "fit out", "fit-out"], "V_PROPERTY_UPGRADE"),
    ],
    "NG7": [
        (["festival", "節", "carnival", "慶典", "parade", "花車", "countdown", "跨年", "new year"], "V_FESTIVAL"),
        (["boh", "renovation", "facility upgrade", "infrastructure"], "V_PROPERTY_UPGRADE"),
    ],
    "NG8": [
        (["fine dining", "casual dining", "restaurant", "餐廳", "gourmet pavilion", "美食廣場",
          "properties renovation", "new facilities", "experience enhancement", "renovation",
          "rebranding", "relaunch", "pavilion", "meeting event space and restaurant"], "V_RESTAURANT"),
        (["partnership event", "demonstration", "demostration", "black pearl", "dinner", "tasting",
          "festival", "美食活動", "美食節", "f&b general promo", "general promo", "pop up", "pop-up",
          "best bars", "best restaurants", "mixologist", "award", "promotion", "promo"], "V_FOOD_EVENT"),
    ],
    "NG9": [
        (["food", "美食", "dining", "餐"], "V_FOOD_EVENT"),
        (["art", "藝術", "exhibition", "展覽"], "V_ART_EXHIBITION"),
        (["mice", "conference", "會議", "會展"], "V_MICE"),
    ],
    "NG10": [],
    "NG11": [
        (["lrt", "輕軌", "public facility", "公共", "road", "bridge", "道路", "橋", "連廊",
          "transport", "infrastructure", "公用", "設施"], "V_PUBLIC_FACILITY"),
    ],
}

ENTITIES = {"galaxy", "sjm", "wynn", "vml", "melco", "mgm"}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--entity", required=True, choices=sorted(ENTITIES))
    ap.add_argument("--write", action="store_true", help="write data/_overrides/{ent}_vertical.tsv")
    ap.add_argument("--show", default=None, help="print every project of one NG (e.g. NG1)")
    args = ap.parse_args()

    from kpi.lib.conf import load_categories
    from kpi.pipelines.step2_tag_projects._logic import candidates_for_ng
    cats = load_categories()
    vlabel = {v["id"]: v["label"] for v in cats["verticals"]}
    cand = {ng: set(candidates_for_ng(ng, cats)) for ng in DEFAULT}

    src = ROOT / "results" / f"ctx_{args.entity}.jsonl"
    if not src.exists():
        print(f"X missing {src}"); return
    R = [json.loads(l) for l in src.open(encoding="utf-8") if l.strip()]

    def classify(r):
        ng = r["NG"] if r["NG"] in DEFAULT else "NG11"
        # Match on the NAMES only (project + name_hint). top_accounts/top_descs inject incidental
        # keyword noise (a show's desc mentions 'venue', a concert's ledger has 'agency'…).
        hay = " ".join([str(r.get("project", "")), str(r.get("name_hint", ""))]).lower()
        for kws, v in RULES.get(ng, []):
            if v in cand[ng] and any(k in hay for k in kws):
                return v
        d = DEFAULT[ng]
        return d if d in cand[ng] else "V_OTHER"

    for r in R:
        r["_V"] = classify(r)

    # ---- review ----
    from collections import defaultdict
    print(f"=== {args.entity}: {len(R)} projects ===")
    # per NG → V split (money)
    by = defaultdict(lambda: defaultdict(lambda: [0, 0.0]))
    for r in R:
        by[r["NG"]][r["_V"]][0] += 1
        by[r["NG"]][r["_V"]][1] += r["Σamt"]
    for ng in sorted(by, key=lambda x: (len(x), x)):
        tot = sum(v[1] for v in by[ng].values())
        print(f"\n{ng} ({DEFAULT.get(ng,'?')}) Σ={tot:,.0f}")
        for v in sorted(by[ng], key=lambda x: -abs(by[ng][x][1])):
            n, amt = by[ng][v]
            flag = "" if v in cand.get(ng, set()) else "  ⚠NOT-CANDIDATE"
            print(f"    {v:28s} n={n:4d}  Σ={amt:>15,.0f}  {vlabel.get(v,'')}{flag}")

    if args.show:
        ng = args.show
        rows = sorted([r for r in R if r["NG"] == ng], key=lambda r: -abs(r["Σamt"]))
        print(f"\n=== --show {ng}: {len(rows)} projects ===")
        for r in rows:
            nm = (r.get("name_hint") or r.get("project"))[:95]
            print(f"  {r['_V']:26s} {r['Σamt']:>13,.0f} | {nm}")

    if args.write:
        ov = ROOT / "data" / "_overrides" / f"{args.entity}_vertical.tsv"
        ov.parent.mkdir(parents=True, exist_ok=True)
        with ov.open("w", encoding="utf-8") as fh:
            fh.write(f"# {args.entity} vertical overrides — V from 項目性質-NG + name keywords (Claude-classify bypass)\n")
            for r in R:
                proj = str(r["project"]).replace("\t", " ").replace("\n", " ").strip()
                fh.write(f"{r['_V']}\t{proj}\n")
        print(f"\n→ wrote {len(R)} rows to {ov.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
