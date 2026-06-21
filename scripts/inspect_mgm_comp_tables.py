r"""MGM "Standard: Table Games Revenue / Slot Revenue / Other Operating" comp rows — 拆細睇
係 gaming credit（自由籌碼/free play）定係 physical comp（房/餐/票）。

拆法：account_code / account_desc / description（JE行） / vertical_id

Run（Windows）:  python scripts\inspect_mgm_comp_tables.py
出 results\inspect_mgm_comp_tables.txt  ← paste 返嚟
"""
from __future__ import annotations
import re
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV  = ROOT / "tableau_combined_25.csv"
OUT  = ROOT / "results" / "inspect_mgm_comp_tables.txt"

COMP_H = {"H_HOTEL_ROOM", "H_VENUE", "H_FNB", "H_COMP_TICKET", "H_COMP_OTHER"}

# comp_type patterns to drill into
TARGET_CT_RE = re.compile(r"Table Games|Slot Revenue|Other Operating", re.I)

PHYS_KW = re.compile(r"Room|Hotel|F&B|Food|Beverage|Ticket|Show|Spa|房|餐|票", re.I)
GAME_KW = re.compile(r"Free Play|Gaming Credit|Credit|Chip|Promo|Casino|Marker|Points", re.I)

L: list[str] = ["# inspect_mgm_comp_tables — MGM Standard:Revenue comp rows 拆細"]


def _s(x): return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    if not CSV.exists():
        L.append("!! tableau_combined_25.csv 揾唔到"); _w(); return

    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    e = df[df["entity"].astype(str).eq("mgm")].copy()

    post = pd.to_numeric(e.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    pre  = pd.to_numeric(e.get("調整前_萬",  0), errors="coerce").fillna(0.0)
    yb   = e["year_bucket"].astype(str)
    e["m"]  = post.where(yb.str[:2].eq("24"), pre)   # comp basis
    e["y2"] = yb.str[:2]

    for c in ("horizontal_id", "comp_type", "account_code", "account_desc",
              "description", "vertical_id", "vertical_label", "ng_code"):
        e[c] = _s(e[c]) if c in e.columns else ""
    e["code"] = e["account_code"].str.replace(r"\.0$", "", regex=True)

    hid = e["horizontal_id"]
    ct  = e["comp_type"]
    ad  = e["account_desc"]

    # ── DEBUG: MGM H_COMP_OTHER distinct values ───────────────────────────────
    L.append("\n── DEBUG: MGM H_COMP_OTHER unique (comp_type, account_desc) top 20 by Σ ──")
    dbg = e[hid.isin(COMP_H)].groupby(["comp_type", "account_desc"])["m"].sum().reset_index()
    dbg = dbg.sort_values("m", key=lambda s: s.abs(), ascending=False)
    for _, r in dbg.head(20).iterrows():
        L.append(f"  ct={r['comp_type'][:20]!r:<22} ad={r['account_desc'][:30]!r:<32}  {r['m']:>8,.0f}萬")

    # 搜 comp_type OR account_desc（原欄位唔確定，兩個都搵）
    _in_ct = ct.str.contains(TARGET_CT_RE, na=False)
    _in_ad = ad.str.contains(TARGET_CT_RE, na=False)
    target = e[hid.isin(COMP_H) & (_in_ct | _in_ad)].copy()

    if target.empty:
        L.append("\n  !! 搵唔到 Target rows（comp_type 同 account_desc 都冇匹配）"); _w(); return

    L.append(f"\n總目標行數: {len(target):,}  Σ={target['m'].sum():,.0f}萬")

    for yr in ["23", "24", "25"]:
        sub = target[target["y2"] == yr]
        if sub.empty:
            continue
        L.append(f"\n{'='*68}\n## MGM  year_bucket prefix = {yr}  ({len(sub):,} 行  Σ={sub['m'].sum():,.0f}萬)")

        # ── 1. by comp_type + horizontal_id ──────────────────────────────────
        L.append("\n  ── 1. comp_type × horizontal_id ──")
        g1 = sub.groupby(["comp_type", "horizontal_id"])["m"].sum().reset_index().sort_values("m", ascending=False)
        for _, r in g1.iterrows():
            L.append(f"    ct={r['comp_type'][:30]:<30}  H={r['horizontal_id']:<15}  {r['m']:>8,.0f}萬")

        # ── 2. by account_code + account_desc ────────────────────────────────
        L.append("\n  ── 2. account_code × account_desc (top 12) ──")
        g2 = sub.groupby(["code", "account_desc"])["m"].sum().reset_index().sort_values("m", key=lambda s: s.abs(), ascending=False)
        for _, r in g2.head(12).iterrows():
            L.append(f"    code={r['code']:<10}  {r['account_desc'][:40]:<40}  {r['m']:>8,.0f}萬")

        # ── 3. by vertical_id ────────────────────────────────────────────────
        L.append("\n  ── 3. vertical_id breakdown ──")
        g3 = sub.groupby(["vertical_id", "vertical_label"])["m"].sum().reset_index().sort_values("m", ascending=False)
        for _, r in g3.iterrows():
            L.append(f"    {r['vertical_id']:<12}  {r['vertical_label'][:20]:<20}  {r['m']:>8,.0f}萬")

        # ── 4. description keyword count ─────────────────────────────────────
        desc_col = "description" if sub["description"].str.strip().ne("").any() else None
        if desc_col:
            L.append("\n  ── 4. description top keywords ──")
            descs = sub["description"].str.strip()
            phys_n = int(descs.str.contains(PHYS_KW, na=False).sum())
            game_n = int(descs.str.contains(GAME_KW, na=False).sum())
            blank_n = int(descs.eq("").sum())
            L.append(f"    Physical comp keywords (Room/Food/Ticket…): {phys_n} 行")
            L.append(f"    Gaming credit keywords (Free Play/Credit…) : {game_n} 行")
            L.append(f"    Blank description                          : {blank_n} 行")
            # top 8 unique description snippets
            top_d = sub.groupby("description")["m"].sum().sort_values(ascending=False).head(8)
            for d, amt in top_d.items():
                L.append(f"      {str(d)[:55]:<55}  {amt:>8,.0f}萬")
        else:
            L.append("\n  ── 4. description: all blank ──")

        # ── 5. ng_code breakdown ─────────────────────────────────────────────
        L.append("\n  ── 5. ng_code breakdown ──")
        g5 = sub.groupby("ng_code")["m"].sum().reset_index().sort_values("m", ascending=False)
        for _, r in g5.head(8).iterrows():
            L.append(f"    {r['ng_code']:<8}  {r['m']:>8,.0f}萬")

    _w()


def _w():
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
