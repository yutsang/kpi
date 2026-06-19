r"""Gap hunter —— 揾 recon 仲對唔上嘅 staff/comp 金額匿喺邊（user step-4：睇 mark 嘅規律）。

針對未 tie 嘅 entity（galaxy/wynn/sjm 等），逐個 metric（人工/客房/會場/食飲/贈票）：
  揾「desc/account/comp_type 含對應關鍵詞」但「而家 H 唔係對應 H」嘅行，group by 現H+comp_type+is_labor，
  列金額 + sample account_desc → 睇佢哋而家落咗去邊、有冇 mark 我哋漏咗。

Run（Windows）:  python scripts\inspect_gap_hunter.py            # 全 6 家
                 python scripts\inspect_gap_hunter.py galaxy wynn # 指定
出 results\inspect_gap_hunter.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
ENTS = sys.argv[1:] or ["galaxy", "sjm", "wynn", "vml", "melco", "mgm"]

# metric → (target H, 關鍵詞)
METRICS = [
    ("人工", "H_LABOR", ["人工", "staff", "Staff", "STAFF", "salary", "Salary", "payroll", "Payroll",
                         "薪", "工資", "labour", "labor", "Labor", "員工", "人力", "headcount", "wage"]),
    ("客房", "H_HOTEL_ROOM", ["客房", "房", "room", "Room", "ROOM", "hotel", "Hotel", "酒店", "ADR", "房晚", "Night"]),
    ("會場", "H_VENUE", ["會場", "場地", "venue", "Venue", "場館", "Function", "Event Space", "Ballroom"]),
    ("食飲", "H_FNB", ["餐", "食", "飲", "F&B", "FnB", "FNB", "Food", "Beverage", "Dining", "宴", "Banquet"]),
    ("贈票", "H_COMP_TICKET", ["票", "ticket", "Ticket", "贈票", "門票", "入場"]),
]
L: list[str] = ["# inspect_gap_hunter —— 揾 staff/comp 金額匿喺邊（萬=調整後_萬）"]


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    if not CSV.exists():
        L.append("!! tableau_combined_25.csv 揾唔到"); _w(); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    df["amt"] = pd.to_numeric(df.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    df["y2"] = df["year_bucket"].astype(str).str[:2]
    for c in ("entity", "horizontal_id", "horizontal_label", "comp_type", "is_labor",
              "is_internal", "account_desc", "description", "科目明細"):
        if c not in df.columns:
            df[c] = ""
        df[c] = _s(df[c])
    df["_blob"] = (df["account_desc"] + " | " + df["description"] + " | " + df["科目明細"] + " | " + df["comp_type"])

    for ent in ENTS:
        e = df[df.entity == ent]
        if not len(e):
            continue
        L.append(f"\n{'='*78}\n# {ent}")
        for y in ("23", "24"):
            sub = e[e.y2 == y]
            if not len(sub):
                continue
            L.append(f"\n  ── {ent} {y} ──")
            for lab, hid, kws in METRICS:
                cur = sub.loc[sub.horizontal_id == hid, "amt"].sum()
                # 含關鍵詞但唔喺 target H 嘅行（= 可能漏咗）
                hit = sub[sub["_blob"].str.contains("|".join(kws), case=False, na=False, regex=True)]
                miss = hit[hit.horizontal_id != hid]
                if abs(miss["amt"].sum()) < 30:
                    continue
                L.append(f"    [{lab}] 現{hid}={cur:,.0f}萬；含關鍵詞但唔喺{hid} 共 {miss['amt'].sum():,.0f}萬 →")
                g = miss.groupby(["horizontal_label", "comp_type", "is_labor"])["amt"].sum().reset_index().sort_values("amt", key=lambda s: s.abs(), ascending=False)
                for _, r in g.head(5).iterrows():
                    samp = miss[(miss.horizontal_label == r['horizontal_label']) & (miss.comp_type == r['comp_type'])]["account_desc"].head(2).tolist()
                    L.append(f"        {r['amt']:>8,.0f}萬 | 現H={r['horizontal_label'][:10]:<10} comp_type={r['comp_type'][:14]:<14} is_labor={r['is_labor'][:8]:<8} | {' ; '.join(str(x)[:22] for x in samp)}")
    _w()


def _w():
    out = ROOT / "results" / "inspect_gap_hunter.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
