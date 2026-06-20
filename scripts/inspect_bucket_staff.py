r"""check H_LABOR(opex) by ALL year_bucket per entity — 睇 staff 係咪散咗喺 25_24SY/25_23SY（exact「25」漏咗）。
mgm 25 golden 21,467 vs exact 16,610 → 睇 25 prefix（25+25_23SY+25_24SY）夾唔夾。

Run（Windows）:  python scripts\inspect_bucket_staff.py mgm galaxy
出 results\inspect_bucket_staff.txt
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
ENTS = sys.argv[1:] or ["mgm"]
T2 = {"galaxy": {"25": 23950}, "mgm": {"25": 21467}, "wynn": {"25": 10958},
      "vml": {"25": 1253}, "melco": {"25": 27806}, "sjm": {"25": 142}}
L = ["# inspect_bucket_staff — H_LABOR(opex,調整前 for 25) by year_bucket"]


def main():
    if not CSV.exists():
        print("!! csv 揾唔到"); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    pre = pd.to_numeric(df.get("調整前_萬", 0), errors="coerce").fillna(0.0)
    post = pd.to_numeric(df.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    yb = df["year_bucket"].astype(str)
    df["pre"] = pre; df["post"] = post; df["yb"] = yb
    hid = df.get("horizontal_id", pd.Series("", index=df.index)).astype(str).str.strip()
    co = df.get("final_capex_opex", pd.Series("", index=df.index)).astype(str).str.strip()
    df["ent"] = df["entity"].astype(str)
    for ent in ENTS:
        L.append(f"\n## {ent}  (opex H_LABOR)")
        m = (df.ent == ent) & hid.eq("H_LABOR") & ~co.eq("Capex")
        sub = df[m]
        L.append(f"   {'bucket':<10}{'調整前':>10}{'調整後':>10}")
        for b in ["23", "24", "24_23SY", "25", "25_23SY", "25_24SY"]:
            s = sub[sub.yb == b]
            if len(s):
                L.append(f"   {b:<10}{s['pre'].sum():>10,.0f}{s['post'].sum():>10,.0f}")
        pref25 = sub[sub.yb.str.startswith("25")]
        g = T2.get(ent, {}).get("25", 0)
        L.append(f"   → 25 exact: 前={sub[sub.yb=='25']['pre'].sum():,.0f} | 25 prefix(25+SY): 前={pref25['pre'].sum():,.0f} 後={pref25['post'].sum():,.0f} | golden25={g:,.0f}")
    out = ROOT / "results" / "inspect_bucket_staff.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"wrote {out.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
