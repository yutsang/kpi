r"""recon_comp_detail —— comp 按 HQ 報告維度拆：性質(房/會場/餐飲/票/其他) × NG × 6家 × 年 × {調整前/調整後}。

對 HQ「贈送內部資源支出和折扣優惠」報告（user 2026-06-22 貼）做格對格 check。
兩個 section（同 HQ 表一樣）：
  1. 按性質分類（horizontal_id → 客房/會場/食品飲料/贈送門票/其他）× entity
  2. 按投資範疇分類（ng_code → 吸引外國客源…）× entity
每個 section：3年(23/24/25) × 2 basis(調整前 pre / 調整後 post)。
年份用 spent-prefix（24=24+24_23SY；25=25+25_23SY+25_24SY；23=exact）。

Run（Windows）:  python scripts\recon_comp_detail.py
出 results\recon_comp_detail.txt（tab 分隔，可貼入 Excel 對）← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
OUT = ROOT / "results" / "recon_comp_detail.txt"

ENTS = ["Galaxy", "Wynn", "VML", "Melco", "MGM", "SJM"]
EKEY = {"Galaxy": "galaxy", "Wynn": "wynn", "VML": "vml", "Melco": "melco", "MGM": "mgm", "SJM": "sjm"}

NATURE = [  # (horizontal_id, HQ 性質 label)
    ("H_HOTEL_ROOM", "客房支出"), ("H_VENUE", "會場支出"), ("H_FNB", "食品飲料支出"),
    ("H_COMP_TICKET", "贈送門票支出"), ("H_COMP_OTHER", "其他"),
]
COMP_HID = {h for h, _ in NATURE}
NG_ORDER = [  # ng_code → HQ 投資範疇 label
    ("NG1", "吸引外國客源"), ("NG2", "會議展覽"), ("NG3", "娛樂表演"), ("NG4", "體育盛事"),
    ("NG5", "文化藝術"), ("NG6", "健康養生"), ("NG7", "主題遊樂"), ("NG8", "美食之都"),
    ("NG9", "社區旅遊"), ("NG10", "海上旅遊"), ("NG11", "其他"), ("NG0", "博彩"),
]


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    L = ["# recon_comp_detail —— comp 按性質/NG × 6家 × 年 × {調整前/調整後}（萬，prefix年）"]
    if not CSV.exists():
        L.append("!! csv 揾唔到"); _w(L); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    df["pre"] = pd.to_numeric(df.get("調整前_萬", 0), errors="coerce").fillna(0.0)
    df["post"] = pd.to_numeric(df.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    df["e"] = df["entity"].astype(str)
    df["hid"] = _s(df.get("horizontal_id", pd.Series("", index=df.index)))
    df["ng"] = _s(df.get("ng_code", pd.Series("", index=df.index))).str.upper()
    yb = df["year_bucket"].astype(str)
    df["y2"] = yb.str[:2]
    df["yb"] = yb
    comp = df[df["hid"].isin(COMP_HID)].copy()

    def amt(sub, basis):
        return sub[basis].sum()

    def yr_mask(sub, yr):
        return sub[sub.yb.eq("23")] if yr == "23" else sub[sub.y2.eq(yr)]

    def block(title, rows, rowkey_fn):
        # rows = list of (key, label); rowkey_fn(comp_subset_for_entity, key) → mask within
        for basis, blabel in [("pre", "調整前"), ("post", "調整後")]:
            for yr in ("23", "24", "25"):
                L.append(f"\n### {title} ｜ {yr}年 {blabel}")
                hdr = "性質/NG\t" + "\t".join(ENTS) + "\t六家合計"
                L.append(hdr)
                col_tot = {e: 0.0 for e in ENTS}
                for key, lab in rows:
                    cells = []
                    rowsum = 0.0
                    for E in ENTS:
                        ce = comp[comp.e.eq(EKEY[E])]
                        ce = yr_mask(ce, yr)
                        ce = rowkey_fn(ce, key)
                        v = ce[basis].sum()
                        cells.append(f"{v:,.0f}")
                        col_tot[E] += v
                        rowsum += v
                    L.append(f"{lab}\t" + "\t".join(cells) + f"\t{rowsum:,.0f}")
                L.append("合計\t" + "\t".join(f"{col_tot[E]:,.0f}" for E in ENTS) + f"\t{sum(col_tot.values()):,.0f}")

    # 1. 按性質
    L.append("\n" + "=" * 70 + "\n## 1. 按贈送內部資源性質分類")
    block("性質", NATURE, lambda ce, hid: ce[ce.hid.eq(hid)])
    # 2. 按 NG
    L.append("\n" + "=" * 70 + "\n## 2. 按投資範疇分類")
    block("NG", NG_ORDER, lambda ce, ng: ce[ce.ng.eq(ng)])

    _w(L)


def _w(L):
    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {OUT.relative_to(ROOT)}  ← paste 返嚟（tab 分隔可入 Excel）")


if __name__ == "__main__":
    main()
