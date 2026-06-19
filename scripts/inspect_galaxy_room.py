r"""追蹤 galaxy（或指定家）comp room + ADR 調整行喺邊：year_bucket × NG × H × 調整前/後。

user：ADR 調整應該已 row-level 做咗 → comp room(含ADR) 應喺數據，可能落錯 bucket/H 或調整方向令淨額細。
呢個 dump：
  §A 項目組H 含 room 嘅行（galaxy Comp-room/Comp|room） → year_bucket × 萬（調整前/調整/調整後）
  §B 同上 by (year_bucket, ng_code) → 對 golden 邊年
  §C ADR/贈房/超過ADR/房晚 關鍵字行 → year_bucket × horizontal_label × 調整前/後（睇調整方向同大細）
  §D Comp房間(H_HOTEL_ROOM) by year_bucket（對 golden 客房 per 年）

Run（Windows）:  python scripts\inspect_galaxy_room.py            # galaxy
                 python scripts\inspect_galaxy_room.py wynn
出 results\inspect_galaxy_room_<ent>.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
ENT = sys.argv[1] if len(sys.argv) > 1 else "galaxy"
L: list[str] = [f"# inspect_galaxy_room — {ENT} comp room + ADR 調整追蹤（萬）"]


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    if not CSV.exists():
        L.append("!! csv 揾唔到"); _w(); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    df = df[df["entity"].astype(str) == ENT].copy()
    df["post"] = pd.to_numeric(df.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    df["pre"] = pd.to_numeric(df.get("調整前_萬", 0), errors="coerce").fillna(0.0)
    df["adj"] = pd.to_numeric(df.get("調整_萬", 0), errors="coerce").fillna(0.0)
    df["yb"] = df["year_bucket"].astype(str)
    for c in ("項目組H", "horizontal_id", "horizontal_label", "ng_code", "account_desc",
              "description", "科目明細", "comp_type"):
        if c not in df.columns:
            df[c] = ""
        df[c] = _s(df[c])
    df["blob"] = df["account_desc"] + " | " + df["description"] + " | " + df["科目明細"] + " | " + df["項目組H"]

    # §A 項目組H 含 room
    room = df[df["項目組H"].str.contains("room", case=False, na=False)]
    L.append(f"\n{'='*72}\n## §A 項目組H 含 'room' 嘅行（{len(room):,}行）→ by year_bucket（前/調整/後 萬）")
    g = room.groupby("yb")[["pre", "adj", "post"]].sum()
    for yb, r in g.iterrows():
        L.append(f"   {yb:<9} 前={r.pre:>10,.0f}  調整={r.adj:>10,.0f}  後={r.post:>10,.0f}")
    L.append(f"   {'總':<9} 前={room['pre'].sum():>10,.0f}  調整={room['adj'].sum():>10,.0f}  後={room['post'].sum():>10,.0f}")

    # §B by (yb, ng)
    L.append(f"\n{'='*72}\n## §B 項目組H room by (year_bucket, NG)（調整後萬，top）")
    g2 = room.groupby(["yb", "ng_code"])["post"].sum().reset_index().sort_values("post", key=lambda s: s.abs(), ascending=False)
    for _, r in g2.head(25).iterrows():
        L.append(f"   {r['yb']:<9} {r['ng_code']:<6} {r['post']:>10,.0f}")

    # §C ADR / 贈房 / 超過ADR / 房晚 關鍵字
    adr = df[df["blob"].str.contains("ADR|贈房|超過ADR|超過 ADR|房晚|comp room|Comp Room|客房", case=False, na=False, regex=True)]
    L.append(f"\n{'='*72}\n## §C ADR/贈房/房晚/客房 關鍵字行（{len(adr):,}行）→ by (year_bucket, H)（前/後 萬）")
    g3 = adr.groupby(["yb", "horizontal_label"])[["pre", "post"]].sum().reset_index().sort_values("post", key=lambda s: s.abs(), ascending=False)
    for _, r in g3.head(25).iterrows():
        L.append(f"   {r['yb']:<9} {r['horizontal_label'][:12]:<12} 前={r['pre']:>9,.0f}  後={r['post']:>9,.0f}")
    L.append(f"   ── ADR 類 sample account_desc ──")
    for a in sorted(set(adr["account_desc"].tolist()))[:18]:
        if a:
            L.append(f"       {a[:50]}")

    # §D H_HOTEL_ROOM by yb
    hr = df[df["horizontal_id"] == "H_HOTEL_ROOM"]
    L.append(f"\n{'='*72}\n## §D 現 Comp房間(H_HOTEL_ROOM) by year_bucket（調整後萬）— 對 golden 客房 per 年")
    for yb, v in hr.groupby("yb")["post"].sum().items():
        L.append(f"   {yb:<9} {v:>10,.0f}")
    L.append(f"   {'總':<9} {hr['post'].sum():>10,.0f}")
    _w()


def _w():
    out = ROOT / "results" / f"inspect_galaxy_room_{ENT}.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
