r"""喺 JE 內揾返房 comp（唔加外部）。galaxy 房 golden 30,188/32,356 但 JE H_HOTEL_ROOM 得 ~5,479/16,707。
睇所有房相關行（description/account_desc/項目組H 含 room）：而家 tag 咗去邊 H、gross(調整前) vs net(調整後/amount_mop)、take/netoff flag。
搵 (1) tag 錯 H 嘅房行  (2) netting 令房 comp 縮水。

Run（Windows）:  python scripts\inspect_room_in_je.py galaxy
                 python scripts\inspect_room_in_je.py sjm
出 results\inspect_room_in_je_<ent>.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
ENTS = sys.argv[1:] or ["galaxy"]
ROOMKW = r"Comp Room|Comp Rooms|Comp Lodging|客房|房晚|Hotel Contra|Lodging|Room Charge|Room Tax|Room Service|COMP ROOM|COMPLIMENTARY ROOM"


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def run(ent, df):
    L = [f"# inspect_room_in_je — {ent}（揾 JE 內房 comp；gross=調整前, net=調整後/amount_mop）"]
    e = df[df["entity"].astype(str) == ent].copy()
    pre = pd.to_numeric(e.get("調整前_萬", 0), errors="coerce").fillna(0.0)
    post = pd.to_numeric(e.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    amt = pd.to_numeric(e.get("amount_mop", 0), errors="coerce").fillna(0.0) / 1e4
    e["pre"] = pre; e["post"] = post; e["amt"] = amt
    e["yb"] = e["year_bucket"].astype(str); e["y2"] = e["yb"].str[:2]
    for c in ("horizontal_label", "項目組H", "account_desc", "description", "take_flag", "netoff_flag", "final_capex_opex"):
        e[c] = _s(e[c]) if c in e.columns else ""
    e["blob"] = e["description"] + " " + e["account_desc"] + " " + e["項目組H"]
    room = e[e["blob"].str.contains(ROOMKW, case=False, regex=True, na=False)]
    L.append(f"\n房相關行：{len(room):,}行 | 調整前Σ={room['pre'].sum():,.0f} | 調整後Σ={room['post'].sum():,.0f} | amount_mopΣ={room['amt'].sum():,.0f}萬")

    # by spent-year prefix
    for yr in ("23", "24", "25"):
        sub = room[room.y2 == yr]
        if not len(sub):
            continue
        L.append(f"\n{'='*60}\n## {ent} {yr}（prefix）房相關 調整前={sub['pre'].sum():,.0f} 調整後={sub['post'].sum():,.0f} amount_mop={sub['amt'].sum():,.0f}")
        L.append(f"  ── 而家 tag 咗去邊 H（調整後）──")
        g = sub.groupby("horizontal_label")["post"].sum().reset_index().sort_values("post", key=lambda s: s.abs(), ascending=False)
        for _, r in g.iterrows():
            if abs(r["post"]) >= 30:
                L.append(f"     {r['horizontal_label'][:14]:<14} {r['post']:>9,.0f}")
        L.append(f"  ── 非 H_HOTEL_ROOM 嘅房行（搬入候選）by 項目組H+desc ──")
        nonr = sub[sub["horizontal_label"].ne("Comp房間")]
        gg = nonr.groupby(["horizontal_label", "項目組H", "account_desc"])["post"].sum().reset_index().sort_values("post", key=lambda s: s.abs(), ascending=False)
        for _, r in gg.head(12).iterrows():
            if abs(r["post"]) >= 30:
                L.append(f"     現H={r['horizontal_label'][:8]:<8} 項H={r['項目組H'][:14]:<14} {r['account_desc'][:22]:<22} {r['post']:>8,.0f}")
        # netting 線索：gross vs net
        L.append(f"  ── netting 線索（gross 調整前 vs net 調整後 差距）──")
        L.append(f"     調整前(gross)={sub['pre'].sum():,.0f}  調整後(net)={sub['post'].sum():,.0f}  差={sub['pre'].sum()-sub['post'].sum():,.0f}")
        # take/netoff flag 分佈
        for fcol in ("take_flag", "netoff_flag"):
            if sub[fcol].astype(bool).any():
                vc = sub.groupby(fcol)["post"].sum()
                L.append(f"     [{fcol}] " + " | ".join(f"{k or '∅'}={v:,.0f}" for k, v in vc.items()))
    p = ROOT / "results" / f"inspect_room_in_je_{ent}.txt"; p.parent.mkdir(exist_ok=True)
    p.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"wrote {p.relative_to(ROOT)}\n")


def main():
    if not CSV.exists():
        print("!! csv 揾唔到"); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    for ent in ENTS:
        run(ent, df)


if __name__ == "__main__":
    main()
