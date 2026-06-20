r"""raw-column hint 深掃 —— 用 CSV 入面嘅原始欄（description/成本中心/vendor/項目組H/調整二級）
揾 bottom-up account 揾唔到嘅 hints。專攻 galaxy 23 staff（adjustment 結構）+ galaxy 房/票 comp。

Run（Windows）:  python scripts\inspect_raw_hints.py
出 results\inspect_raw_hints.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
L = ["# inspect_raw_hints — galaxy 23 staff + 房/票 comp 嘅 raw-column 信號"]


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    if not CSV.exists():
        print("!! csv 揾唔到"); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    post = pd.to_numeric(df.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    pre = pd.to_numeric(df.get("調整前_萬", 0), errors="coerce").fillna(0.0)
    df["post"] = post; df["pre"] = pre; df["adj"] = pre - post
    df["yb"] = df["year_bucket"].astype(str)
    df["ent"] = df["entity"].astype(str)
    for c in ("horizontal_id", "horizontal_label", "final_capex_opex", "account_code", "account_desc",
              "description", "成本中心", "vendor", "項目組H", "項目組V", "調整一級", "調整二級", "source", "ng_code"):
        df[c] = _s(df[c]) if c in df.columns else ""
    df["code"] = df["account_code"].str.replace(r"\.0$", "", regex=True)
    g = df[df.ent == "galaxy"]

    # ── A. galaxy 23 source 分佈（opex labor vs all）──
    g23 = g[g.yb == "23"]
    L.append(f"\n{'='*64}\n## A. galaxy 23 source 分佈（萬，調整後）")
    for src, s in g23.groupby("source"):
        lab = s[s.horizontal_id.eq("H_LABOR")]["post"].sum()
        L.append(f"   source={src[:30]:<30} 全opex={s[~s.final_capex_opex.eq('Capex')]['post'].sum():>9,.0f} | H_LABOR={lab:>8,.0f}")

    # ── B. galaxy 23 H_LABOR 嘅 adjustment 結構（−50,723 喺邊；睇 over-adjust）──
    lab23 = g23[g23.horizontal_id.eq("H_LABOR")]
    L.append(f"\n{'='*64}\n## B. galaxy 23 H_LABOR adjustment（前={lab23['pre'].sum():,.0f} 調整={lab23['adj'].sum():,.0f} 後={lab23['post'].sum():,.0f}；golden 17,082）")
    L.append("   by 調整二級（adjustment 原因）:")
    for k, s in lab23.groupby("調整二級"):
        if abs(s["adj"].sum()) >= 50 or abs(s["post"].sum()) >= 50:
            L.append(f"     調二={k[:30]:<30} 前={s['pre'].sum():>9,.0f} 調整={s['adj'].sum():>9,.0f} 後={s['post'].sum():>8,.0f}")

    # ── C. galaxy 23 opex 非人工 by 成本中心（staff dept 信號）──
    o23 = g23[~g23.final_capex_opex.eq("Capex") & ~g23.horizontal_id.eq("H_LABOR")]
    L.append(f"\n{'='*64}\n## C. galaxy 23 opex 非人工 by 成本中心 top（staff 部門信號）")
    gc = o23.groupby("成本中心")["post"].sum().reset_index().sort_values("post", key=lambda s: s.abs(), ascending=False)
    for _, r in gc.head(12).iterrows():
        L.append(f"     成本中心={r['成本中心'][:34]:<34} {r['post']:>9,.0f}")

    # ── D. galaxy 房 comp 信號：description/account 含 room/lodging/hotel 但非 H_HOTEL_ROOM（23+24）──
    for yv in ("23", "24"):
        gy = g[g.yb == yv]
        room_kw = r"Lodging|Room|Hotel|客房|Suite|Accommod|Guest|Marriott|Ritz|Banyan|Andaz|JW "
        cand = gy[gy["description"].str.contains(room_kw, case=False, regex=True, na=False) | gy["account_desc"].str.contains(room_kw, case=False, regex=True, na=False)]
        cand = cand[~cand.horizontal_id.eq("H_HOTEL_ROOM")]
        L.append(f"\n## D{yv}. galaxy {yv} 房信號 非客房H（golden 客房 {'30,188' if yv=='23' else '32,356'}，我 {'5,479' if yv=='23' else '16,708'}）")
        gd = cand.groupby(["horizontal_label", "account_desc"])["post"].sum().reset_index()
        gd = gd[gd["post"].abs() >= 100].sort_values("post", key=lambda s: s.abs(), ascending=False)
        for _, r in gd.head(8).iterrows():
            L.append(f"     現H={r['horizontal_label'][:9]:<9} {r['account_desc'][:30]:<30} {r['post']:>8,.0f}")

    # ── E. galaxy 贈票 信號：Ticket/Admission/Club/Spa/Show 非 H_COMP_TICKET（23+24）──
    for yv in ("23", "24"):
        gy = g[g.yb == yv]
        tkw = r"Ticket|Admission|Club Use|Spa|Show|門票|入場|Concert|Entertainment Ticket"
        cand = gy[(gy["description"].str.contains(tkw, case=False, regex=True, na=False) | gy["account_desc"].str.contains(tkw, case=False, regex=True, na=False)) & ~gy.horizontal_id.eq("H_COMP_TICKET")]
        L.append(f"\n## E{yv}. galaxy {yv} 贈票信號 非贈票H（golden 贈票 {'3,990' if yv=='23' else '4,024'}，我 ~3/319）")
        ge = cand.groupby(["horizontal_label", "account_desc"])["post"].sum().reset_index()
        ge = ge[ge["post"].abs() >= 50].sort_values("post", key=lambda s: s.abs(), ascending=False)
        for _, r in ge.head(8).iterrows():
            L.append(f"     現H={r['horizontal_label'][:9]:<9} {r['account_desc'][:30]:<30} {r['post']:>8,.0f}")

    out = ROOT / "results" / "inspect_raw_hints.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
