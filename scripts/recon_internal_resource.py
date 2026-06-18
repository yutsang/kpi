r"""對數：內部資源 golden（root\internal_resource_golden.xlsx）逐格 vs 我哋 tableau_combined_25.csv。

項目組要求每 entity × year × NG（bucket）嘅 staff cost / comp（5 拆 + 合計）/ capex / opex
大致相同。呢個 script 將 golden 每一格 同 我哋對應數比，列出差異。

Golden 欄 → 我哋 horizontal_id：
  internal_staff_cost → H_LABOR        客房支出 → H_HOTEL_ROOM   會場支出 → H_VENUE
  食品飲料支出 → H_FNB                 贈票支出 → H_COMP_TICKET   comp_其他 → H_COMP_OTHER
  comp合計 = 5 個 comp 之和             capex/opex → final_capex_opex
金額用「調整後_萬」（= 報告投資金額，萬澳門元，同 golden 單位）。

year 對應：golden「23」↔ 我哋 year_bucket「23」；golden「24」↔「24」。同時印「24+24_23SY」
合計作參考（睇 golden 24 有冇含期後）。galaxy24 / mgm24 staff 只有總數 → 睇 entity×year TOTAL 行。

Run（Windows）:  python scripts\recon_internal_resource.py
出 results\recon_internal_resource.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd
import re

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
GOLD = ROOT / "internal_resource_golden.xlsx"
THR = 50.0  # 萬：只列 |diff|>THR 嘅格

# golden NG 中文名 → ng_code（容錯：博彩娛樂/博彩項目 都 →NG0）
NG_NAME2CODE = {
    "博彩": "NG0", "吸引外國客源": "NG1", "會議展覽": "NG2", "娛樂表演": "NG3", "體育盛事": "NG4",
    "文化藝術": "NG5", "健康養生": "NG6", "主題遊樂": "NG7", "美食之都": "NG8", "社區旅遊": "NG9",
    "海上旅遊": "NG10", "其他": "NG11",
}
# 我哋 H → golden metric key
H2KEY = {"H_LABOR": "staff", "H_HOTEL_ROOM": "room", "H_VENUE": "venue", "H_FNB": "fnb",
         "H_COMP_TICKET": "ticket", "H_COMP_OTHER": "comp_other"}
COMP_KEYS = ["room", "venue", "fnb", "ticket", "comp_other"]
METRICS = ["staff", "room", "venue", "fnb", "ticket", "comp_other", "comp_total", "capex", "opex"]
MLAB = {"staff": "人工", "room": "客房", "venue": "會場", "fnb": "食飲", "ticket": "贈票",
        "comp_other": "comp其他", "comp_total": "comp合計", "capex": "capex", "opex": "opex"}

L: list[str] = ["# recon_internal_resource —— 內部資源 golden vs 我哋（萬澳門元）"]


def _num(x) -> float:
    s = str(x).strip().replace(",", "")
    if s in ("", "-", "nan", "None", "NaN", "–"):
        return 0.0
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    try:
        v = float(s)
    except Exception:
        return 0.0
    return -v if neg else v


def _ng_from_name(s) -> str:
    s = str(s)
    for k, c in NG_NAME2CODE.items():
        if k in s:
            return c
    return ""


def load_golden() -> pd.DataFrame:
    raw = pd.read_excel(GOLD, dtype=str, header=None)
    # 揾 header row（含 entity/year/capex/opex 字眼）
    hdr = None
    for i in range(min(8, len(raw))):
        row = " ".join(str(x) for x in raw.iloc[i].tolist())
        if ("entity" in row.lower()) and ("opex" in row.lower() or "capex" in row.lower()):
            hdr = i; break
    df = pd.read_excel(GOLD, dtype=str, header=hdr if hdr is not None else 0)
    df.columns = [str(c).strip() for c in df.columns]

    def col(*kw):
        for c in df.columns:
            cl = str(c).lower().replace(" ", "")
            if any(k.lower().replace(" ", "") in cl for k in kw):
                return c
        return None
    cmap = {
        "entity": col("entity"), "year": col("year", "年"),
        "capex": col("capex"), "opex": col("opex"), "ng": col("NG", "範疇", "性質", "bucket"),
        "staff": col("staff", "人工", "internal_staff"),
        "room": col("客房"), "venue": col("會場", "会场"), "fnb": col("食品", "飲料", "饮料"),
        "ticket": col("贈票", "赠票", "門票", "门票", "ticket"),
        "comp_other": col("comp_其他", "comp其他", "comp_other"),
        "comp_total": col("comp合計", "comp合计", "comp_total", "合計", "合计"),
    }
    L.append(f"  golden 欄對應: " + ", ".join(f"{k}={v}" for k, v in cmap.items() if v))
    out = pd.DataFrame()
    out["entity"] = df[cmap["entity"]].astype(str).str.strip().str.lower().ffill() if cmap["entity"] else ""
    out["year"] = df[cmap["year"]].astype(str).str.strip().str.extract(r"(\d{2})")[0] if cmap["year"] else ""
    out["ng_code"] = df[cmap["ng"]].map(_ng_from_name) if cmap["ng"] else ""
    for k in ["capex", "opex", "staff", "room", "venue", "fnb", "ticket", "comp_other"]:
        out[k] = df[cmap[k]].map(_num) if cmap.get(k) else 0.0
    out["comp_total"] = out[COMP_KEYS].sum(axis=1)
    out = out[out["entity"].isin(["galaxy", "sjm", "wynn", "vml", "melco", "mgm"]) & out["year"].isin(["23", "24"])]
    return out


def our_metrics(df: pd.DataFrame) -> pd.DataFrame:
    """df = tableau csv 某 entity 某 bucket 子集 → per ng_code 各 metric（萬）。"""
    amt = pd.to_numeric(df.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    h = df.get("horizontal_id", pd.Series("", index=df.index)).astype(str).str.strip()
    co = df.get("final_capex_opex", pd.Series("", index=df.index)).astype(str).str.strip()
    ng = df.get("ng_code", pd.Series("", index=df.index)).astype(str).str.strip()
    base = pd.DataFrame({"ng_code": ng, "amt": amt, "h": h, "co": co})
    rows = []
    for ngc, g in base.groupby("ng_code"):
        r = {"ng_code": ngc}
        for hid, key in H2KEY.items():
            r[key] = g.loc[g["h"].eq(hid), "amt"].sum()
        r["comp_total"] = sum(r[k] for k in COMP_KEYS)
        r["capex"] = g.loc[g["co"].eq("Capex"), "amt"].sum()
        r["opex"] = g.loc[g["co"].eq("Opex"), "amt"].sum()
        rows.append(r)
    return pd.DataFrame(rows)


def main():
    if not CSV.exists():
        L.append("!! tableau_combined_25.csv 揾唔到"); _w(); return
    if not GOLD.exists():
        L.append(f"!! {GOLD.name} 揾唔到（root 放 internal_resource_golden.xlsx）"); _w(); return
    gold = load_golden()
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    df["_yb"] = df["year_bucket"].astype(str).str.strip()

    diffs = []  # (entity, year, ng, metric, gold, ours, diff)
    for (ent, yr), gsub in gold.groupby(["entity", "year"]):
        oursub = df[(df["entity"] == ent) & (df["_yb"] == yr)]
        om = our_metrics(oursub).set_index("ng_code")
        # 同時計 24+24_23SY 作參考
        alt = df[(df["entity"] == ent) & (df["_yb"].str.startswith(yr))]
        oalt = our_metrics(alt).set_index("ng_code")

        L.append(f"\n{'='*78}\n## {ent} {yr}（左=golden 右=我哋[{yr}] | Δ；括號=我哋[{yr}+SY]）")
        # entity×year TOTAL（galaxy24/mgm24 staff 只有總數 → 睇呢度）
        gt = {m: gsub[m].sum() for m in METRICS}
        ot = {m: om[m].sum() if m in om.columns else 0.0 for m in METRICS}
        at = {m: oalt[m].sum() if m in oalt.columns else 0.0 for m in METRICS}
        L.append("  ── TOTAL ──")
        for m in METRICS:
            d = ot[m] - gt[m]
            flag = "  ⚠" if abs(d) > THR else ""
            L.append(f"     {MLAB[m]:<8} golden={gt[m]:>10,.0f}  我哋={ot[m]:>10,.0f}  Δ={d:>10,.0f}  (含SY={at[m]:,.0f}){flag}")
        # per-NG
        for _, gr in gsub.iterrows():
            ngc = gr["ng_code"]
            if not ngc:
                continue
            orow = om.loc[ngc] if ngc in om.index else None
            arow = oalt.loc[ngc] if ngc in oalt.index else None
            for m in METRICS:
                gv = float(gr[m]); ov = float(orow[m]) if orow is not None else 0.0
                av = float(arow[m]) if arow is not None else 0.0
                d = ov - gv
                if abs(d) > THR:
                    diffs.append((ent, yr, ngc, m, gv, ov, av, d))

    # 大差異榜（per cell |Δ|>THR）
    L.append(f"\n{'='*78}\n## 逐格大差異（|Δ|>{THR:.0f}萬，按 |Δ| 排）")
    diffs.sort(key=lambda t: -abs(t[7]))
    L.append(f"   {'entity':<7}{'yr':<4}{'NG':<6}{'metric':<8}{'golden':>10}{'我哋':>10}{'Δ':>10}{'含SY':>10}")
    for ent, yr, ngc, m, gv, ov, av, d in diffs[:120]:
        L.append(f"   {ent:<7}{yr:<4}{ngc:<6}{MLAB[m]:<8}{gv:>10,.0f}{ov:>10,.0f}{d:>10,.0f}{av:>10,.0f}")
    L.append(f"\n   （共 {len(diffs)} 格超 {THR:.0f}萬；只顯示頭 120）")
    _w()


def _w():
    out = ROOT / "results" / "recon_internal_resource.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
