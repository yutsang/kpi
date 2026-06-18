r"""診斷 staff cost + comp 內部資源 點解對唔上 golden（recon 大差異）。

recon 發現我哋 H_LABOR / comp(客房/會場/贈票) 比 golden 內部資源細好多。
tie 冇穿 → 啲錢喺數據入面，只係 H tag 錯（多數落咗 其他/建設 等）。
呢個 script 睇我哋有冇 flag 可以揾返佢哋再 re-tag：

  §A staff：每 entity×year，現 H_LABOR 金額 vs is_labor-flag 行嘅金額（+ 佢哋而家咩 H）。
  §B comp：comp_type / is_internal 各值 × 而家 H × 金額（睇 room/venue/ticket 落咗去邊）。
  §C 估算：如果 is_labor→人工、comp_type→對應 comp H，可以補返幾多。

Run（Windows）:  python scripts\inspect_comp_staff.py
出 results\inspect_comp_staff.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
ENTS = ["galaxy", "sjm", "wynn", "vml", "melco", "mgm"]
TRUE = {"1", "1.0", "y", "yes", "true", "t", "x", "是", "係", "labor", "staff"}
L: list[str] = ["# inspect_comp_staff —— staff/comp 內部資源點解對唔上 golden（萬=調整後_萬）"]


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def _truthy(series):
    s = series.astype(str).str.strip().str.lower()
    return ~s.isin(["", "0", "0.0", "nan", "none", "false", "f", "n", "no", "<na>"])


def main():
    if not CSV.exists():
        L.append("!! tableau_combined_25.csv 揾唔到"); _w(); return
    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    df["amt"] = pd.to_numeric(df.get("調整後_萬", 0), errors="coerce").fillna(0.0)
    df["y2"] = df["year_bucket"].astype(str).str[:2]
    for c in ("entity", "horizontal_label", "horizontal_id", "is_labor", "comp_type",
              "is_internal", "comp_scope", "scope", "account_desc"):
        if c not in df.columns:
            df[c] = ""
        df[c] = _s(df[c])

    have = [c for c in ("is_labor", "comp_type", "is_internal", "comp_scope", "scope") if df[c].ne("").any()]
    L.append(f"  可用 flag 欄（有非空值）：{have or '冇！'}")
    L.append(f"  comp_type 出現過嘅值：{sorted(set(df['comp_type'].unique()) - {''})[:20]}")
    L.append(f"  is_labor 出現過嘅值：{sorted(set(df['is_labor'].unique()) - {''})[:12]}")

    # ── §A staff ──
    L.append(f"\n{'='*74}\n## §A staff（人工）每 entity×year：現 H_LABOR vs is_labor-flag")
    lab_t = _truthy(df["is_labor"])
    for ent in ENTS:
        for y in ("23", "24", "25"):
            sub = df[(df.entity == ent) & (df.y2 == y)]
            if not len(sub):
                continue
            our = sub.loc[sub.horizontal_id == "H_LABOR", "amt"].sum()
            flag = sub[lab_t.loc[sub.index]]
            ft = flag["amt"].sum()
            if abs(our) < 1 and abs(ft) < 1:
                continue
            byh = flag.groupby("horizontal_label")["amt"].sum().sort_values(key=lambda s: s.abs(), ascending=False)
            extra = " | ".join(f"{k or '(空)'}={v:,.0f}" for k, v in byh.head(4).items())
            L.append(f"   {ent:<7}{y}  現H_LABOR={our:>8,.0f}萬  is_labor行={ft:>8,.0f}萬  →現H: {extra}")

    # ── §B comp_type / is_internal × 現 H ──
    L.append(f"\n{'='*74}\n## §B comp_type × 現 H（睇 room/venue/ticket 落咗去邊）")
    ct = df[df.comp_type.ne("")]
    if len(ct):
        g = ct.groupby(["comp_type", "horizontal_label"])["amt"].sum().reset_index().sort_values("amt", key=lambda s: s.abs(), ascending=False)
        for _, r in g.head(30).iterrows():
            L.append(f"   comp_type={r['comp_type'][:20]:<20} 現H={r['horizontal_label'][:12]:<12} {r['amt']:>9,.0f}萬")
    else:
        L.append("   （comp_type 全空 → 用唔到呢個 flag）")
    ii = df[_truthy(df["is_internal"])]
    if len(ii):
        L.append(f"\n   is_internal=true × 現 H：")
        g = ii.groupby("horizontal_label")["amt"].sum().sort_values(key=lambda s: s.abs(), ascending=False)
        for k, v in g.head(12).items():
            L.append(f"     {k or '(空)':<14} {v:,.0f}萬")

    # ── §C 各 entity comp H 現況（對 golden）──
    L.append(f"\n{'='*74}\n## §C 現有 comp H 各 entity×year 金額（對 golden 客房/會場/食飲/贈票/comp其他）")
    COMPH = {"H_HOTEL_ROOM": "客房", "H_VENUE": "會場", "H_FNB": "食飲",
             "H_COMP_TICKET": "贈票", "H_COMP_OTHER": "comp其他"}
    for ent in ENTS:
        for y in ("23", "24", "25"):
            sub = df[(df.entity == ent) & (df.y2 == y)]
            if not len(sub):
                continue
            vals = {lab: sub.loc[sub.horizontal_id == hid, "amt"].sum() for hid, lab in COMPH.items()}
            if sum(abs(v) for v in vals.values()) < 1:
                continue
            L.append(f"   {ent:<7}{y}  " + "  ".join(f"{lab}={v:,.0f}" for lab, v in vals.items()))
    _w()


def _w():
    out = ROOT / "results" / "inspect_comp_staff.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
