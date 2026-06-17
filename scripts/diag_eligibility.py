r"""Report-only：用 categories.yml 真 eligible_verticals，睇 enforcement 會 reset 邊啲掛錯枝。
NOT 1:1 —— 每 NG 一組多個 eligible V，只有完全唔喺組嘅 V 先算掛錯枝。

睇兩樣：
  ① 每 NG eligible 集有幾多個 V（證明唔係 1:1）
  ② 掛錯枝（V 唔喺該 NG eligible 集）by 金額 —— 即 enforcement 會郁嘅

Run:  python scripts/diag_eligibility.py
出 results/diag_eligibility.txt  ← paste 返嚟（睇真會 reset 邊啲先決定 apply）
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd, yaml

ROOT = Path(__file__).resolve().parent.parent
CSV = ROOT / "tableau_combined_25.csv"
CATS = ROOT / "conf" / "base" / "categories.yml"


def main():
    L = ["# diag_eligibility — report-only：enforcement 會 reset 邊啲掛錯枝（用 categories.yml eligible 集）"]
    if not CSV.exists() or not CATS.exists():
        L.append("!! csv 或 categories.yml 揾唔到"); _w(L); return
    cats = yaml.safe_load(CATS.read_text(encoding="utf-8")) or {}
    vlabel = {v["id"]: v.get("label", v["id"]) for v in (cats.get("verticals") or [])}
    ngcats = cats.get("ng_categories") or {}
    CROSS_CUTTING = {"V_PROPERTY_UPGRADE", "V_PUBLIC_FACILITY", "V_COMMUNITY"}  # 任何 NG 合理，永不 reset
    eligible = {ng: (set(d.get("eligible_verticals") or []) | CROSS_CUTTING) for ng, d in ngcats.items()}
    nglabel = {ng: d.get("label", "") for ng, d in ngcats.items()}

    L.append("\n① 每 NG eligible 集大小（證明唔係 1:1）：")
    for ng in sorted(eligible, key=lambda x: int(x[2:]) if x[2:].isdigit() else 99):
        evs = [vlabel.get(v, v) for v in eligible[ng]]
        L.append(f"   {ng:<5}{nglabel.get(ng,'')[:6]:<6} {len(evs)} 個: {evs}")

    df = pd.read_csv(CSV, dtype=str, keep_default_na=False)
    df["amt"] = pd.to_numeric(df["amount_mop"], errors="coerce").fillna(0.0).abs()
    vid = df["vertical_id"].astype(str).str.strip()
    ng = df["ng_code"].astype(str).str.strip()
    adj = df["v_調整"].astype(str).str.strip() if "v_調整" in df.columns else pd.Series("", index=df.index)
    # 掛錯枝：V 非空、未經調整(原始真V)、NG 有 eligible 集、V 唔喺集（已含 cross-cutting）
    bad = [bool(v) and f == "" and (n in eligible) and (v not in eligible[n]) and v != "V_OTHER"
           for v, n, f in zip(vid.tolist(), ng.tolist(), adj.tolist())]
    bad = pd.Series(bad, index=df.index)
    L.append(f"\n② 掛錯枝行: {int(bad.sum()):,}  Σ={df.loc[bad,'amt'].sum()/1e6:,.1f}M（總 {df['amt'].sum()/1e6:,.0f}M 嘅 {df.loc[bad,'amt'].sum()/df['amt'].sum()*100:.1f}%）")
    g = (df[bad].groupby(["entity", "ng_code", "ng_label", "vertical_label"])["amt"]
            .sum().reset_index().sort_values("amt", ascending=False))
    L.append("   ── 會 reset 嘅 combo（按金額 top 30）：entity | NG | V掛錯 | Σ萬 ──")
    for r in g.head(30).itertuples():
        L.append(f"     {r.entity:<7} {r.ng_code:<5}{str(r.ng_label)[:6]:<6} ← {str(r.vertical_label)[:12]:<12} {r.amt/1e4:>9,.0f}萬")
    L.append("\n   ── per entity 掛錯枝 Σ ──")
    for e, sub in df[bad].groupby("entity"):
        L.append(f"     {e:<7} {len(sub):>7,} 行  Σ={sub['amt'].sum()/1e6:>7,.1f}M")
    _w(L)


def _w(L):
    out = ROOT / "results" / "diag_eligibility.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
