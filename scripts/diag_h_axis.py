r"""H-axis 診斷 (全 6 家)：答緊呢幾條 —
  #1 opex/capex 欄 format 完唔完好（值有冇亂：Capex/CAPEX/Opex/空…）
  #5 邊啲行冇 H label（melco ~19億）→ 數 + 佔額 + top account
  #7 人工成本(H_LABOR) 點解有負數 → Σ + top account/desc
  #3 Comp活動場地(H_VENUE) 有冇 capex（應該去建設/器具）→ by capex/opex
  #2 藝人演出費(H_PERFORMER) 入面有冇其他合約成本（licence/贊助）→ top account_desc
  #4 utility/水電 而家喺邊個 H（應該去其他）→ by H label

用 horizontal_id（穩定，唔受 label 改名影響）。amt = conf amount /1e4 (萬)。
Run:  python scripts/diag_h_axis.py            # all 6
      python scripts/diag_h_axis.py melco mgm
出 results/diag_h_axis.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd
import yaml

ROOT = Path(__file__).resolve().parent.parent
ENT = {"galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
       "vml": "company_4", "melco": "company_5", "mgm": "company_6"}
UTIL_KW = ["utilit", "電費", "水電", "水費", "electric", "eletric", "chilled water",
           "power charge", "gas & oil", "煤氣", "燃氣"]


def _conf(com):
    return yaml.safe_load((ROOT / "conf" / com / "parameters.yml").read_text(encoding="utf-8"))


def _col(df, cands):
    low = {str(c).strip().lower(): c for c in df.columns}
    for c in cands:
        if not c: continue
        if c in df.columns: return c
        if str(c).strip().lower() in low: return low[str(c).strip().lower()]
    return None


def _amtcol(df, cfg):
    a = (cfg.get("columns") or {}).get("amount", "")
    return _col(df, [a, "Debit minus Credit", "MOP Amt", "Amount - Amended", "amount_mop", "amount",
                     "Val/COArea Crcy", "Entry Voucher Amount/ Expense Amount ", "adjusted_amount"])


def _top(L, s, amt, n=8, unit="萬"):
    g = pd.DataFrame({"k": s.astype(str).fillna(""), "a": amt}).groupby("k")["a"].agg(["sum", "size"])
    g = g.reindex(g["sum"].abs().sort_values(ascending=False).index)
    for k, r in g.head(n).iterrows():
        L.append(f"        {str(k)[:46]:<46} {r['sum']:>11,.0f}{unit}  ({int(r['size']):,}行)")


def main():
    want = [a.lower() for a in sys.argv[1:]] or list(ENT)
    L = ["# diag_h_axis — opex/capex format / no-H / 人工負數 / Comp場地 capex / 藝人演出費內容 / utility 位置"]
    for alias in want:
        com = ENT.get(alias)
        if not com: continue
        L.append(f"\n{'='*74}\n## {alias}")
        p = ROOT / "data" / alias / "output" / f"{com}_kpi_report.parquet"
        if not p.exists():
            L.append("   (no kpi_report parquet — 未跑)"); continue
        cfg = _conf(com)
        df = pd.read_parquet(p)
        Hid = df["horizontal_id"].astype(str).fillna("") if "horizontal_id" in df.columns else pd.Series("", index=df.index)
        Hlb = df["horizontal_label"].astype(str).fillna("") if "horizontal_label" in df.columns else pd.Series("", index=df.index)
        amt = pd.to_numeric(df[_amtcol(df, cfg)], errors="coerce").fillna(0.0) / 1e4
        cox = _col(df, ["final_capex_opex", (cfg.get("columns") or {}).get("capex_opex", ""), "Capex_Opex", "capex_opex", "CAPEX_OPEX"])
        adc = _col(df, [(cfg.get("columns") or {}).get("account_desc", ""), "account_desc", "A/C Name", "Nature of Expenses", "Ledger Hierarchy Level 5"])
        acc = _col(df, [(cfg.get("columns") or {}).get("account_code", ""), "account_code", "Account", "Ledger Account"])
        rp = df["report_period"].astype(str) if "report_period" in df.columns else pd.Series("", index=df.index)
        desc = df[adc].astype(str).fillna("") if adc else pd.Series("", index=df.index)

        # ── #1 opex/capex format ──
        L.append(f"   #1 capex/opex 欄 '{cox}' 值分佈 (format check):")
        if cox:
            vc = df[cox].astype(str).fillna("(空)").replace({"": "(空)", "nan": "(空)"}).value_counts()
            for k, v in vc.head(12).items():
                L.append(f"        {str(k)[:30]:<30} {v:,}行")
        else:
            L.append("        (揾唔到 capex/opex 欄!)")

        # ── #5 no-H label ──
        noH = Hid.isin(["", "nan", "None"]) | Hlb.isin(["", "(未分類)", "nan", "None"])
        L.append(f"\n   #5 冇 H label: {int(noH.sum()):,} 行  Σ={amt[noH].sum():,.0f}萬  (= melco 嗰 ~19億?)")
        if noH.any() and adc:
            L.append("      top account_desc:")
            _top(L, desc[noH], amt[noH])
        if noH.any():
            L.append("      by bucket: " + ", ".join(f"{k}={v:,.0f}萬" for k, v in amt[noH].groupby(rp[noH]).sum().items() if abs(v) > 1))

        # ── #7 人工成本 (H_LABOR) 負數 ──
        lab_neg = (Hid == "H_LABOR") & (amt < 0)
        L.append(f"\n   #7 人工成本(H_LABOR) 負數: {int(lab_neg.sum()):,} 行  Σ={amt[lab_neg].sum():,.0f}萬 (H_LABOR 總 Σ={amt[Hid=='H_LABOR'].sum():,.0f}萬)")
        if lab_neg.any() and adc:
            _top(L, desc[lab_neg], amt[lab_neg])

        # ── #3 Comp活動場地 (H_VENUE) capex/opex ──
        ven = Hid == "H_VENUE"
        L.append(f"\n   #3 Comp活動場地(H_VENUE) 總 Σ={amt[ven].sum():,.0f}萬，by capex/opex:")
        if ven.any() and cox:
            for k, v in amt[ven].groupby(df[cox].astype(str)).sum().items():
                flag = "  ← capex 應去建設/器具" if ("cap" in str(k).lower()) else ""
                L.append(f"        {str(k)[:20]:<20} {v:,.0f}萬{flag}")

        # ── #2 藝人演出費 (H_PERFORMER) 內容 (揾非藝人嘅合約: licence/贊助) ──
        perf = Hid == "H_PERFORMER"
        L.append(f"\n   #2 藝人演出費(H_PERFORMER) 總 Σ={amt[perf].sum():,.0f}萬  top account_desc (睇下有冇 licence/贊助/版權):")
        if perf.any() and adc:
            _top(L, desc[perf], amt[perf])

        # ── #4 utility/水電 而家喺邊個 H ──
        util = desc.str.lower().apply(lambda s: any(k in s for k in UTIL_KW))
        L.append(f"\n   #4 utility/水電 ({int(util.sum()):,} 行 Σ={amt[util].sum():,.0f}萬) 而家落咗邊個 H (應該→其他):")
        if util.any():
            for k, v in amt[util].groupby(Hlb[util]).sum().items():
                L.append(f"        {str(k)[:24]:<24} {v:,.0f}萬")
    out = ROOT / "results" / "diag_h_axis.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
