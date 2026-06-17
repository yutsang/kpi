r"""答「為什麼一個 dicj 同一 bucket 有多 NG」—— dump 原始 databook 主題欄。
讀每家 kpi_report.parquet + conf columns.ng11_category（NG 來源欄），
對同一 (dicj_code, report_period) 列出**原始主題欄嘅唔同值 + 金額**，
睇實係咪項目組自己逐行標咗唔同主題（= 佢哋 data 粒度，唔係我哋填錯）。

Run:  python scripts/dump_ng_why.py
出 results/dump_ng_why.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd, yaml

ROOT = Path(__file__).resolve().parent.parent
ENT = {"galaxy": 1, "sjm": 2, "wynn": 3, "vml": 4, "melco": 5, "mgm": 6}
AMT = ["amount_mop", "Amount - Amended", "ledger_amount"]


def _s(x):
    return x.astype(str).str.strip().replace({"nan": "", "None": "", "NaN": ""})


def main():
    L = ["# dump_ng_why — 同一 (dicj, bucket) 嘅原始 databook 主題欄唔同值（睇 NG 多值根因）"]
    for ent, n in ENT.items():
        p = ROOT / "data" / ent / "output" / f"company_{n}_kpi_report.parquet"
        cf = ROOT / "conf" / f"company_{n}" / "parameters.yml"
        L.append(f"\n{'='*64}\n── {ent} ──")
        if not p.exists() or not cf.exists():
            L.append("   !! parquet/conf 揾唔到"); continue
        cfg = yaml.safe_load(cf.read_text(encoding="utf-8")) or {}
        tcol = (cfg.get("columns") or {}).get("ng11_category", "")
        df = pd.read_parquet(p)
        if tcol not in df.columns:
            tcol = next((c for c in ("ng_theme", "項目性質", "項目類型", "Section.1", "項目性質") if c in df.columns), None)
        if not tcol:
            L.append("   !! 主題欄揾唔到"); continue
        amtc = next((c for c in AMT if c in df.columns), None)
        amt = pd.to_numeric(df[amtc], errors="coerce").fillna(0.0) / 1e4 if amtc else pd.Series(0.0, index=df.index)
        dc = _s(df["dicj_code"]) if "dicj_code" in df.columns else pd.Series("", index=df.index)
        rp = _s(df["report_period"]) if "report_period" in df.columns else pd.Series("", index=df.index)
        th = _s(df[tcol].astype(str))
        L.append(f"   NG 來源欄 = '{tcol}'")
        g = pd.DataFrame({"d": dc, "b": rp, "t": th, "a": amt})
        g = g[g["d"] != ""]
        # 揾同一 (d,b) 有 >1 主題值嘅，按金額 top
        flagged = []
        for (d, b), sub in g.groupby(["d", "b"]):
            ts = sub.groupby("t")["a"].sum()
            ts = ts[ts.abs() > 0.5]
            if ts.nunique() > 1 and len(ts) > 1:
                flagged.append((d, b, ts.abs().sum(), ts.sort_values(ascending=False)))
        flagged.sort(key=lambda x: -x[2])
        L.append(f"   {len(flagged)} 個 (dicj,bucket) 主題欄有 >1 值（top 8）：")
        for d, b, tot, ts in flagged[:8]:
            L.append(f"     {d} [{b}]:")
            for t, a in ts.items():
                L.append(f"         「{str(t)[:34]}」 = {a:,.0f}萬")
    _w(L)


def _w(L):
    out = ROOT / "results" / "dump_ng_why.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
