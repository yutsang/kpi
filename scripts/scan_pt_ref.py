r"""Scan 全 18 個 raw 檔（6 entity × 各 yearly_source），揾項目組自己嘅分類 reference 欄
（H類/V類/comp/類別/分類/性質/費用大类/識別…），列出欄名 + distinct 值，做對比參考。

⚠ 唔照抄（taxonomy 唔同）—— 只係睇佢哋有冇分類欄、我哋邊度可參考 / 有冇 contradiction。

Run:  python scripts/scan_pt_ref.py
出 results/scan_pt_ref.txt  ← paste 返嚟
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd, yaml

ROOT = Path(__file__).resolve().parent.parent
ENT = {"galaxy": 1, "sjm": 2, "wynn": 3, "vml": 4, "melco": 5, "mgm": 6}
# 分類-ref 欄名關鍵字（項目組自己嘅 H/V/comp/類別 等）
KW = ["分類", "類別", "类别", "性質", "性质", "進一步", "费用大类", "費用大类", "comp", "Comp", "COMP",
      "横向", "縱向", "类型", "類型", "標籤", "标签", "識別", "识别", "表演", "演出", "場館", "场馆",
      "nature", "Nature", "category", "Category", "comp费用", "建築", "建筑", "設施", "设施", "劃分", "划分"]


def _hit(col):
    c = str(col)
    return any(k in c for k in KW)


def main():
    L = ["# scan_pt_ref — 18 raw 檔嘅項目組分類 ref 欄（H/V/comp/類別…）+ distinct 值"]
    for ent, n in ENT.items():
        confp = ROOT / "conf" / f"company_{n}" / "parameters.yml"
        L.append(f"\n{'='*70}\n── {ent} ──")
        if not confp.exists():
            L.append("   !! conf 揾唔到"); continue
        cfg = yaml.safe_load(confp.read_text(encoding="utf-8")) or {}
        rawdir = ROOT / "data" / ent / "raw"
        ys = cfg.get("yearly_sources") or [{"year": cfg.get("raw_filename", "?"),
                                            "raw_filename": cfg.get("raw_filename"),
                                            "raw_sheet": cfg.get("raw_sheet", 0)}]
        seen = set()
        for y in ys:
            fn = y.get("raw_filename") or ""
            sh = y.get("raw_sheet", 0)
            hdr = y.get("header_row", cfg.get("header_row", 0))
            key = (fn, str(sh))
            if key in seen:
                continue
            seen.add(key)
            fp = rawdir / fn
            if not fp.exists():
                L.append(f"   ─ {fn}[{sh}] 揾唔到"); continue
            try:
                if isinstance(sh, list):
                    sh = sh[0]
                df = pd.read_excel(fp, sheet_name=sh, header=hdr, dtype=str, nrows=30000)
            except Exception as e:
                L.append(f"   ─ {fn}[{sh}] 讀唔到: {e}"); continue
            cols = [c for c in df.columns if _hit(c)]
            L.append(f"   ─ {fn} [{sh}]：{len(cols)} 個 ref 欄")
            for c in cols:
                s = df[c].astype(str).str.strip().replace({"nan": "", "None": ""})
                nb = s[s != ""]
                if len(nb) == 0:
                    continue
                vc = nb.value_counts()
                ex = [f"{k[:22]}({v})" for k, v in vc.head(10).items()]
                L.append(f"      「{str(c)[:30]}」 cov{len(nb)/len(df)*100:.0f}% distinct={nb.nunique()}: {ex}")
    _w(L)


def _w(L):
    out = ROOT / "results" / "scan_pt_ref.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
