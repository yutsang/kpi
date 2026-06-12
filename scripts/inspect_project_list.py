"""Deep-inspect the 6 golden 投資項目清單 files (data/投資項目清單/) — per entity:
columns inventory, 清單序號/承批公司項目序號/項目名稱 samples, EVERY numeric column Σ, and
★ flags on columns whose Σ ≈ a golden bucket (萬→MOP ×1e4) — i.e. which 清單 columns
reproduce 25/25_24SY/25_23SY/24/24_23SY/23, so the 清單 can drive per-project tie-out.

Run (Windows):  python scripts/inspect_project_list.py
Output: prints + results/inspect_project_list.txt (paste back)
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
LDIR = ROOT / "data" / "投資項目清單"
GOLDEN_WAN = {  # 萬 MOP (user 2026-06-12)
    "galaxy": {"25": 326626, "25_24SY": 18766, "25_23SY": 3, "24": 248981, "24_23SY": 17489, "23": 119395},
    "sjm":    {"25": 124661, "25_24SY": 63594, "25_23SY": 7971, "24": 155980, "24_23SY": 59681, "23": 59346},
    "wynn":   {"25": 209699, "25_24SY": 5611, "25_23SY": 40, "24": 194885, "24_23SY": 4090, "23": 134168},
    "vml":    {"25": 252310, "25_24SY": 39, "25_23SY": 0, "24": 445636, "24_23SY": -841, "23": 134729},
    "melco":  {"25": 243470, "25_24SY": 13125, "25_23SY": 0, "24": 174252, "24_23SY": 26011, "23": 99522},
    "mgm":    {"25": 186142, "25_24SY": 64264, "25_23SY": 5527, "24": 149763, "24_23SY": 63985, "23": 95033},
}


def _num(s):
    return pd.to_numeric(s.astype("string").str.replace(",", "", regex=False), errors="coerce")


def _s(df, col):
    v = df.get(col)
    if v is None:
        return pd.Series("", index=df.index)
    if isinstance(v, pd.DataFrame):
        v = v.iloc[:, 0]
    return v.astype("string").fillna("").str.strip()


def main():
    L = ["# inspect_project_list — 6 golden 清單 deep dump (Σ欄 vs golden buckets)"]
    if not LDIR.exists():
        L.append("X data/投資項目清單 missing"); _w(L); return
    for p in sorted(LDIR.glob("*.xls*")):
        if p.name.startswith("~$"):
            continue
        ent = next((e for e in ("galaxy", "mgm", "melco", "sjm", "vml", "wynn") if e in p.name.lower()), None)
        print(f"[清單] reading {p.name[:46]}...", flush=True)
        try:
            xl = pd.ExcelFile(p)
            shn = next((s for s in xl.sheet_names if "database" in str(s).lower()
                        or "分項目數據簿" in str(s)), xl.sheet_names[0])
            probe = pd.read_excel(p, sheet_name=shn, header=None, dtype=str, nrows=15)
            hdr = next((i for i in range(len(probe))
                        if probe.iloc[i].astype("string").fillna("").str.strip()
                        .isin(["清單序號", "承批公司項目序號"]).any()), 0)
            t = pd.read_excel(p, sheet_name=shn, header=hdr, dtype=str)
            t.columns = [str(c).strip() for c in t.columns]
            t = t.loc[:, ~t.columns.duplicated()]
            L.append(f"\n{'='*76}\n## {ent} — {p.name[:50]} [{shn}] header={hdr} rows={len(t):,} cols={len(t.columns)}")
            L.append("   ALL columns:")
            for i in range(0, len(t.columns), 6):
                L.append("      " + " | ".join(str(c)[:24] for c in t.columns[i:i + 6]))
            pc, cc, nc = _s(t, "承批公司項目序號"), _s(t, "清單序號"), _s(t, "項目名稱")
            samp = [(a, b, str(c)[:18]) for a, b, c in zip(cc, pc, nc) if a or b][:8]
            L.append(f"   sample (清單序號, 承批序號, 名稱): {samp}")
            # numeric Σ census + golden proximity flags
            gold = {k: v * 1e4 for k, v in (GOLDEN_WAN.get(ent) or {}).items()}
            L.append("   numeric Σ census (全部 |Σ|>1萬 嘅欄):")
            for c in t.columns:
                v = _num(t[c].astype("string"))
                n = int(v.notna().sum())
                s = float(v.sum() or 0)
                if n < 3 or abs(s) < 1e4:
                    continue
                star = ""
                for b, g in gold.items():
                    if g and abs(s - g) <= max(abs(g) * 0.02, 2e4):
                        star += f"  ★≈golden {b}"
                # 清單係萬定 MOP? 兩個 scale 都對一次
                for b, g in gold.items():
                    if g and abs(s * 1e4 - g) <= max(abs(g) * 0.02, 2e4):
                        star += f"  ★(萬)≈golden {b}"
                L.append(f"      {str(c)[:44]:44s} Σ={s:>16,.0f}  n={n:>4}{star}")
        except Exception as e:
            L.append(f"   !! {p.name}: {type(e).__name__}: {e}")
    _w(L)


def _w(L):
    out = ROOT / "results" / "inspect_project_list.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste back")


if __name__ == "__main__":
    main()
