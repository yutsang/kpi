r"""Scan data/<ent>/raw/project_team/*.xlsx for a per-subproject DICJ column (year-aware, no ambiguity).

Why: name-substr 反填對 generic「其他XXX」桶有年份歧義 (同一中文名 → 唔同年唔同 DICJ)。項目組自己版本好
可能逐行已有正確 DICJ code (已分年)。呢個 script 自動揾：
  - 邊個 col 嘅值同 golden DICJ set 高度 overlap  → DICJ-code col
  - 邊個 col 似中文項目名 (CJK 比例高)            → name col
  - 邊個 col 似年份 (2023/2024/2025 / 23/24/25)   → year col
然後印 col 清單 + sample 行，等 Claude 決定用邊幾個 col 砌 (subproject/名,年 → DICJ)。
amt 只係 check 數,唔做 mapping 條件。

Run (Windows):  python scripts/inspect_ptteam.py            # all 6
                python scripts/inspect_ptteam.py wynn sjm   # 指定
"""
from __future__ import annotations
import sys, re, unicodedata
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
GCAND = [ROOT / "HQ 投資方向_audit_0616.xlsx", ROOT / "data" / "HQ 投資方向_audit_0616.xlsx"]
TAB = "Database combine"
ENT = {"galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
       "vml": "company_4", "melco": "company_5", "mgm": "company_6"}
NAME2ALIAS = [(["galaxy", "銀河"], "galaxy"), (["wynn", "永利"], "wynn"), (["mgm", "美高梅"], "mgm"),
              (["melco", "新濠", "摩珀斯", "影匯", "影滙", "studio city", "city of dreams"], "melco"),
              (["sjm", "澳娛", "葡京", "回力", "上葡京"], "sjm"),
              (["威尼斯", "金沙", "sands", "londoner", "倫敦人", "parisian", "巴黎人", "venetian", "vml"], "vml")]


def _alias(s):
    sl = str(s).strip().lower()
    for kws, a in NAME2ALIAS:
        if any(k.lower() in sl for k in kws):
            return a
    return None


def _ndicj(v):
    v = str(v).strip()
    m = re.match(r"^(項目)0*(\d.*)$", v)
    return m.group(1) + m.group(2) if m else v


def _cjk(s):
    s = str(s)
    if not s.strip():
        return 0.0
    return sum(1 for c in s if "一" <= c <= "鿿") / max(len(s), 1)


def _golden_dicj(g, alias):
    ga = g[g["_a"] == alias]
    return set(_ndicj(d).lower() for d in ga["_d"].unique() if str(d).strip() and str(d).strip() != "nan")


def main():
    want = [a.lower() for a in sys.argv[1:]] or list(ENT)
    L = ["# inspect_ptteam — 揾項目組版本逐行 DICJ col (年份正確,解歧義)"]
    gpath = next((p for p in GCAND if p.exists()), None)
    g = None
    if gpath:
        g = pd.read_excel(gpath, sheet_name=TAB, dtype=str)
        g.columns = [str(c).strip() for c in g.columns]
        g["_a"] = g[next(c for c in g.columns if "承批" in c)].map(_alias)
        g["_d"] = g[next(c for c in g.columns if c.strip() in ("DICJ Code", "DICJ"))].astype(str).str.strip()
    else:
        L.append("!! golden 揾唔到 → 只做欄位偵測,冇 DICJ-overlap")

    for alias in want:
        L.append(f"\n{'='*72}\n## {alias}")
        ptdir = ROOT / "data" / alias / "raw" / "project_team"
        if not ptdir.exists():
            L.append(f"   (冇 {ptdir.relative_to(ROOT)} — 跳過)")
            continue
        gset = _golden_dicj(g, alias) if g is not None else set()
        L.append(f"   golden DICJ set: {len(gset)} codes")
        files = sorted(list(ptdir.glob("*.xlsx")) + list(ptdir.glob("*.xls")))
        if not files:
            L.append("   (folder 入面冇 xlsx)")
            continue
        for f in files:
            try:
                xl = pd.ExcelFile(f)
            except Exception as e:
                L.append(f"   !! 讀唔到 {f.name}: {e}")
                continue
            for sh in xl.sheet_names:
                try:
                    df = xl.parse(sh, dtype=str)
                except Exception:
                    continue
                if df.empty:
                    continue
                df = df.dropna(how="all").dropna(axis=1, how="all")
                if df.empty:
                    continue
                L.append(f"\n   ── {f.name} :: [{sh}]  {df.shape[0]}r × {df.shape[1]}c")
                dicj_cols, name_cols, year_cols = [], [], []
                for c in df.columns:
                    vals = df[c].dropna().astype(str).str.strip()
                    vals = vals[vals != ""]
                    if len(vals) == 0:
                        continue
                    uniq = vals.unique()
                    samp = uniq[:3000]
                    # DICJ overlap
                    if gset:
                        hit = sum(1 for v in samp if _ndicj(v).lower() in gset)
                        ov = hit / max(len(samp), 1)
                        if ov >= 0.30:
                            dicj_cols.append((c, ov, len(uniq)))
                    # year
                    yv = sum(1 for v in samp if re.fullmatch(r"(20)?2[2-6](\.0)?", v.strip()))
                    if yv / max(len(samp), 1) >= 0.6:
                        year_cols.append(c)
                    # 中文名
                    cjk = sum(_cjk(v) for v in samp[:500]) / max(min(len(samp), 500), 1)
                    avglen = sum(len(v) for v in samp[:500]) / max(min(len(samp), 500), 1)
                    if cjk >= 0.35 and avglen >= 5 and len(uniq) >= 5:
                        name_cols.append((c, round(cjk, 2), round(avglen, 1)))
                L.append(f"      cols: {list(df.columns)}")
                if dicj_cols:
                    for c, ov, nu in sorted(dicj_cols, key=lambda x: -x[1]):
                        L.append(f"      ★ DICJ-col '{c}': {ov*100:.0f}% 值 match golden DICJ  ({nu} uniq)")
                else:
                    L.append("      (冇 col 同 golden DICJ overlap ≥30%)")
                if year_cols:
                    L.append(f"      year-col: {year_cols}")
                if name_cols:
                    L.append(f"      name-col: {[c for c, _, _ in name_cols]}")
                # sample rows of best DICJ + name + year cols
                bc = sorted(dicj_cols, key=lambda x: -x[1])[0][0] if dicj_cols else None
                nc = name_cols[0][0] if name_cols else None
                yc = year_cols[0] if year_cols else None
                show = [x for x in [bc, nc, yc] if x]
                if show:
                    L.append(f"      sample [{' | '.join(show)}]:")
                    for _, r in df[show].dropna(how="all").drop_duplicates().head(8).iterrows():
                        L.append("        " + "  |  ".join(str(r[x]).strip()[:40] for x in show))
    out = ROOT / "results" / "inspect_ptteam.txt"
    out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L))
    print(f"\nwrote {out.relative_to(ROOT)}  ← paste 返嚟")


if __name__ == "__main__":
    main()
