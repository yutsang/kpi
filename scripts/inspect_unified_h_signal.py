"""H-signal sufficiency check on the unified raw — USING ONLY the 3 sig fields
(account_code / account_desc / description). pt_class_H/V & comp_type are reference cols
and deliberately EXCLUDED. Question answered: 每行有冇足夠資訊分 H？

Per file, every row is graded:
  S1  account_desc non-blank                      → strong  (GL name readable: rules/LLM OK)
  S2  account_desc blank, account_code non-blank  → medium  (code-only: prefix rules possible)
  S3  acct both blank, description informative    → desc-only (LLM on text)
  S4  acct both blank + description blank/junk    → ★NO SIGNAL — cannot classify H from data
Junk description = after normalize_description (same as pipeline) and stripping <num>/<date>/
<netoff>/<id> placeholders, fewer than 3 letter/CJK chars remain (e.g. pure numbers, dates,
netoff markers).
Also flags S1 rows whose description is junk (acct-name is then the ONLY signal — fine, but
explains why desc-level sigs explode; candidates for account-level signature_fields).

Run (Windows):  python scripts/inspect_unified_h_signal.py          # all files in data\raw
                python scripts/inspect_unified_h_signal.py wynn     # subset
Output: prints + results/inspect_unified_h_signal.txt (small — paste back)
"""
from __future__ import annotations
import re
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))
from kpi.lib.text import normalize_description  # noqa: E402

RAW_DIR = ROOT / "data" / "raw"
NEED = ["account_code", "account_desc", "description", "amount_mop", "adjusted_amount"]
_PLACEHOLDER = re.compile(r"<[^>]+>")
_INFO_CHARS = re.compile(r"[A-Za-z一-鿿]")


def _amt(df):
    a = pd.to_numeric(df.get("adjusted_amount", pd.Series(dtype=object)).astype("string")
                      .str.replace(",", "", regex=False), errors="coerce")
    m = pd.to_numeric(df.get("amount_mop", pd.Series(dtype=object)).astype("string")
                      .str.replace(",", "", regex=False), errors="coerce")
    return a.fillna(m).fillna(0.0)


def _s(df, col):
    return df.get(col, pd.Series("", index=df.index)).astype("string").fillna("").str.strip()


def _informative(dn: str) -> bool:
    t = _PLACEHOLDER.sub("", dn)
    return len(_INFO_CHARS.findall(t)) >= 3


def main():
    want = [a.lower() for a in sys.argv[1:]]
    files = sorted(p for p in RAW_DIR.iterdir()
                   if p.suffix.lower() in (".xlsx", ".xlsm", ".xls", ".csv")
                   and not p.name.startswith("~$"))
    if want:
        files = [f for f in files if any(f.name.lower().startswith(w) for w in want)]
    L = ["# inspect_unified_h_signal — acct_code/acct_desc/description ONLY (pt_class excluded)",
         "# S1=acct_desc有 S2=淨code S3=淨desc S4=★乜都冇(分唔到H)"]
    for f in files:
        try:
            hdr = pd.read_excel(f, nrows=0) if f.suffix != ".csv" else pd.read_csv(f, nrows=0)
            cols = [str(c).strip() for c in hdr.columns]
            use = [c for c in NEED if c in cols]
            df = (pd.read_excel(f, dtype=str, usecols=use) if f.suffix != ".csv"
                  else pd.read_csv(f, dtype=str, usecols=use))
        except Exception as e:
            L.append(f"\n## {f.name}: read error {e}"); continue
        df.columns = [str(c).strip() for c in df.columns]
        ac, ad, de = _s(df, "account_code"), _s(df, "account_desc"), _s(df, "description")
        dn = de.map(normalize_description)
        informative = dn.map(_informative).astype(bool)   # .map on string dtype keeps string dtype → mean() blows up
        amt = _amt(df).abs()
        tot_a = float(amt.sum()) or 1.0

        s1 = ad.ne("")
        s2 = ~s1 & ac.ne("")
        s3 = ~s1 & ~s2 & informative
        s4 = ~s1 & ~s2 & ~s3
        s1_junkdesc = s1 & ~informative

        L.append(f"\n{'='*72}\n## {f.name}  rows={len(df):,}  Σ|amt|={tot_a/1e6:,.1f}M")
        L.append(f"   fill%: acct_code={ac.ne('').mean()*100:.0f}%  acct_desc={ad.ne('').mean()*100:.0f}%  "
                 f"description={de.ne('').mean()*100:.0f}%  (desc informative={informative.mean()*100:.0f}%)")
        for name, m, note in (("S1 acct_desc有(強)", s1, ""),
                              ("   └ 其中desc係junk", s1_junkdesc, "  ← acct名係唯一信號(account級sig候選)"),
                              ("S2 淨code(中)", s2, ""),
                              ("S3 淨desc(LLM文本)", s3, ""),
                              ("S4 ★冇信號", s4, "  ← 分唔到H,要靠參考欄/問項目組")):
            L.append(f"   {name:22s} {int(m.sum()):>8,}r ({m.mean()*100:5.1f}%)  "
                     f"|amt| {float(amt[m].sum())/1e6:>10.2f}M ({float(amt[m].sum())/tot_a*100:5.1f}%){note}")
        if int(s4.sum()):
            grp = (pd.DataFrame({"k": (ac + "|" + ad + "|" + dn)[s4], "amt": amt[s4]})
                   .groupby("k")["amt"].agg(["sum", "size"]).sort_values("sum", ascending=False).head(6))
            for k, r in grp.iterrows():
                L.append(f"        S4例 {r['sum']/1e6:8.2f}M ({int(r['size'])}r)  '{str(k)[:80]}'")
        if int(s2.sum()):
            grp = (pd.DataFrame({"k": ac[s2], "amt": amt[s2]})
                   .groupby("k")["amt"].agg(["sum", "size"]).sort_values("sum", ascending=False).head(4))
            for k, r in grp.iterrows():
                L.append(f"        S2例 {r['sum']/1e6:8.2f}M ({int(r['size'])}r)  code='{str(k)[:50]}'")
    out = ROOT / "results" / "inspect_unified_h_signal.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste back")


if __name__ == "__main__":
    main()
