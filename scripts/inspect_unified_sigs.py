"""Per-entity: how well do the NEW unified-raw rows match the EXISTING unique_signatures
(tags we already paid for)? Decision input for the per-entity sig strategy — run BEFORE
wiring conf, so we know per entity whether old tags carry over or a migration is needed.

For each entity it reports, per unified file and combined:
  • old sig inventory: count, fields-per-sig (3 or 4 — job_code entities), tagged count
  • NEW sigs built the same way (account_code|account_desc|normalize_description(description))
  • overlap coverage of new ROWS and |amount|:
      - full-sig exact          (tags reuse as-is)
      - full-sig w/ acct dash→space  (vml: old acct key '80010 005' vs unified '80010-005')
      - account-level (code+desc) where old tagged sigs are UNANIMOUS on H
        (reusable via account-level injection even when desc_norm differs)
  • top uncovered new sigs by |amount| (what would actually need new tagging)

Run (Windows):  python scripts/inspect_unified_sigs.py            # all entities found in data\raw
                python scripts/inspect_unified_sigs.py vml galaxy # subset
Output: prints + results/inspect_unified_sigs.txt  (small — paste back)
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd
import yaml

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))
from kpi.lib.text import normalize_description  # noqa: E402  (same normalizer as step1/step4)

RAW_DIR = ROOT / "data" / "raw"
ENTITIES = {"galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
            "vml": "company_4", "melco": "company_5", "mgm": "company_6"}
NEED = ["account_code", "account_desc", "description", "amount_mop", "adjusted_amount"]


def _amt(df):
    a = pd.to_numeric(df.get("adjusted_amount", pd.Series(dtype=object)).astype("string")
                      .str.replace(",", "", regex=False), errors="coerce")
    m = pd.to_numeric(df.get("amount_mop", pd.Series(dtype=object)).astype("string")
                      .str.replace(",", "", regex=False), errors="coerce")
    return a.fillna(m).fillna(0.0)


def _s(df, col):
    return df.get(col, pd.Series("", index=df.index)).astype("string").fillna("").str.strip()


def main():
    want = [a.lower() for a in sys.argv[1:]] or list(ENTITIES)
    L = ["# inspect_unified_sigs — old unique_signatures vs NEW unified raw (per entity)"]
    for alias, com in ENTITIES.items():
        if alias not in want:
            continue
        files = sorted(RAW_DIR.glob(f"{alias}_*.xls*")) + sorted(RAW_DIR.glob(f"{alias}_*.csv"))
        files = [f for f in files if not f.name.startswith("~$")]
        L.append(f"\n{'='*76}\n## {alias} ({com}) — {len(files)} unified file(s)")
        if not files:
            L.append("   (no unified files yet — skip)"); continue

        sigp = ROOT / "data" / alias / "interim" / f"{com}_unique_signatures.xlsx"
        if not sigp.exists():
            L.append(f"   !! old unique_signatures missing: {sigp}"); continue
        old = pd.read_excel(sigp, dtype=str)
        old.columns = [str(c).strip() for c in old.columns]
        sig_col = "signature" if "signature" in old.columns else old.columns[0]
        old_sigs = old[sig_col].astype("string").fillna("")
        n_parts = int(old_sigs.str.count(r"\|").mode().iloc[0]) + 1 if len(old) else 3
        man = _s(old, "manual_horizontal"); llm = _s(old, "llm_horizontal")
        cur_h = man.where(man.ne(""), llm)
        tagged = cur_h.ne("")
        L.append(f"   old sigs: {len(old):,}  fields/sig={n_parts}  tagged={int(tagged.sum()):,}")
        old_keys = set(old_sigs)
        old_tagged_keys = set(old_sigs[tagged])
        # account-level (code|desc) unanimous-H map from tagged sigs
        ak = (_s(old, "account_code") + "|" + _s(old, "account_desc"))
        gv = pd.DataFrame({"k": ak[tagged], "h": cur_h[tagged]}).groupby("k")["h"].nunique()
        acct_unanimous = set(gv[gv == 1].index)

        tot_rows = tot_amt = 0
        cov = {"full": [0, 0.0], "dash": [0, 0.0], "acct": [0, 0.0], "tagged": [0, 0.0]}
        uncovered = []
        for f in files:
            try:
                hdr = pd.read_excel(f, nrows=0) if f.suffix != ".csv" else pd.read_csv(f, nrows=0)
                use = [c for c in NEED if c in [str(x).strip() for x in hdr.columns]]
                df = (pd.read_excel(f, dtype=str, usecols=use) if f.suffix != ".csv"
                      else pd.read_csv(f, dtype=str, usecols=use))
            except Exception as e:
                L.append(f"   !! {f.name}: read error {e}"); continue
            df.columns = [str(c).strip() for c in df.columns]
            ac, ad = _s(df, "account_code"), _s(df, "account_desc")
            dn = _s(df, "description").map(normalize_description)
            sig = ac + "|" + ad + "|" + dn
            sig_dash = ac.str.replace("-", " ", regex=False) + "|" + ad + "|" + dn
            akey = ac + "|" + ad
            akey_dash = ac.str.replace("-", " ", regex=False) + "|" + ad
            amt = _amt(df).abs()
            n, a = len(df), float(amt.sum())
            tot_rows += n; tot_amt += a
            m_full = sig.isin(old_keys)
            m_dash = m_full | sig_dash.isin(old_keys)
            m_tag = sig.isin(old_tagged_keys) | sig_dash.isin(old_tagged_keys)
            m_acct = m_dash | akey.isin(acct_unanimous) | akey_dash.isin(acct_unanimous)
            for key, m in (("full", m_full), ("dash", m_dash), ("tagged", m_tag), ("acct", m_acct)):
                cov[key][0] += int(m.sum()); cov[key][1] += float(amt[m].sum())
            L.append(f"   {f.name:22s} rows={n:>7,}  new-sigs={sig[~m_dash].nunique():>6,}  "
                     f"full={m_full.mean()*100:5.1f}%r  +dashfix={m_dash.mean()*100:5.1f}%r  "
                     f"+acctlvl={m_acct.mean()*100:5.1f}%r  amt-cov(acct)={float(amt[m_acct].sum())/a*100 if a else 0:5.1f}%")
            top = (pd.DataFrame({"sig": sig[~m_acct], "amt": amt[~m_acct]})
                   .groupby("sig")["amt"].agg(["sum", "size"]).sort_values("sum", ascending=False).head(8))
            for s_, r_ in top.iterrows():
                L.append(f"        UNCOV {r_['sum']/1e6:8.2f}M ({int(r_['size'])}r)  {str(s_)[:90]}")
        if tot_rows:
            L.append(f"   == {alias} TOTAL rows={tot_rows:,}  Σ|amt|={tot_amt/1e6:,.1f}M")
            for key, lab in (("full", "exact-sig"), ("dash", "+dash-fix"),
                             ("tagged", "already-TAGGED sig"), ("acct", "+acct-level inject")):
                L.append(f"      {lab:20s} rows {cov[key][0]/tot_rows*100:5.1f}%   |amt| {cov[key][1]/tot_amt*100 if tot_amt else 0:5.1f}%")
    out = ROOT / "results" / "inspect_unified_sigs.txt"; out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {out.relative_to(ROOT)}  ← paste back")


if __name__ == "__main__":
    main()
