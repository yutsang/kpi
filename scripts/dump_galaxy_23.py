"""Dump Galaxy 2023 row-level signatures (H) + projects (V) for baseline LLM classify.

Per user: don't hard-code H from each source's "system meaning" — dump every distinct
account|desc signature across the 9 取數 sources so I classify row-by-row, THEN inspect &
correct. NG + V(項目類別) are given per DICJ (never overridden), carried only to compare.

Sources (sap = galaxy_sap.xlsx; book = galaxy_raw_databook.xlsx, password via --pw):
  SAP opex/capex (dual-block non-gaming) + SAP Gaming capex + PMS + VMS + VR + Simulation
  + Pre-Opening + Corporate support + Other Source Detail.

Run (Windows):
  python scripts/dump_galaxy_23.py --pw dicj_kpmg
Output → results/galaxy_sigs_23.tsv (H baseline) + results/galaxy_projects_23.tsv (V compare)
"""
from __future__ import annotations
import argparse, io, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent


def find_file(name):
    p = Path(name)
    if p.exists(): return p
    d = ROOT / "data"
    hits = list(d.glob(f"*/raw/{name}")) + list(d.rglob(name)) + list(d.rglob(f"*{name}*"))
    exact = [h for h in hits if h.name == name]
    return (exact or hits)[0] if (exact or hits) else None


def handle(fp, pw):
    if not pw: return fp
    import msoffcrypto
    buf = io.BytesIO()
    with open(fp, "rb") as f:
        of = msoffcrypto.OfficeFile(f); of.load_key(password=pw); of.decrypt(buf)
    buf.seek(0); return buf


def rd(fp, sheet, header, pw):
    return pd.read_excel(handle(fp, pw), sheet_name=sheet, header=header, dtype=object)


def col(df, name):
    if name is None: return None
    if name in df.columns: return name
    for c in df.columns:
        if str(c).strip() == str(name).strip(): return c
    return None


def num(s): return pd.to_numeric(s, errors="coerce")


# each H-source → how to read its signature. (file_key, sheet, header, ac, ad, amt, capex, ng, cat, dicj, proj, filter)
def H_SOURCES(pw_set):
    book = "book" if pw_set else None  # gaming capex only if pw provided
    S = [
        dict(src="SAP_opex",     fk="sap", sheet="SAP Record-non gaming", hdr=0,
             ac="Account Code", ad="Account Description", amt="Reported Amount (MOP)", capex="Opex"),
        dict(src="SAP_capex_ng", fk="sap", sheet="SAP Record-non gaming", hdr=0,
             ac="Account Code.1", ad="Account Description.1", amt="Reported Amount (MOP).1", capex="Capx"),
        dict(src="PMS",          fk="sap", sheet="PMS", hdr=0,
             ac="Account Code", ad="Account Description", amt="Amount (MOP)", capex="Comp"),
        dict(src="VMS",          fk="sap", sheet="VMS", hdr=3,
             ac=None, ad="Descriptions", amt="Amount (Reported Amt) (MOP)", capex="Comp"),
        dict(src="VR",           fk="sap", sheet="VR Record", hdr=3,
             ac=None, ad="Description.1", amt="Amount", capex="Staff"),
        dict(src="Simulation",   fk="sap", sheet="Simulation&Pre-Opening", hdr=2,
             ac="Account Code", ad="Account Description", amt="Reported Amount", capex="Sim",
             filt=("Nature", "Simulation")),
        dict(src="PreOpening",   fk="sap", sheet="Simulation&Pre-Opening", hdr=2,
             ac="Account Code", ad="Account Description", amt="Reported Amount", capex="PreOpen",
             filt=("Nature", "Pre-Opening")),
        dict(src="Other",        fk="sap", sheet="Other Source Detail", hdr=0,
             ac="GL Account", ad="GL Account Description", amt="Investment Amount", capex="Other"),
    ]
    if book:
        S.append(dict(src="SAP_capex_gam", fk="book", sheet="SAP Gaming", hdr=1,
                      ac="Account Code.1", ad="Account Description.1", amt="Reported Amount (MOP).1", capex="Capx"))
    return S

NG_COL, CAT_COL = "NG11 Category", "項目類別"


def dump_sigs(files, pw):
    rows = []
    for s in H_SOURCES(bool(pw)):
        fp = files[s["fk"]]
        if not fp: print(f"[{s['src']}] file missing"); continue
        try:
            df = rd(fp, s["sheet"], s["hdr"], pw if s["fk"] == "book" else None)
        except Exception as e:
            print(f"[{s['src']}] read failed: {e}"); continue
        amt = col(df, s["amt"])
        if not amt:
            print(f"[{s['src']}] amount col {s['amt']!r} missing — skip"); continue
        if s.get("filt"):
            fc = col(df, s["filt"][0])
            if fc: df = df[df[fc].astype(str).str.strip().str.lower() == s["filt"][1].lower()]
        ac, ad = col(df, s["ac"]), col(df, s["ad"])
        ng, cat = col(df, NG_COL), col(df, CAT_COL)
        sub = pd.DataFrame({
            "source": s["src"], "capex_opex": s["capex"],
            "account_code": df[ac].astype(str).str.strip() if ac else "",
            "account_desc": df[ad].astype(str).str.strip() if ad else "",
            "_amt": num(df[amt]).fillna(0.0),
            "_ng": df[ng].astype(str).str.strip() if ng else "",
            "_cat": df[cat].astype(str).str.strip() if cat else "",
        })
        rows.append(sub)
        print(f"[{s['src']}] {len(sub):,} rows  Σ={sub['_amt'].sum():,.0f}")
    if not rows: return
    allr = pd.concat(rows, ignore_index=True)
    g = (allr.groupby(["source", "capex_opex", "account_code", "account_desc"])
              .agg(amount=("_amt", "sum"), n=("_amt", "size"),
                   NG=("_ng", lambda x: x.mode().iloc[0] if not x.mode().empty else ""),
                   項目類別=("_cat", lambda x: " / ".join(x.value_counts().head(2).index)))
              .reset_index().sort_values("amount", ascending=False, key=abs))
    out = ROOT / "results" / "galaxy_sigs_23.tsv"
    out.parent.mkdir(exist_ok=True)
    g.to_csv(out, sep="\t", index=False, encoding="utf-8-sig")
    print(f"\n✓ {len(g):,} sigs → results/galaxy_sigs_23.tsv  (Σ all={allr['_amt'].sum():,.0f})")


def dump_projects(files, pw):
    fp = files["sap"]
    df = rd(fp, "Cover", 2, None)
    dicj, proj, cat, ng = col(df, "DICJ Code"), col(df, "Project"), col(df, CAT_COL), col(df, NG_COL)
    amt = col(df, "Cover Amount")
    g = (df.assign(_d=df[dicj].astype(str).str.strip(),
                   _p=df[proj].astype(str).str.replace(r"[\r\n]+", " ", regex=True).str.strip(),
                   _c=df[cat].astype(str).str.strip(), _n=df[ng].astype(str).str.strip(),
                   _a=num(df[amt]).fillna(0.0))
           .groupby("_d").agg(project=("_p", lambda x: x.mode().iloc[0] if not x.mode().empty else ""),
                              項目類別=("_c", lambda x: x.mode().iloc[0] if not x.mode().empty else ""),
                              NG=("_n", lambda x: x.mode().iloc[0] if not x.mode().empty else ""),
                              amount=("_a", "sum"), n=("_a", "size"))
           .reset_index().rename(columns={"_d": "DICJ"}).sort_values("amount", ascending=False, key=abs))
    out = ROOT / "results" / "galaxy_projects_23.tsv"
    out.parent.mkdir(exist_ok=True)
    g.to_csv(out, sep="\t", index=False, encoding="utf-8-sig")
    print(f"✓ {len(g):,} projects → results/galaxy_projects_23.tsv  (Σ Cover={g['amount'].sum():,.0f})")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--sap", default="galaxy_sap.xlsx")
    ap.add_argument("--book", default="galaxy_raw_databook.xlsx")
    ap.add_argument("--pw", default=None, help="password for the encrypted databook (runtime only)")
    a = ap.parse_args()
    files = {"sap": find_file(a.sap), "book": find_file(a.book)}
    print(f"sap={files['sap']}  book={files['book']}  pw={'set' if a.pw else 'none'}")
    dump_projects(files, a.pw)
    dump_sigs(files, a.pw)


if __name__ == "__main__":
    main()
