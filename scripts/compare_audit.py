"""Compare 項目組 manual labels (原表類別 / 原表科目分類) vs our pipeline result
(大表類別匹配 / 大表科目匹配) in an audit xlsx tab — to see WHERE they differ, whether the
differences are PATTERNED, and whether the coarse NG × capex/opex picture agrees.

The two label sets use DIFFERENT vocabularies (項目組 category vs our V/H labels), so a plain
string-equality "match" is meaningless. Instead we build a CONFUSION view: for each 原表 value,
how our 大表 value distributes (by amount). A 原表 value that maps cleanly to ONE 大表 value =
consistent; one that scatters across many = the problem area.

Run on Windows (where the xlsx is):
  python scripts/compare_audit.py --file vml_audit_25_v2.xlsx --sheet "4_大表-0601"
Writes results/audit_compare.xlsx (full crosstabs) + prints a summary to paste back.
"""
from __future__ import annotations
import argparse, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent


def _find(cols, *cands):
    for c in cands:
        for col in cols:
            if str(col).strip() == c:
                return col
    # loose contains
    for c in cands:
        for col in cols:
            if c in str(col):
                return col
    return None


def _concentration(df, src, dst, amt, topn=3):
    """For each src value: total amt + its top-N dst values (by amt) + top1 share %."""
    g = df.groupby([src, dst])[amt].sum().abs().reset_index()
    out = []
    for sv, sub in g.groupby(src):
        tot = sub[amt].sum()
        if tot <= 0:
            continue
        sub = sub.sort_values(amt, ascending=False)
        top1 = sub.iloc[0]
        tops = " ; ".join(f"{r[dst]}={r[amt]/tot*100:.0f}%" for _, r in sub.head(topn).iterrows())
        out.append({src: sv, "Σamt": round(tot), "top1%": round(top1[amt] / tot * 100, 1),
                    "ndst": sub[dst].nunique(), f"top{topn}_大表": tops})
    return pd.DataFrame(out).sort_values("Σamt", ascending=False, key=lambda s: s.abs())


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--file", default="vml_audit_25_v2.xlsx")
    ap.add_argument("--sheet", default="4_大表-0601")
    ap.add_argument("--n", type=int, default=25)
    args = ap.parse_args()
    fp = ROOT / args.file
    if not fp.exists():
        fp = Path(args.file)
    if not fp.exists():
        print(f"X {args.file} not found (cwd or repo root)"); return
    df = pd.read_excel(fp, sheet_name=args.sheet)
    cols = list(df.columns)
    print(f"file={fp.name} sheet={args.sheet!r} rows={len(df):,}\ncolumns: {cols}")

    their_v = _find(cols, "原表類別")
    their_h = _find(cols, "原表科目分類", "原表科目")
    our_v = _find(cols, "大表類別匹配", "大表類別", "vertical_label")
    our_h = _find(cols, "大表科目匹配", "大表科目", "horizontal_label")
    amt = _find(cols, "amount_mop", "MOP Amt", "金額", "Amount", "amt")
    ng = _find(cols, "ng_code", "ng_label", "NG")
    co = _find(cols, "final_capex_opex", "CAPEX_OPEX", "capex_opex", "Capex/Opex")
    print(f"detected → 原表V={their_v!r} 原表H={their_h!r} 大表V={our_v!r} 大表H={our_h!r} "
          f"amt={amt!r} ng={ng!r} capex/opex={co!r}")
    if not amt:
        df["_amt"] = 1.0; amt = "_amt"; print("  (no amount col — counting rows)")
    df[amt] = pd.to_numeric(df[amt], errors="coerce").fillna(0)
    for c in (their_v, their_h, our_v, our_h, ng, co):
        if c:
            df[c] = df[c].astype(str).str.strip()

    out = ROOT / "results" / "audit_compare.xlsx"
    out.parent.mkdir(parents=True, exist_ok=True)
    blocks = []   # (section_title, dataframe) — all written to ONE sheet, stacked

    def _block(title, src, dst):
        if not src or not dst:
            print(f"\n[{title}] skip — missing column"); return
        conc = _concentration(df, src, dst, amt)
        scattered = conc[conc["top1%"] < 80]
        print(f"\n=== {title}: 原表 → 大表 (Σamt) ===")
        print(f"  {conc[src].nunique()} 原表值;{(conc['top1%']>=80).sum()} 個乾淨對應(top1≥80%),"
              f"{len(scattered)} 個分散(<80%)")
        print(f"  最大 {args.n} 個 原表值(scattered 標 ⚠):")
        for _, r in conc.head(args.n).iterrows():
            flag = "  ⚠分散" if r["top1%"] < 80 else ""
            print(f"   {str(r[src])[:20]:20} Σ={r['Σamt']:>14,.0f}  top1={r['top1%']:>5}% ndst={r['ndst']:>2}  "
                  f"{r['top'+'3_大表'][:70]}{flag}")
        blocks.append((f"{title} — concentration (原表→大表)", conc.rename(columns={src: "原表"})))
        ct = df.groupby([src, dst])[amt].sum().reset_index().sort_values(amt, ascending=False,
                                                                          key=lambda s: s.abs())
        blocks.append((f"{title} — pairs", ct.rename(columns={src: "原表", dst: "大表"})))

    _block("類別V", their_v, our_v)
    _block("科目H", their_h, our_h)

    if ng and co:
        piv = df.groupby([ng, co])[amt].sum().reset_index().rename(columns={amt: "Σamt"}).sort_values([ng, co])
        print(f"\n=== NG × capex/opex (我們 ng/co) Σamt ===")
        for _, r in piv.iterrows():
            print(f"   {str(r[ng]):8} {str(r[co]):10} {r['Σamt']:>16,.0f}")
        blocks.append(("NG × capex/opex (Σamt)", piv))
        if their_v and our_v:
            rowsum = [{ng: ngv, co: cov, "Σamt": round(sub.groupby(their_v)[amt].sum().abs().sum()),
                       "原表V數": int(sub.groupby(their_v)[amt].sum().abs().gt(0).sum()),
                       "原表V_top": (sub.groupby(their_v)[amt].sum().abs().idxmax() if len(sub) else "")}
                      for (ngv, cov), sub in df.groupby([ng, co])]
            blocks.append(("NG×capex/opex — 原表V 集中度", pd.DataFrame(rowsum)))

    # write ALL blocks to ONE sheet, stacked with a title row + blank separator
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        row = 0
        for title, d in blocks:
            pd.DataFrame([[title]]).to_excel(writer, sheet_name="compare", startrow=row,
                                             header=False, index=False)
            d.to_excel(writer, sheet_name="compare", startrow=row + 1, index=False)
            row += len(d) + 3
    print(f"\n→ {out}  (單一 sheet 'compare',全部 stacked)  — paste 上面 console 畀我分析")


if __name__ == "__main__":
    main()
