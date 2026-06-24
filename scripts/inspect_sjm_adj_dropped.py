"""inspect_sjm_adj_dropped.py — SJM 25 調整數點解 0：net-zero 撤回行被 step0.5 drop
Run: python scripts\inspect_sjm_adj_dropped.py
Out: results\inspect_sjm_adj_dropped.txt
比較每個 25-bucket：kept(net≠0) vs dropped(net=0但adj≠0) 的調整金額總額。
"""
from __future__ import annotations
import sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
F    = ROOT / "data" / "sjm" / "raw" / "sjm_2025.xlsx"
OUT  = ROOT / "results" / "inspect_sjm_adj_dropped.txt"

# bucket → (amount_col, adjust_col)  —— 同 conf year_split 2025 一致
BUCKETS = [("25", "25跨年", "25調整金額"),
           ("25_24SY", "24跨年", "24調整金額"),
           ("25_23SY", "23跨年", "23調整金額")]


def _num(s): return pd.to_numeric(s.astype(str).str.replace(",", "", regex=False), errors="coerce").fillna(0.0)


def main():
    L = ["# inspect_sjm_adj_dropped — SJM 25 調整數 0 根因", ""]
    if not F.exists():
        L.append(f"!! {F} 揾唔到"); OUT.write_text("\n".join(L), encoding="utf-8"); print("\n".join(L)); return

    df = pd.read_excel(F, sheet_name="data", dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    L.append(f"rows={len(df):,}")

    for bk, amt_c, adj_c in BUCKETS:
        L += ["", "=" * 60, f"## bucket {bk}  (amount={amt_c}  adjust={adj_c})"]
        if amt_c not in df.columns or adj_c not in df.columns:
            L.append(f"  !! 欄缺: amt={amt_c in df.columns} adj={adj_c in df.columns}"); continue
        amt = _num(df[amt_c]); adj = _num(df[adj_c])
        post = amt + adj
        kept = post != 0.0
        dropped_adj = (post == 0.0) & (adj != 0.0)   # net-zero 但有調整 → 而家被 drop
        L.append(f"  kept (post≠0):          {int(kept.sum()):>6} 行   調整Σ={adj[kept].sum()/1e4:>12,.0f}萬   post(amount_mop)Σ={post[kept].sum()/1e4:>12,.0f}萬")
        L.append(f"  DROPPED (post=0,adj≠0): {int(dropped_adj.sum()):>6} 行   調整Σ={adj[dropped_adj].sum()/1e4:>12,.0f}萬   調整前Σ={amt[dropped_adj].sum()/1e4:>12,.0f}萬")
        L.append(f"  → 全 bucket 調整Σ(kept+dropped) = {adj[kept | dropped_adj].sum()/1e4:,.0f}萬")

    OUT.parent.mkdir(exist_ok=True)
    OUT.write_text("\n".join(L), encoding="utf-8")
    print("\n".join(L)); print(f"\nwrote {OUT.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
