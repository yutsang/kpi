"""Inject OUR horizontal (H) classification into the step3 feedback override file.

The H-axis twin of inject_manual_vertical. Reads a TSV of
    signature <TAB> H_CODE          (our per-unique-signature H, classified from
                                     account_code / account_desc / desc / projects —
                                     NOT copied from the project team's 分類1 column)
from data/_overrides/{ent}_horizontal.tsv (or --tsv PATH) and writes/merges it into
    data/{ent}/output/{com}_feedback.xlsx     (signature + correct_horizontal)

step3 loads that file via load_feedback() and treats filled rows as DIRECT OVERRIDES
("User-verified feedback override"), applied BEFORE the LLM and BEFORE the
skip_signature_llm default — so OUR H wins. The project team's 分類1 column_map is then
used only for reconciliation/comparison, not as the source (demote it in conf so it
doesn't overwrite our H at step4). Existing feedback rows are preserved.

Run (Windows), after writing the TSV:
    python scripts/inject_horizontal_feedback.py --entity vml
    python scripts/inject_horizontal_feedback.py --entity vml --tsv results/vml_horizontal.tsv
Then: kedro run --pipeline=vml
"""
from __future__ import annotations
import argparse, sys
from pathlib import Path
try: sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
except Exception: pass
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
ENTITIES = {"galaxy": "company_1", "sjm": "company_2", "wynn": "company_3",
            "vml": "company_4", "melco": "company_5", "mgm": "company_6"}
OVERRIDE_COL = "correct_horizontal"


def _read_pairs(src: Path) -> dict[str, str]:
    """signature<TAB>H_CODE  (tolerates header / blank / # comment / reversed order)."""
    pairs: dict[str, str] = {}
    for raw in src.read_text(encoding="utf-8").splitlines():
        line = raw.rstrip("\n")
        if not line.strip() or line.lstrip().startswith("#"):
            continue
        parts = [p.strip() for p in line.split("\t")]
        if len(parts) < 2:
            continue
        a, b = parts[0], parts[1]
        # accept either "signature<TAB>H_CODE" or "H_CODE<TAB>signature"
        if a.upper().startswith("H_") and not b.upper().startswith("H_"):
            sig, h = b, a
        else:
            sig, h = a, b
        if not sig or not h or sig.lower() == "signature" or not h.upper().startswith("H_"):
            continue
        pairs[sig] = h.upper()
    return pairs


def inject(ent: str, tsv: Path | None):
    com = ENTITIES[ent]
    src = tsv or (ROOT / "data" / "_overrides" / f"{ent}_horizontal.tsv")
    if not src.exists():
        print(f"[{ent}] X missing {src} — write signature<TAB>H_CODE rows first.")
        return
    pairs = _read_pairs(src)
    if not pairs:
        print(f"[{ent}] X no valid signature<TAB>H_CODE rows in {src.name}.")
        return

    out_dir = ROOT / "data" / ent / "output"
    out_dir.mkdir(parents=True, exist_ok=True)
    fb = out_dir / f"{com}_feedback.xlsx"
    if fb.exists():
        df = pd.read_excel(fb)
    else:
        df = pd.DataFrame(columns=["signature", OVERRIDE_COL, "notes"])
    for c in ("signature", OVERRIDE_COL):
        if c not in df.columns:
            df[c] = ""
    df["signature"] = df["signature"].astype(str)

    existing = set(df["signature"])
    updated = 0
    for i, sig in df["signature"].items():
        if sig in pairs:
            df.at[i, OVERRIDE_COL] = pairs[sig]
            updated += 1
    new_sigs = [s for s in pairs if s not in existing]
    if new_sigs:
        add = pd.DataFrame({"signature": new_sigs,
                            OVERRIDE_COL: [pairs[s] for s in new_sigs]})
        df = pd.concat([df, add], ignore_index=True)

    df.to_excel(fb, index=False)
    print(f"[{ent}] {fb.relative_to(ROOT)}: {updated} updated + {len(new_sigs)} new "
          f"= {len(pairs)} H overrides ({len(df)} rows total). "
          f"step3 applies OUR H via feedback (skip LLM); demote 分類1 column_map in conf.")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--entity", nargs="+", required=True, choices=sorted(ENTITIES))
    ap.add_argument("--tsv", type=Path, default=None, help="override TSV (default data/_overrides/{ent}_horizontal.tsv)")
    a = ap.parse_args()
    for e in a.entity:
        inject(e, a.tsv)


if __name__ == "__main__":
    main()
