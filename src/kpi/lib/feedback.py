"""LLM tuning via human feedback.

After step 3, a small `<entity>_feedback.xlsx` is produced under data/<entity>/output/
containing the 30 most impactful signatures the LLM tagged (top by |amount|, plus
H_OTHER and low-confidence ones). It has empty `correct_horizontal` and `notes` columns
for the user to fill.

Next time step 3 runs:
  - Rows where the user filled `correct_horizontal` become DIRECT OVERRIDES (skip LLM).
  - The first few overrides are also injected into the LLM system prompt as few-shot
    examples so the LLM generalises from your judgement.
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd

# Columns the user fills
OVERRIDE_COL = "correct_horizontal"
NOTES_COL = "notes"

FEEDBACK_DISPLAY_COLS = [
    "signature",
    "account_code",
    "account_desc",
    "desc_norm",
    "desc_samples",
    "row_count",
    "total_amount",
    "has_negative",
    "project_samples",
    "llm_horizontal",
    "llm_horizontal_label",
    "llm_confidence",
    "llm_reasoning",
    OVERRIDE_COL,
    NOTES_COL,
]


def feedback_path(output_dir: Path, company: str) -> Path:
    return output_dir / f"{company}_feedback.xlsx"


def load_feedback(output_dir: Path, company: str) -> tuple[dict[str, str], list[dict]]:
    """Returns (overrides, few_shot_examples).
    overrides: signature → correct_horizontal_id (filled rows only).
    few_shot_examples: up to 8 dicts {account_desc, desc_norm, correct_horizontal} for prompt injection.
    """
    path = feedback_path(output_dir, company)
    if not path.exists():
        return {}, []
    try:
        df = pd.read_excel(path)
    except Exception as e:
        print(f"  [feedback] could not read {path.name}: {e}")
        return {}, []

    if OVERRIDE_COL not in df.columns or "signature" not in df.columns:
        return {}, []

    overrides: dict[str, str] = {}
    examples: list[dict] = []
    for _, r in df.iterrows():
        sig = str(r.get("signature", "")).strip()
        correct = str(r.get(OVERRIDE_COL, "")).strip()
        if not sig or not correct or correct.lower() == "nan":
            continue
        overrides[sig] = correct
        if len(examples) < 8:
            examples.append({
                "account_desc": str(r.get("account_desc", "")),
                "desc_norm": str(r.get("desc_norm", "")),
                "correct_horizontal": correct,
                "notes": str(r.get(NOTES_COL, "")) if pd.notna(r.get(NOTES_COL)) else "",
            })
    print(
        f"  [feedback] loaded {len(overrides)} overrides + "
        f"{len(examples)} few-shot examples from {path.name}"
    )
    return overrides, examples


def write_feedback_template(
    sig_df: pd.DataFrame,
    output_dir: Path,
    company: str,
    n_top_amount: int = 100,
    n_h_other: int = 60,
    n_low_conf: int = 40,
) -> Path:
    """Pick a focused review set for human spot-check and LLM tuning.

    Sections (de-duplicated before writing):
      - Top N signatures by |amount|          → highest-impact items
      - Top N H_OTHER signatures by |amount|  → diagnose prompt over-use of '其他'
      - Top N low-confidence (<0.7) by amount → most likely to be wrong

    Preserves any existing user-filled `correct_horizontal` / `notes` so re-running
    step 3 doesn't blow away your feedback work.
    """
    path = feedback_path(output_dir, company)
    output_dir.mkdir(parents=True, exist_ok=True)

    # Preserve existing user inputs
    prev_overrides: dict[str, dict] = {}
    if path.exists():
        try:
            prev = pd.read_excel(path)
            for _, r in prev.iterrows():
                sig = str(r.get("signature", "")).strip()
                if not sig:
                    continue
                prev_overrides[sig] = {
                    OVERRIDE_COL: str(r.get(OVERRIDE_COL, "")) if pd.notna(r.get(OVERRIDE_COL)) else "",
                    NOTES_COL: str(r.get(NOTES_COL, "")) if pd.notna(r.get(NOTES_COL)) else "",
                }
        except Exception:
            pass

    df = sig_df.copy()
    df["_abs"] = pd.to_numeric(df.get("total_amount"), errors="coerce").abs().fillna(0)

    parts = []
    # Section 1: top by amount (always — the most material items)
    parts.append(df.sort_values("_abs", ascending=False).head(n_top_amount).assign(_section="top_amount"))
    # Section 2: H_OTHER bloat — helps diagnose if prompt is under-confident
    if "llm_horizontal" in df.columns:
        h_other = df[df["llm_horizontal"].astype("string").fillna("") == "H_OTHER"]
        parts.append(h_other.sort_values("_abs", ascending=False).head(n_h_other).assign(_section="h_other"))
    # Section 3: low confidence — likely wrong tags, good for correcting
    if "llm_confidence" in df.columns:
        low_conf = df[pd.to_numeric(df["llm_confidence"], errors="coerce") < 0.7]
        parts.append(low_conf.sort_values("_abs", ascending=False).head(n_low_conf).assign(_section="low_conf"))

    sample = pd.concat(parts).drop_duplicates(subset=["signature"])

    # Preserve ALL rows that have a non-empty user override in the prev xlsx,
    # not just the ones that happen to fall in the top-N sample. This was a
    # silent data-loss bug: injected feedback for sigs outside the sample
    # used to be dropped at the next step3 run.
    prev_filled_sigs = {
        sig for sig, d in prev_overrides.items()
        if d.get(OVERRIDE_COL, "").strip() and d.get(OVERRIDE_COL, "").strip().lower() != "nan"
    }

    # Rows from sig_df whose signature has a prior override → keep all
    if prev_filled_sigs:
        preserved = sig_df[sig_df["signature"].astype(str).isin(prev_filled_sigs)].copy()
        preserved["_section"] = "user_override"
        preserved["_abs"] = pd.to_numeric(preserved.get("total_amount"), errors="coerce").abs().fillna(0)
        # Drop sample rows already in preserved (avoid dup)
        sample = sample[~sample["signature"].astype(str).isin(prev_filled_sigs)]
        review = pd.concat([preserved, sample], ignore_index=True)
    else:
        review = sample

    # Re-apply prior user input across all rows (sample rows get empty values
    # unless the sig is in prev_overrides)
    for col in (OVERRIDE_COL, NOTES_COL):
        review[col] = review["signature"].astype(str).map(
            lambda s: prev_overrides.get(s, {}).get(col, "")
        )

    # Orphan-rescue: prior-override sigs whose signature no longer appears in
    # sig_df (e.g., text drift after a normalize change) — keep them as bare
    # rows so the work isn't lost. They won't have sig_df enrichment but the
    # signature + correct_horizontal pair is enough for step3's override path.
    sig_set = set(review["signature"].astype(str))
    orphans = prev_filled_sigs - sig_set
    if orphans:
        orphan_rows = []
        for sig in orphans:
            d = prev_overrides[sig]
            orphan_rows.append({
                "signature": sig,
                OVERRIDE_COL: d.get(OVERRIDE_COL, ""),
                NOTES_COL: d.get(NOTES_COL, ""),
                "_section": "user_override_orphan",
            })
        review = pd.concat([review, pd.DataFrame(orphan_rows)], ignore_index=True)

    display_cols = FEEDBACK_DISPLAY_COLS + ["_section"]
    cols_present = [c for c in display_cols if c in review.columns]
    review = review[cols_present]
    review.to_excel(path, index=False)
    n_overrides = (review[OVERRIDE_COL].astype(str).str.strip().ne("")
                                       & review[OVERRIDE_COL].astype(str).str.strip().ne("nan")).sum()
    print(f"  [feedback] wrote review template: {path.name}  ({len(review)} rows, "
          f"{n_overrides} with user overrides preserved)")
    return path


def build_few_shot_block(examples: list[dict]) -> str:
    """Render examples as a string to append to the LLM system prompt."""
    if not examples:
        return ""
    lines = ["", "USER-VERIFIED EXAMPLES (treat these patterns as authoritative):"]
    for i, ex in enumerate(examples, 1):
        line = (
            f"  {i}. account_desc='{ex.get('account_desc', '')}' "
            f"desc='{ex.get('desc_norm', '')}' → {ex.get('correct_horizontal')}"
        )
        if ex.get("notes"):
            line += f"   ({ex['notes']})"
        lines.append(line)
    return "\n".join(lines)
