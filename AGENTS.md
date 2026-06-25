# Instructions for Codex in this repo

## Git commits

**Do NOT add `Co-Authored-By: Codex ...` (or any AI-coauthor trailer) to commit messages.**

The user does not want Codex credited in commit history. When committing, write a plain
message with no Co-Authored-By line and no `🤖 Generated with Codex` footer.

This applies to:
- `git commit` messages (including `--amend`)
- PR descriptions created via `gh pr create`

If you have already added such a trailer, amend the commit to remove it before pushing.
If it has already been pushed, amend + force-push (the user has authorized force-push for
this specific cleanup case).

## Parameter / config file edits (`conf/**` — `conf/company_*/parameters.yml`, `conf/base/categories.yml`, `conf/table_2/parameters.yml`, etc.)

**ALL `conf/**` files are edited by Codex DIRECTLY in the local Mac files (Edit tool).**
`conf/` is gitignored and is **NOT synced through git** — the user does NOT `git pull` to
get config. The user copies config from the **Mac side directly** (manual copy Mac→Windows).

Therefore, **after every conf edit Codex MUST guide the user explicitly**:
1. Edit the local Mac `conf/...` file directly with the Edit tool.
2. Then give the user, in the reply:
   - the **full path** edited,
   - the **exact text block to copy** + **exactly where to put it** (anchor line),
   - a **verify command** (e.g. `findstr /c:"..." conf\...`),
   - what to run after.
3. Do NOT assume the user will diff/find the change themselves, and NEVER assume a conf
   change reaches Windows via git — it does not. Spell out the copy.

Contrast with **code** (`src/`, `scripts/`): that DOES go via git (user grabs from GitHub,
e.g. a SHA-pinned raw URL). Only `conf/**` is the manual Mac-copy channel.

This overrides the earlier "no edit gitignored conf" guidance — that approach caused
churn. User confirmed 2026-05-28 (edit local + describe) and 2026-05-30 (conf is manual
Mac-copy, not git; always spell out the exact block + location to paste).

## `results/` folder — communication scratch (gitignored)

The user drops script-output `.txt` / `.tsv` files into `results/` so Codex can read
them (chat paste is limited to text ≤~3000 rows; this folder bypasses that). Conventions:

- **It is gitignored** — never commit it, never `git add` it.
- **Ephemeral** — the user deletes files there intermittently. Do NOT rely on a file in
  `results/` persisting after a given run; read it in the turn it's referenced, extract
  what's needed, then treat it as gone. Don't reference old `results/` contents later.
- It is purely a Mac-side Codex↔user channel for large outputs (rule audits, sig dumps,
  pasted-back Excel splits, project-team golden masters for reconciliation).
- When the user says "結果喺 results/" — go read the relevant files immediately.

## Every-turn habits (user asked 2026-05-31 — applies to EVERY Codex)

1. **End every turn with (a) a progress table and (b) a "要提你嘅 / reminders" list.**
   - The 6×2 progress table = entity (galaxy/sjm/wynn/vml/melco/mgm) × {25, 24}, showing where each stands. Keep it current.
   - Reminders = open questions + the standing "⏰ 項目組 rules (25/24/23) user will give separately — remind them" item. User fears forgetting; never drop these.
2. **Do NOT spell out how to fetch files from GitHub** (no curl / Invoke-WebRequest walkthroughs). The user knows how to sync. Just say what changed / what to grab; they handle transfer. (conf = manual Mac→Win copy; code = via git.)
3. **Periodically clean up `results/` and `scripts/`** — delete consumed/superseded files (scripts are git-tracked = recoverable; results/ is ephemeral). Don't let them accumulate.

## Per-entity classification method (settled on VML 25, 2026-05-31)

All 6 entities have the **vertical (V) label problem** (per-project broadcast / unreliable LLM).
The fix per entity = **use the project team's own manual classification column** (they label
their data; that column is the ground truth we lack via 表1/2), mapped to OUR taxonomy:

- Find the project-team manual category column in the raw/tagged_rows (e.g. VML `分類1`;
  Wynn `項目性質` + `Nature of Expenses` + `comp费用大类`; Galaxy 一級/二級標籤; etc.).
- Build an **internal mapping** (their category → our V_* / H_*) since taxonomies differ.
- Apply via step4 `row_vertical_overrides` / `row_horizontal_overrides` (`column_map` for a
  whole category column; `when/set` for targeted account/keyword/helper-column rules).
- Keep OUR-taxonomy exceptions where ours is better (e.g. VML 健康養生 → V_WELLNESS).
- Then reconcile against the project team's audit/golden (`build_vml_check.py`) + give them
  the deliverable (`build_master_audit_25.py`) to re-check.
