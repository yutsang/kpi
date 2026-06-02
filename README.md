# KPI Pipeline

Produces per-company Excel reports (pivot + full tagged rows) and a PowerBI-ready parquet.

---

## Quick start

```bash
# 1. Select an entity and run the full pipeline (steps 0–5)
kedro run --pipeline=company_1

# 2. Run all configured entities at once
kedro run --pipeline=all

# 3. After reviewing signatures, generate reports only
kedro run --pipeline=company_1_audit
```

---

## Configuration

```
conf/
├── base/
│   ├── parameters.yml      # shared defaults (concurrency, LLM settings, reporting_cycle)
│   └── categories.yml      # vertical / horizontal / NG taxonomy (all 6 companies share this)
├── local/
│   ├── credentials.yml     # LLM api_base / api_key / chat_model  ← gitignored
│   └── parameters.yml      # local overrides                       ← gitignored
└── company_1/
│   └── parameters.yml      # alias, name, raw_filename, column mappings
└── company_2/ … company_6/
```

**Each `conf/company_N/parameters.yml` defines:**

| Field | Purpose |
|---|---|
| `alias` | Short code used in Kedro commands and folder names |
| `name` | Display name |
| `raw_filename` | Source file placed in `data/<alias>/raw/` |
| `raw_sheet` | Sheet name or index (0-based) |
| `columns.*` | Maps pipeline column roles to actual Excel column names |
| `account_code_system` | Tells the LLM which GL/account format to expect |
| `prior_raw_files` | Prior-year source files for multi-year stacking |
| `prebuild_sources` | Source extraction config (company_6 only) |

**Data folders** are auto-created under `data/<alias>/`:
```
data/<alias>/
├── raw/          ← place source files here
├── interim/      ← parquet + xlsx produced by steps 0–4
│   └── cache/    ← LLM response cache (safe to delete to force re-tag)
└── output/       ← final Excel reports + audit files
```

**Central audit folder** — all audit files also land here:
```
data/audit/       ← company_N_audit_review.xlsx for all companies
                  ← sig_sample_YYYY-MM-DD.xlsx from sample_sigs tool
```

---

## Pipeline steps

Each step is idempotent — re-running is safe. Delete interim files to force re-run.

### Step 0 — Convert (`step0_convert`)
Reads raw source (`.xlsx` or `.parquet`) → writes `<company>_raw.parquet`.  
Validates that all required columns are present; prints the actual column list if not.

> **company_6 only:** Run the prebuild step first (see below) to generate the raw parquet from multiple source files.

### Step 1 — Extract (`step1_extract`)
Extracts unique **projects**, **signatures**, **accounts**, **vendors** from the raw parquet.

**Signature** = `account_code | account_desc | normalized(description)` — the unit for 橫向 tagging.

Outputs:
- `<company>_unique_projects.xlsx` — one row per project; fill `manual_vertical` to override LLM
- `<company>_unique_signatures.xlsx` — one row per signature; fill `manual_horizontal` to override LLM
- `<company>_unique_accounts.xlsx`
- `<company>_unique_vendors.xlsx`
- `<company>_summary.txt` — row counts, period ranges, amount totals

### Step 2 — Tag projects (`step2_tag_projects`)
LLM classifies each unique project → **縱向 (vertical)** category.

Priority: `manual_vertical` (user) → `llm_vertical` (LLM)

Writes back to `<company>_unique_projects.xlsx`:
`llm_vertical`, `llm_vertical_label`, `llm_confidence`, `llm_reasoning`, `llm_alt_candidates`

### Step 3 — Tag signatures (`step3_tag_signatures`)
LLM classifies each unique signature → **橫向 (horizontal)** spend type.

Priority: user feedback overrides → `manual_horizontal` → `llm_horizontal`

Writes back to `<company>_unique_signatures.xlsx`:  
`llm_horizontal`, `llm_horizontal_label`, `llm_confidence`, `llm_reasoning`

Also writes `<company>_feedback.xlsx` (~200 rows) to `output/` for human review.

### Step 4 — Apply tags (`step4_apply_tags`)
Joins project + signature tags to every row.

Adds columns to each row:
- `vertical_id`, `vertical_label`, `vertical_source`
- `horizontal_id`, `horizontal_label`, `horizontal_source`
- `final_capex_opex` (normalised Capex/Opex)
- `ng_scope` (gaming / non_gaming)
- `row_type` (normal / adjustment)

Outputs: `<company>_tagged_rows.parquet`

### Step 5 — Build report (`step5_build_report`)
Produces `<company>_kpi_report.xlsx` with:

| Sheet | Contents |
|---|---|
| `0_index` | Sheet directory |
| `1_master_pivot` | Vertical × horizontal cross-tab with row counts |
| `2_year_<YYYY>` | Same pivot split by posting year (one sheet per year) |
| `6_audit_top500` | Top 500 rows by amount for drill-through |
| `7_all_rows` | Up to 50,000 rows sorted by amount (full fact table for Excel users) |

The `<company>_tagged_rows.parquet` is also the PowerBI source.  
> PowerBI: rows=`vertical_label`, columns=`horizontal_label`, values=`Sum(amount)`, slicers=`final_capex_opex`, `period`, `posting_year`

---

## Side pipelines

Run these explicitly after the main pipeline.

```bash
# Single company
kedro run --pipeline=company_1_audit
kedro run --pipeline=company_1_offsets
kedro run --pipeline=company_1_export
kedro run --pipeline=company_1_validate
kedro run --pipeline=company_1_summary

# All companies at once
kedro run --pipeline=audit
kedro run --pipeline=offsets
```

| Pipeline | When to run | Output |
|---|---|---|
| `audit` | After step 3 | `<company>_audit_review.xlsx` — low-conf + H_OTHER + top sigs + account cross-ref |
| `offsets` | After step 4 | `<company>_offsets_review.xlsx` — offset/reversal detection (4 layers) |
| `export` | After step 4 | Tagged rows as xlsx (for users without parquet tooling) |
| `validate` | After step 3–4 | LLM tagging quality checks |
| `summary` | Any time after step 3 | Re-prints step 3 statistics |

### Prebuild (company_6 only)

company_6's raw data comes from multiple source Excel files that require custom extraction.  
Place the source files in `data/<alias>/raw/` (filenames defined in `conf/company_6/parameters.yml` under `prebuild_sources`), then:

```bash
kedro run --pipeline=company_6_prebuild
# then run the main pipeline as normal:
kedro run --pipeline=company_6
```

---

## Manual correction tools

After step 3, review and fix LLM tags before building the report.

### Generate correction workbook

```bash
# Single company
python -m kpi.tools.quickfix --entity company_1

# All companies at once
python -m kpi.tools.quickfix --all
```

Output: `data/<alias>/output/<company>_quickfix.xlsx`

| Sheet | What to do |
|---|---|
| `0_zero_analysis` | Verticals/horizontals showing ⚠ ZERO — start here |
| `1_projects` | Fill `fix_vertical` column with a vertical ID (e.g. `V_CONCERT`) |
| `2_signatures` | Fill `fix_horizontal` column with a horizontal ID (e.g. `H_FNB`) |
| `3_untagged_rows` | Rows still untagged after step 4 — find their projects/signatures |
| `4_valid_ids` | Reference list of all valid vertical/horizontal IDs |

### Apply corrections

```bash
python -m kpi.tools.apply_quickfix --entity company_1
```

Writes your `fix_vertical` / `fix_horizontal` entries back to `unique_projects.xlsx` and  
`unique_signatures.xlsx`, then automatically re-runs step 4 + step 5.

---

## Signature sampling (for LLM prompt diagnosis)

```bash
# All companies → data/audit/sig_sample_YYYY-MM-DD.xlsx
python -m kpi.tools.sample_sigs

# Single company
python -m kpi.tools.sample_sigs --entity company_2
```

Produces ~60 rows per company (per-company sheet + `ALL_combined`):
- No colour = top amount (must be tagged correctly)
- **Yellow** = H_OTHER (LLM fell back to '其他')
- **Red** = low confidence < 0.7 (LLM was unsure)

Share the `ALL_combined` sheet for cross-company prompt diagnosis.

---

## Multi-year data

To include prior years (e.g. 2023, 2024) alongside 2025:

1. Place prior-year source files in `data/<alias>/raw/`
2. Add them to `conf/company_N/parameters.yml`:
   ```yaml
   prior_raw_files:
     - filename: "company_N_actuals_2023.xlsx"
       raw_sheet: 0
     - filename: "company_N_actuals_2024.xlsx"
       raw_sheet: 0
   ```
3. Stack all years into one parquet:
   ```bash
   python -m kpi.lib.stack_years --entity company_N
   ```
4. Run the pipeline normally — step 5 auto-splits by `posting_year`.

> **company_6:** Generate prior-year parquets by running the prebuild step with the old source files first (rename the output), then run `stack_years`.

---

## Year rollover checklist

Update `conf/base/parameters.yml` → `reporting_cycle` section each year:

```yaml
reporting_cycle:
  current_year: 2025
  report_years: [2023, 2024, 2025]
  cutoff_month: null   # null = full year; "2025-03" = YTD March
```

Then for each entity:
1. Update `raw_filename` (or `prebuild_sources`) to point to the new year's file
2. Verify column names haven't changed in the new file
3. Delete `data/<alias>/interim/<company>_raw.parquet` and `<company>_tagged_rows.parquet`
4. Keep `<company>_unique_*.xlsx` if manual tags exist — they carry over
5. Re-run: `kedro run --pipeline=all`

---

## Force re-run a step

```bash
# Re-run only step 3 for company_1 (clears LLM cache for that step)
FORCE_REDO=step3_tag_signatures kedro run --pipeline=company_1

# Re-run everything
FORCE_REDO=all kedro run --pipeline=company_1
```

---

## Taxonomy reference

Defined in `conf/base/categories.yml` — shared across all companies.

**Verticals (縱向)** — 23 categories covering NG0–NG11 investment types  
**Horizontals (橫向)** — 8 spend types (H_CONSTRUCTION, H_LABOR, H_HOTEL_ROOM, H_FNB, H_VENUE, H_PROFESSIONAL, H_SPONSORSHIP, H_OTHER)  
**NG scopes** — NG0 (gaming) + NG1–NG11 (non-gaming)

---

## Troubleshooting

| Symptom | Fix |
|---|---|
| `No entity selected` | Pass `--entity company_N` or set `KPI_ENTITY=company_N` |
| Required columns missing | Check `conf/company_N/parameters.yml` columns mapping vs actual Excel headers |
| Vertical/horizontal all `(未分類)` | Step 2/3 not run yet, or LLM failed — check `llm_status` in unique_*.xlsx |
| Many `H_OTHER` | Run `sample_sigs` and review; likely a prompt issue for that company's account format |
| Zero in pivot cell | Run `quickfix` — check `0_zero_analysis` sheet, then fix in `1_projects` / `2_signatures` |
| LLM cache stale | Delete `data/<alias>/interim/cache/` and re-run step 2/3 |
| Duplicate column error in quickfix | Excel has duplicate column headers — quickfix auto-deduplicates, safe to ignore warning |
