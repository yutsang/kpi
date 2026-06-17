"""Step 3: LLM-tag each unique row signature with a 橫向 (horizontal) + row_type.

A signature = (account_code, account_desc, normalized(description)). LLM sees:
  - account_code + account_desc + description samples
  - sample projects this signature appears in
  - amount statistics including negative count (catches refunds/reversals)
  - candidate horizontals + row_types

Outputs llm_horizontal / llm_row_type / llm_confidence / llm_reasoning back to
<company>_unique_signatures.xlsx.

Run after Step 0 + Step 1.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

import pandas as pd
from tqdm import tqdm

from kpi.lib.conf import ROOT, load_config, load_categories, path  # noqa: E402
from kpi.lib.feedback import build_few_shot_block, load_feedback, write_feedback_template  # noqa: E402
from kpi.lib.heartbeat import Heartbeat  # noqa: E402
from kpi.lib.io_setup import force_unbuffered_io  # noqa: E402
from kpi.lib.llm import LLMClient  # noqa: E402

force_unbuffered_io()


SYSTEM_PROMPT = (
    "You are a domain expert in Macau gaming concessionaire (Galaxy / SJM / Wynn / "
    "Melco / VML / MGM) SAP/PMS spend categorization. Given a row signature pick the "
    "single best 橫向 (horizontal) spend type from the candidate list.\n\n"

    "16 CATEGORIES (apply the decision tree below):\n"
    "  H_CONSTRUCTION  建設與設施支出   — buildings, fit-out, MEP, structural\n"
    "  H_EQUIP         設施及器具採購   — movable equipment, slot machines, IT hardware, furniture\n"
    "  H_LABOR         人工成本         — salaries, payroll, payroll allocation\n"
    "  H_HOTEL_ROOM    酒店客房 [Comp]  — comp lodging / room nights / hotel tax tied to room\n"
    "  H_FNB           餐飲    [Comp]   — comp food / beverage\n"
    "  H_VENUE         活動場地 [Comp]  — Comp barter venue / Comp event venue ONLY\n"
    "  H_COMP_TICKET   贈票支出 [Comp]  — show tickets / event passes given to clients\n"
    "  H_COMP_OTHER    Comp 其他 [Comp] — other comp (spa, transport, retail, misc)\n"
    "  H_PROFESSIONAL  專業服務費       — consultancy / accounting / legal / agency / PR\n"
    "  H_PERFORMER     合約成本（演藝） — artist / performer / show production fees\n"
    "  H_SPONSORSHIP   贊助費           — outgoing sponsorship payments\n"
    "  H_ADVERTISING   廣告及推廣       — media buys / digital ads / outdoor / influencer\n"
    "  H_LICENSE       授權費           — IP licensing / software licence / venue licence\n"
    "  H_LEASE         租賃費           — office / warehouse / equipment / vehicle / IT lease;\n"
    "                                     ALSO commercial (non-Comp) event venue rental\n"
    "  H_MAINTENANCE   維護費           — R&M / repair / preventive maintenance\n"
    "  H_OTHER         其他             — true catch-all; AVOID when possible\n\n"

    "DECISION TREE (follow in order):\n"
    "  STEP 1 — IS IT COMP / CLIENT HOSPITALITY?\n"
    "    Trigger: account_desc starts with 'Comp ', or account_code in 7120/7131/7132/7141/7151/7153,\n"
    "             or description mentions comp/giveaway-to-client.\n"
    "    If yes, drill into the FIVE Comp sub-categories:\n"
    "      • Room / Lodging / Stay / Hotel tax    → H_HOTEL_ROOM\n"
    "      • Food / Beverage / F&B / 餐 / 飲     → H_FNB\n"
    "      • Comp barter venue / external venue swap → H_VENUE\n"
    "      • Comp show ticket / event pass / 贈票  → H_COMP_TICKET\n"
    "      • Anything else Comp (spa, transport, retail) → H_COMP_OTHER\n"
    "    If unambiguously Comp but unsure of sub-type → H_COMP_OTHER (not H_OTHER).\n\n"

    "  STEP 2 — IS IT CAPEX FACILITY / EQUIPMENT?\n"
    "    Trigger: account_code 1xxxxxx (FA / WIP / AUC) / 3610xxx (Accrued Projects),\n"
    "             OR Capex-dominant + fit-out / construction / hardware description.\n"
    "    Distinguish:\n"
    "      • Buildings, fit-out, structural, MEP, immovable fixtures  → H_CONSTRUCTION\n"
    "      • Movable equipment / slot machines / gaming tables / IT hardware /\n"
    "        furniture / kitchen equipment / cabinets / cameras       → H_EQUIP\n"
    "    Rule of thumb: 'can you wheel it out?' → H_EQUIP. 'is it built-in/structural?' → H_CONSTRUCTION.\n"
    "    Exceptions even when account_code says CAPEX:\n"
    "      • description has consultancy/advisory   → H_PROFESSIONAL\n"
    "      • description has F&B/coffee for project → H_FNB (not H_CONSTRUCTION)\n\n"

    "  STEP 3 — IS IT PAYROLL?\n"
    "    Trigger: account_code 8xxxxxx, account_desc has Salaries / Payroll / Wages / Bonus / Allocation - Payroll.\n"
    "    → H_LABOR.\n\n"

    "  STEP 4 — SERVICE / PAYMENT TYPE:\n"
    "    • Artist / performer / show production / DJ / band / 藝人 / 表演 / 演出費\n"
    "        → H_PERFORMER\n"
    "    • Sports team appearance / athlete fee for promo\n"
    "        → H_PERFORMER (treat athlete same as performer)\n"
    "    • Consultancy / advisory / accounting / legal / PR / market research / agency retainer\n"
    "        → H_PROFESSIONAL\n"
    "    • Sponsorship Fee, Donation, Charity (對外贊助 — concessionaire pays out)\n"
    "        → H_SPONSORSHIP\n"
    "    • Media buy / Digital ad / OOH billboard / TV / radio / SEO ad / influencer post fee /\n"
    "      Giveaways & Merchandise / Direct Mailing\n"
    "        → H_ADVERTISING\n"
    "    • IP / brand / software / venue license fee / DICJ permits\n"
    "        → H_LICENSE\n"
    "    • Rent / Lease / 租金: office, warehouse, equipment, vehicle, IT system,\n"
    "      AND commercial non-Comp venue rental for events\n"
    "        → H_LEASE\n"
    "    • R&M / Repair / Maintenance / Preventive Maintenance / 維修 / 保養\n"
    "        → H_MAINTENANCE\n"
    "    • Genuinely uncategorisable → H_OTHER (last resort)\n\n"

    "GAMING-INDUSTRY JARGON GLOSSARY (→ H_EQUIP for movables; H_CONSTRUCTION for fixed):\n"
    "  EGM / Slot Machine / 角子機 / Cabinet                 → H_EQUIP\n"
    "  Gaming table / Smart table / Roulette / Baccarat     → H_EQUIP\n"
    "  Kit Conversion / Conversion Kit (machine upgrade)    → H_EQUIP\n"
    "  Chip / Plaque / Chip purchase                        → H_EQUIP\n"
    "  Vendor names (gaming hardware): IGT, Konami, Sega Sammy, LNW (Light & Wonder),\n"
    "  Spintec, Aristocrat, Scientific Games, AGS, Galaxy Gaming  → H_EQUIP\n"
    "  Hikvision / HIK / CAM / PTRZ surveillance camera     → H_EQUIP (gaming-floor device)\n"
    "  LED display / Display Marketing / 4K screen          → H_EQUIP\n"
    "  IT HW / Server / Storage                              → H_EQUIP\n"
    "  MEP (Mechanical/Electrical/Plumbing)                  → H_CONSTRUCTION\n"
    "  Fit-out / Hard Cost / Soft Cost / Building works      → H_CONSTRUCTION\n"
    "  WIP / AUC / FA wrappers — default H_CONSTRUCTION unless description points to H_EQUIP\n\n"

    "CHART OF ACCOUNTS HINTS — Galaxy GL 7-digit codes (where account_code_system='galaxy_gl'):\n"
    "  1xxxxxx (FA / WIP / AUC / Accrued Projects)         → H_CONSTRUCTION or H_EQUIP (per Step 2)\n"
    "  3610xxx Accrued Projects                             → H_CONSTRUCTION (capital project accrual)\n"
    "  7120xxx Comp Lodging                                  → H_HOTEL_ROOM\n"
    "  7131xxx Comp Food                                     → H_FNB\n"
    "  7132xxx Comp Beverage                                 → H_FNB\n"
    "  7141020 Comp Service Charge — by description:\n"
    "      'Rooms - Service Charge' / 'Room'                 → H_HOTEL_ROOM\n"
    "      'F&B' / 'Food/Beverage'                           → H_FNB\n"
    "      'Spa / BTSpa / Foothub / Club Use'                → H_COMP_OTHER\n"
    "      Bare 'Service charge'                              → H_PROFESSIONAL\n"
    "  7141050 Comp Club Use / Spa                          → H_COMP_OTHER\n"
    "  7141078 Comp External - Barter Events                → H_VENUE\n"
    "  7141080 Comp External - Others                       → H_COMP_OTHER (default) or specific\n"
    "  7151xxx Comp Tourism Tax (tied to hotel)             → H_HOTEL_ROOM\n"
    "  7153xxx Comp External - Transportation               → H_COMP_OTHER\n"
    "  7511010 Sponsorship Fee                              → H_SPONSORSHIP\n"
    "  7511100 Giveaways & Direct Mailing / Merchandise     → H_ADVERTISING\n"
    "  7611010 Production / Production Service              → H_PERFORMER (if show prod) else H_PROFESSIONAL\n"
    "  7615000 Digital Marketing — by description:\n"
    "      agency retainer                                   → H_PROFESSIONAL\n"
    "      endorsement / influencer / KOL                    → H_ADVERTISING\n"
    "      generic ad placement / media buy                  → H_ADVERTISING\n"
    "  7616000 Marketing - Barter                           → H_VENUE (Comp swap) or H_SPONSORSHIP\n"
    "  7710100 Special Events — by description:\n"
    "      external venue / barter                           → H_VENUE\n"
    "      commercial venue rental                           → H_LEASE\n"
    "      pure performer fee                                → H_PERFORMER\n"
    "      generic event ops                                 → H_ADVERTISING\n"
    "  7921xxx Professional and Consultancy Fee             → H_PROFESSIONAL\n"
    "  7926xxx Contract Services — by description:\n"
    "      consulting                                        → H_PROFESSIONAL\n"
    "      hall / venue rental (commercial)                  → H_LEASE\n"
    "      performer / show production                       → H_PERFORMER\n"
    "  7958xxx Allocation - R&M                             → H_MAINTENANCE\n"
    "  8xxxxxx Salaries / Payroll / Wages / Payroll Allocation → H_LABOR\n"
    "  8107350 Venue License (DICJ)                         → H_LICENSE\n\n"

    "Non-Galaxy entities (SJM SAP / Wynn / VML / Melco internal codes): the number prefixes\n"
    "do NOT match Galaxy. Rely on account_desc + description, and use the decision tree above.\n\n"

    "AVOID H_OTHER. Try harder before falling back:\n"
    "  • Empty/generic 'Accrual for PO#xxx' → use account_code + Capex/Opex.\n"
    "  • Vendor-name-only → infer (gaming vendor → H_EQUIP; consulting firm → H_PROFESSIONAL;\n"
    "    artist agency → H_PERFORMER; media agency → H_ADVERTISING).\n"
    "  • Repeating Food / F&B / Beverage → H_FNB regardless of NG context.\n"
    "  • Repeating Room / Lodging → H_HOTEL_ROOM.\n"
    "  • Clearly Comp but sub-type unclear → H_COMP_OTHER (not H_OTHER).\n\n"

    "OUTPUT RULES:\n"
    "  The horizontal_id you pick MUST appear in the candidate list.\n"
    "  Always reply in valid JSON only."
)

BATCH_SYSTEM_PROMPT = SYSTEM_PROMPT + (
    "\n\nBATCH MODE: You will be given a BATCH of N signatures. Reply with a single JSON "
    'object {"results": [...]} where the array has exactly N entries in the SAME order '
    "as the input items. Each entry must contain horizontal_id, confidence, reasoning."
)


def _sig_block(sig_row: dict) -> list[str]:
    co_dom = str(sig_row.get("capex_opex_dominant", "") or "").strip()
    co_pct = sig_row.get("capex_opex_dominant_pct", "")
    co_line = ""
    if co_dom:
        try:
            co_line = f"Capex/Opex (dominant): {co_dom} ({int(co_pct)}% of rows)"
        except (ValueError, TypeError):
            co_line = f"Capex/Opex (dominant): {co_dom}"

    row_count = sig_row.get("row_count", "")
    total_amt = sig_row.get("total_amount", "")
    has_neg = sig_row.get("has_negative", False)
    try:
        stat_line = (
            f"Stats: row_count={int(row_count):,}  "
            f"total_amount(MOP)={float(total_amt):,.0f}  "
            f"has_negative={'yes' if bool(has_neg) else 'no'}"
        )
    except (ValueError, TypeError):
        stat_line = ""

    block = [
        f"Account code: {sig_row['account_code']}",
        f"Account description: {sig_row['account_desc']}",
        f"Normalized description: {sig_row['desc_norm']}",
        f"Description samples (raw): {sig_row['desc_samples']}",
    ]
    if co_line:
        block.append(co_line)
    if stat_line:
        block.append(stat_line)
    return block


def _candidates_block(horizontals: list[dict]) -> list[str]:
    lines = ["Candidate horizontals (pick exactly ONE id):"]
    for h in horizontals:
        lines.append(f"  - id: {h['id']}   label: {h['label']}")
    return lines


def build_user_prompt(sig_row: dict, horizontals: list[dict]) -> str:
    lines = _sig_block(sig_row)
    lines.append("")
    lines.extend(_candidates_block(horizontals))
    lines.append("")
    lines.append(
        "Reply with JSON only, schema:\n"
        "{\n"
        '  "horizontal_id": "<one of candidate horizontal ids>",\n'
        '  "confidence": <float 0..1>,\n'
        '  "reasoning": "<one sentence in Chinese>"\n'
        "}"
    )
    return "\n".join(lines)


def build_batch_user_prompt(batch: list[dict], horizontals: list[dict]) -> str:
    lines = [f"BATCH of {len(batch)} signatures. Process each in order.", ""]
    for i, sig in enumerate(batch, 1):
        lines.append(f"--- Item {i} ---")
        lines.extend(_sig_block(sig))
        lines.append("")
    lines.extend(_candidates_block(horizontals))
    lines.append("")
    lines.append(
        "Reply with JSON ONLY, schema:\n"
        "{\n"
        '  "results": [\n'
        '    {"horizontal_id": "...", "confidence": 0.0, "reasoning": "..."},\n'
        "    ...\n"
        "  ]\n"
        "}\n"
        f"The 'results' array MUST have exactly {len(batch)} entries in the same order as the input items."
    )
    return "\n".join(lines)


_ACCOUNT_CODE_SYSTEM_HINTS: dict[str, str] = {
    "galaxy_gl": (
        "ACCOUNT CODE FORMAT: This entity uses Galaxy-style 7-digit GL codes "
        "(e.g. 7120xxx Comp Lodging, 7131xxx Comp F&B, 7511010 Sponsorship, "
        "7921xxx Prof Fee, 8xxxxxx Payroll, 1xxxxxx Fixed Assets). "
        "The CHART OF ACCOUNTS HINTS in the instructions apply directly."
    ),
    "sjm_sap_cost_element": (
        "ACCOUNT CODE FORMAT: This entity uses SAP Cost Elements (6-digit codes, "
        "e.g. 4xxxxx materials, 6xxxxx services, 9xxxxx overheads). "
        "The Galaxy GL prefix hints (7120xxx etc.) do NOT apply. "
        "Rely primarily on account_desc (Cost element descr.) and description to classify."
    ),
    "wynn_internal": (
        "ACCOUNT CODE FORMAT: This entity uses Wynn internal account numbers. "
        "The Galaxy GL prefix hints do NOT apply. "
        "Use account_desc ('Nature of Expenses') and description as the primary classification signals."
    ),
    "vml_internal": (
        "ACCOUNT CODE FORMAT: This entity uses VML internal account codes. "
        "The Galaxy GL prefix hints do NOT apply. "
        "Use account_desc ('A/C Name') and description ('Transaction Description') as the primary signals."
    ),
    "melco_internal": (
        "ACCOUNT CODE FORMAT: This entity uses Melco internal ledger account IDs. "
        "The Galaxy GL prefix hints do NOT apply. "
        "Use account_desc (ledger_account) and description (line_memo) as the primary signals."
    ),
}


def _build_entity_system_prompt(base: str, account_code_system: str) -> str:
    hint = _ACCOUNT_CODE_SYSTEM_HINTS.get(account_code_system, "")
    if not hint:
        return base
    return f"{hint}\n\n{base}"


def main():
    cfg = load_config()
    cats = load_categories()
    company = cfg["company"]["code"]
    interim = path(cfg, "interim_dir")
    master = cfg["_master"]
    ent = (master.get("companies") or {}).get(company, {})
    account_code_system = ent.get("account_code_system", "")

    sig_xlsx = interim / f"{company}_unique_signatures.xlsx"
    if not sig_xlsx.exists():
        raise FileNotFoundError("Run Step 1 first to produce unique_signatures.xlsx")

    sig_df = pd.read_excel(sig_xlsx, engine="openpyxl")

    # KPI_SKIP_STEP3=1 — 完全 skip 本 step (連 LLM)，沿用現有 sig 檔嘅 H tag。
    # 啱「唔關 H 分類」嘅 raw 修改 (opex/capex 手調、少量污染)：del raw → step0 重建 → step4 re-apply，
    # 但 H 維持上次結果，慳成個 LLM pass。注意：改咗 H feedback / 規則 / categories label 就唔好設呢個掣。
    import os as _os
    if _os.environ.get("KPI_SKIP_STEP3"):
        _hcols = [c for c in ("manual_horizontal", "llm_horizontal") if c in sig_df.columns]
        _tagged = pd.Series(False, index=sig_df.index)
        for _c in _hcols:
            _tagged |= sig_df[_c].astype(str).str.strip().replace({"nan": "", "None": ""}).ne("")
        if _tagged.any():
            print(f"  [skip] KPI_SKIP_STEP3=1 — 沿用現有 {int(_tagged.sum()):,}/{len(sig_df):,} 個 sig 嘅 H tag，"
                  f"唔跑 LLM (opex/capex/污染 等唔關 H 嘅修改用呢個；改 H feedback/rule/label 就唔好設)", flush=True)
            return
        print("  [skip] KPI_SKIP_STEP3=1 但 sig 檔暫無 H tag → 照常跑 LLM", flush=True)

    sig_df["account_code"] = sig_df["account_code"].astype("string").fillna("")
    sig_df["account_desc"] = sig_df["account_desc"].astype("string").fillna("")
    sig_df["desc_norm"] = sig_df["desc_norm"].astype("string").fillna("")
    sig_df["desc_samples"] = sig_df["desc_samples"].astype("string").fillna("")
    sig_df["project_samples"] = sig_df["project_samples"].astype("string").fillna("")

    horizontals = [h for h in cats["horizontals"] if h["id"] != "H_COUNT"]
    valid_h_ids = {h["id"] for h in horizontals}
    h_label = {h["id"]: h["label"] for h in horizontals}

    def derive_row_type(has_negative: bool) -> str:
        return "adjustment" if has_negative else "normal"

    rt_label = {"normal": "正常支出", "adjustment": "會計調整/退款"}

    cache_path = interim / "cache" / f"{company}_step3_signatures.jsonl"
    llm = LLMClient(cfg, cache_path=cache_path)

    output_dir = path(cfg, "output_dir")
    overrides, few_shot = load_feedback(output_dir, company)

    # Build entity-aware base prompts (inject account_code_system hint before global instructions)
    entity_system = _build_entity_system_prompt(SYSTEM_PROMPT, account_code_system)
    entity_batch_system = _build_entity_system_prompt(BATCH_SYSTEM_PROMPT, account_code_system)
    if account_code_system:
        print(f"  [step3] account_code_system={account_code_system!r} — entity-specific GL hint injected", flush=True)

    if few_shot:
        SYSTEM_PROMPT_FINAL = entity_system + build_few_shot_block(few_shot)
        BATCH_SYSTEM_PROMPT_FINAL = entity_batch_system + build_few_shot_block(few_shot)
    else:
        SYSTEM_PROMPT_FINAL = entity_system
        BATCH_SYSTEM_PROMPT_FINAL = entity_batch_system

    payloads_all = sig_df.to_dict(orient="records")
    print(f"Total unique signatures: {len(payloads_all)}.")

    pre_results: list[dict | None] = [None] * len(payloads_all)
    feedback_hits = 0
    for i, sig in enumerate(payloads_all):
        sig_key = str(sig.get("signature", ""))
        if sig_key and sig_key in overrides:
            pre_results[i] = {
                "horizontal_id": overrides[sig_key],
                "confidence": 1.0,
                "reasoning": "User-verified feedback override",
                "_tag_source": "user_feedback",
            }
            feedback_hits += 1
    if feedback_hits:
        print(f"  [feedback] {feedback_hits} signatures resolved by user override (skip LLM)")

    # Predominant rules from parameters.yml — applied to remaining unmatched signatures.
    # Three rule modes:
    #   then: <H_ID>          → skip LLM, use that ID directly
    #   then_scope: <scope>   → narrow LLM candidate list to horizontals with matching scope
    #   then_any_of: [...]    → narrow LLM candidate list to given ids
    from kpi.lib.rules import load_predominant_rules, find_predominant_rule, resolve_then
    pr_rules = load_predominant_rules(cfg, kind="horizontal")
    valid_h_ids = {h["id"] for h in cats.get("horizontals", [])}
    rule_skip_hits = 0
    rule_constrain_hits = 0
    if pr_rules:
        for i, sig in enumerate(payloads_all):
            if pre_results[i] is not None:
                continue
            rule, idx = find_predominant_rule(pr_rules, sig)
            if not rule:
                continue
            resolved = resolve_then(rule, horizontals)
            if resolved["mode"] == "fixed":
                target = resolved["id"]
                if target in valid_h_ids:
                    pre_results[i] = {
                        "horizontal_id": target,
                        "confidence": 0.99,
                        "reasoning": f"Predominant rule #{idx} ({rule.get('if')!r}) → fixed",
                        "_tag_source": f"rule:{idx}:fixed",
                    }
                    rule_skip_hits += 1
            elif resolved["mode"] == "subset":
                subset_ids = [hid for hid in resolved["ids"] if hid in valid_h_ids]
                if subset_ids:
                    # Attach constrained candidate list to the sig dict — call_one picks it up,
                    # and the batched path is skipped for these items.
                    sig["_candidates"] = [h for h in horizontals if h["id"] in subset_ids]
                    sig["_constraint_rule_idx"] = idx
                    rule_constrain_hits += 1
        if rule_skip_hits:
            print(f"  [rules] {rule_skip_hits} signatures skipped LLM (fixed-target rules)")
        if rule_constrain_hits:
            print(f"  [rules] {rule_constrain_hits} signatures with narrowed candidate lists (subset rules)")

    # Conf flag `skip_signature_llm: true` — for entities whose H is fully driven by a project-team
    # column_map applied at step4 (e.g. VML 分類1/分類1.1). Don't run the slow sig LLM at all: every
    # still-unresolved sig is defaulted (H_OTHER), and step4's row_horizontal_overrides set the real H.
    # Feedback overrides + predominant rules above still win. Avoids a multi-hour wasted LLM pass.
    if bool(ent.get("skip_signature_llm", False)):
        _dft = ent.get("signature_default_horizontal", "H_OTHER")
        _n = 0
        for i in range(len(payloads_all)):
            if pre_results[i] is None:
                pre_results[i] = {
                    "horizontal_id": _dft, "confidence": 0.0,
                    "reasoning": "signature LLM skipped (conf skip_signature_llm); final H from step4 row overrides",
                    "_tag_source": "skip_llm",
                }
                _n += 1
        print(f"  [skip-llm] signature LLM disabled — {_n} sigs defaulted to {_dft} "
              f"(step4 column_map / overrides set final H)", flush=True)

    payloads = [p for p, r in zip(payloads_all, pre_results) if r is None]
    batched_payloads = [p for p in payloads if not p.get("_candidates")]
    single_payloads = [p for p in payloads if p.get("_candidates")]
    print(
        f"  LLM workload: {len(payloads)} signatures "
        f"({len(batched_payloads)} batched + {len(single_payloads)} constrained single-call)"
    )

    heartbeat = Heartbeat(
        path=interim / f"{company}_step3_progress.json",
        step_name="step3_signatures",
        total=len(payloads),
    )
    llm.attach_heartbeat(heartbeat)

    batch_size = int(cfg.get("llm", {}).get("step3_batch_size", 10) or 10)

    def coerce_one(r: dict) -> dict:
        h = r.get("horizontal_id")
        if h not in valid_h_ids:
            r["_invalid_horizontal"] = h
            r["horizontal_id"] = "H_OTHER"
            r["confidence"] = min(float(r.get("confidence", 0.0)), 0.3)
        return r

    def parse_batch(response: dict, batch_items: list[dict]) -> list[dict]:
        arr = response.get("results")
        if not isinstance(arr, list):
            raise ValueError(f"batch response missing 'results' array: {list(response.keys())}")
        return [coerce_one(dict(r)) for r in arr]

    # Phase config: batched for first N rounds, then one-by-one for remainder
    total_rounds   = int(cfg.get("llm", {}).get("step3_max_rounds", llm.max_round_retries) or llm.max_round_retries)
    batch_switch   = int(cfg.get("llm", {}).get("step3_batch_switch_round", 5) or 5)
    batch_rounds   = min(batch_switch, total_rounds)
    single_rounds  = max(0, total_rounds - batch_rounds)

    def call_one(row):
        # Use per-row constrained candidate list when present (set by subset rules).
        cands = row.get("_candidates") or horizontals
        return coerce_one(
            llm.chat_json(
                system=SYSTEM_PROMPT_FINAL,
                user=build_user_prompt(row, cands),
            )
        )

    def _is_err(r):
        return r is None or (isinstance(r, dict) and "_error" in r)

    if not payloads:
        results: list = []
    elif batch_size <= 1 or batch_rounds == 0:
        llm.max_round_retries = total_rounds
        results = llm.map_parallel(payloads, call_one, desc="step3 signatures")
    else:
        # Phase 1a: batched path — only items WITHOUT constrained candidate lists
        llm.max_round_retries = batch_rounds
        batched_results = (
            llm.map_batched(
                items=batched_payloads,
                cache_user_fn=lambda row: build_user_prompt(row, horizontals),
                cache_system=SYSTEM_PROMPT_FINAL,
                batch_user_fn=lambda batch: build_batch_user_prompt(batch, horizontals),
                batch_system=BATCH_SYSTEM_PROMPT_FINAL,
                parse_batch_fn=parse_batch,
                desc="step3 signatures (batched)",
                batch_size=batch_size,
            )
            if batched_payloads else []
        )
        # Phase 1b: single-call path — items WITH constrained candidate lists
        single_results = (
            llm.map_parallel(single_payloads, call_one, desc="step3 signatures (constrained)")
            if single_payloads else []
        )
        # Merge results back into the order of `payloads`.
        b_iter = iter(batched_results)
        s_iter = iter(single_results)
        results = [
            (next(s_iter) if p.get("_candidates") else next(b_iter))
            for p in payloads
        ]
        # Phase 2: retry any still-failing items one-by-one
        still_failing = [i for i, r in enumerate(results) if _is_err(r)]
        if still_failing and single_rounds > 0:
            print(f"  {len(still_failing)} items still failing after batch phase — switching to one-by-one ({single_rounds} more rounds)")
            llm.max_round_retries = single_rounds
            sub_payloads = [payloads[i] for i in still_failing]
            sub_results  = llm.map_parallel(sub_payloads, call_one, desc="step3 signatures (1-by-1)")
            for idx, res in zip(still_failing, sub_results):
                results[idx] = res

    # Splice LLM results back into the full result array (pre_results has rule/manual/rerank hits).
    full_results: list[dict] = []
    llm_iter = iter(results)
    for r in pre_results:
        if r is not None:
            full_results.append(r)
        else:
            try:
                full_results.append(next(llm_iter))
            except StopIteration:
                full_results.append({"_error": "missing LLM result"})
    results = full_results

    def is_err(r: dict) -> bool:
        return isinstance(r, dict) and "_error" in r

    derived_row_types = [
        derive_row_type(bool(p.get("has_negative", False))) for p in payloads_all
    ]
    out = pd.DataFrame({
        "signature": [p["signature"] for p in payloads_all],
        "llm_status": ["needs_manual" if is_err(r) else "ok" for r in results],
        "tag_source": [r.get("_tag_source", "llm") if not is_err(r) else "" for r in results],
        "llm_horizontal": ["" if is_err(r) else r.get("horizontal_id", "H_OTHER") for r in results],
        "llm_horizontal_label": ["" if is_err(r) else h_label.get(r.get("horizontal_id", ""), "") for r in results],
        "row_type": ["" if is_err(r) else rt for r, rt in zip(results, derived_row_types)],
        "row_type_label": ["" if is_err(r) else rt_label.get(rt, "") for r, rt in zip(results, derived_row_types)],
        "llm_confidence": [None if is_err(r) else r.get("confidence") for r in results],
        "llm_reasoning": ["" if is_err(r) else r.get("reasoning") for r in results],
        "llm_error": [r.get("_error", "") if is_err(r) else "" for r in results],
        "llm_invalid_horizontal": ["" if is_err(r) else r.get("_invalid_horizontal", "") for r in results],
    })

    drop_cols = [c for c in ["llm_status", "tag_source", "llm_horizontal", "llm_horizontal_label", "row_type", "row_type_label", "llm_confidence", "llm_reasoning", "llm_error", "llm_invalid_horizontal"] if c in sig_df.columns]
    sig_df = sig_df.drop(columns=drop_cols, errors="ignore")
    merged = sig_df.merge(out, on="signature", how="left")
    merged.to_excel(sig_xlsx, index=False)

    needs_manual = (merged["llm_status"] == "needs_manual").sum()
    rule_tagged = merged["tag_source"].fillna("").astype(str).str.startswith("rule:").sum()
    rerank_tagged = merged["tag_source"].fillna("").astype(str).str.startswith("rerank:").sum()
    llm_tagged = ((merged["llm_status"] == "ok") & (merged["tag_source"] == "llm")).sum()
    low_conf = ((merged["tag_source"] == "llm") & (merged["llm_confidence"] < 0.7)).sum()
    adjustments = ((merged["llm_status"] == "ok") & (merged["row_type"] == "adjustment")).sum()
    print(f"Done. Wrote {sig_xlsx}")
    print(f"  rule-tagged signatures:       {rule_tagged}")
    print(f"  rerank-tagged signatures:     {rerank_tagged}")
    print(f"  LLM-tagged signatures:        {llm_tagged}")
    print(f"  needs_manual (LLM failed):    {needs_manual}  ← fill manual_horizontal for these")
    print(f"  low-confidence LLM (<0.7):    {low_conf}")
    print(f"  signatures w/ adjustments:    {adjustments}  (auto-flagged via has_negative)")
    ok_results = sum(1 for r in results if not is_err(r))
    err_results = sum(1 for r in results if is_err(r))
    heartbeat.finalize(done=len(results), ok=ok_results, err=err_results)

    # Generate / refresh the small feedback xlsx for the user to review
    write_feedback_template(merged, output_dir, company)


if __name__ == "__main__":
    main()
