"""Step 2: LLM-tag each unique Project with a 縱向 (vertical) category.

Reads the parquet to compute per-project context (NG distribution, Capex/Opex split,
top accounts, top descriptions), then asks the LLM to pick one vertical from the
NG-constrained candidate list.

Output is written back to <company>_unique_projects.xlsx, populating:
  llm_vertical, llm_vertical_label, llm_confidence, llm_reasoning, llm_alt_candidates

Run after Step 0 + Step 1.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

import pandas as pd
from tqdm import tqdm

from kpi.lib.conf import ROOT, load_config, load_categories, path  # noqa: E402
from kpi.lib.heartbeat import Heartbeat  # noqa: E402
from kpi.lib.io_setup import force_unbuffered_io  # noqa: E402
from kpi.lib.llm import LLMClient  # noqa: E402

force_unbuffered_io()

TOP_N_ACCOUNTS = 8
TOP_N_DESCRIPTIONS = 8


def build_vertical_lookup(categories: dict) -> dict:
    return {v["id"]: v for v in categories["verticals"]}


def normalize_ng_code(ng: str | None, categories: dict) -> str | None:
    """Resolve ng to its canonical code ('NG0'..'NG11').

    Handles several formats seen across companies:
      - Direct code: "NG0" .. "NG11"
      - Chinese label: "娛樂表演" → NG3
      - "7.x - Chinese_Label" (MGM CAPEX): "7.3 - 娛樂表演" → NG3
      - "Chinese_Label - 活動/設施" (MGM WD/Patron): "娛樂表演 - 活動" → NG3
      - Internal whitespace: "吸引 外國客源" → "吸引外國客源" → NG1
      - Starts with "博彩": VML / MGM gaming entries → NG0
      - "Gaming" keyword → NG0
    Returns canonical code if resolved, else original value unchanged.
    """
    if not ng:
        return ng
    ng = ng.strip()
    ng_cats = categories.get("ng_categories", {})

    def _try_resolve(s: str) -> str | None:
        if s in ng_cats:
            return s
        for code, defn in ng_cats.items():
            if defn.get("label") == s:
                return code
        if s.startswith("博彩"):
            return "NG0"
        return None

    # Try direct
    r = _try_resolve(ng)
    if r:
        return r

    # Try with internal whitespace stripped (e.g. "吸引 外國客源" → "吸引外國客源")
    no_ws = "".join(ng.split())
    if no_ws != ng:
        r = _try_resolve(no_ws)
        if r:
            return r

    # Try splitting on " - " — try BOTH halves (handle "7.3 - 娛樂表演" and "娛樂表演 - 活動")
    if " - " in ng:
        for part in ng.split(" - "):
            part = part.strip()
            r = _try_resolve(part)
            if r:
                return r
            # Also try whitespace-stripped variant of each part
            part_no_ws = "".join(part.split())
            if part_no_ws != part:
                r = _try_resolve(part_no_ws)
                if r:
                    return r

    if ng.lower() == "gaming":
        return "NG0"
    return ng


def candidates_for_ng(ng: str, categories: dict) -> list[str]:
    """Eligible vertical IDs for a given NG, plus V_OTHER as fallback."""
    ng_def = categories.get("ng_categories", {}).get(ng)
    eligible = list(ng_def["eligible_verticals"]) if ng_def else []
    if "V_OTHER" not in eligible:
        eligible.append("V_OTHER")
    return eligible


def project_context_from_group(project: str, sub: pd.DataFrame, cols: dict) -> dict:
    ng_counts = sub[cols["ng11_category"]].value_counts(dropna=False).to_dict()
    co_counts = sub[cols["capex_opex"]].value_counts(dropna=False).to_dict()
    # primary_ng: pick most-common NON-EMPTY ng. Empty/NaN dominance was forcing
    # NG11 fallback → narrow candidate list → LLM choice rejected → V_OTHER.
    non_empty_ng = {
        k: v for k, v in ng_counts.items()
        if pd.notna(k) and str(k).strip() and str(k).strip().lower() != "nan"
    }
    primary_ng = max(non_empty_ng, key=non_empty_ng.get) if non_empty_ng else None
    top_accts = (
        sub.groupby(cols["account_desc"], dropna=False).size()
        .sort_values(ascending=False).head(TOP_N_ACCOUNTS).index.tolist()
    )
    top_descs = (
        sub.groupby(cols["description"], dropna=False).size()
        .sort_values(ascending=False).head(TOP_N_DESCRIPTIONS).index.tolist()
    )
    ctx: dict = {
        "project": project,
        "row_count": int(len(sub)),
        "total_amount": float(pd.to_numeric(sub[cols["amount"]], errors="coerce").sum()),
        "ng_counts": {str(k): int(v) for k, v in ng_counts.items()},
        "primary_ng": str(primary_ng) if primary_ng is not None else None,
        "capex_opex_counts": {str(k): int(v) for k, v in co_counts.items()},
        "top_account_descs": [str(x) for x in top_accts if pd.notna(x)],
        "top_descriptions": [str(x) for x in top_descs if pd.notna(x) and str(x).strip() not in ("", "-")],
    }
    # MGM prebuild adds source/section; include as classification hints when present.
    if "source" in sub.columns:
        src_counts = sub["source"].value_counts(dropna=True).head(3).to_dict()
        ctx["source_hints"] = {str(k): int(v) for k, v in src_counts.items()}
    if "section" in sub.columns:
        sec_counts = sub["section"].value_counts(dropna=True).head(3).to_dict()
        ctx["section_hints"] = {str(k): int(v) for k, v in sec_counts.items()}
    return ctx


SYSTEM_PROMPT = (
    "You are a domain expert classifier for Macau gaming concessionaires' NG0-NG11 "
    "non-gaming investment KPI reporting. Given a project's metadata, pick the single "
    "best 縱向 (vertical) category from a constrained candidate list.\n\n"
    "PRIORITY OF SIGNALS (most to least reliable):\n"
    "1. PROJECT NAME — most reliable. Look for specific keywords (Chinese AND English).\n"
    "2. Top descriptions — secondary signal.\n"
    "3. Top accounts — WEAKEST signal. Comp Lodging / Comp F&B appear in MANY project "
    "types (sponsorships, invites, MICE, etc.) so account types alone are misleading.\n\n"
    "KEY PATTERNS to recognize in PROJECT NAME:\n"
    "- 'BOH', 'Back of House', 'Cage Office', '賬房', 'Lobby', 'Atrium', 'WiFi/Wi-Fi', "
    "'Network', 'Server', 'Storage', 'Infrastructure', 'IT Hardware', 'Reception' "
    "→ V_PROPERTY_UPGRADE (NOT V_GAMING_EQUIP).\n"
    "- 'EG', 'Slot Machine', '角子機', 'Gaming Table', 'Smart Table', '娛樂桌', 'DICJ' "
    "→ V_GAMING_EQUIP.\n"
    "- 'Gaming Area', '博彩區', 'Casino Floor' → V_GAMING_VENUE.\n"
    "- 'Wellness', 'Spa', 'Gym', 'Fitness', 'Beauty', 'Health', '健康', '養生', 'Clinic', "
    "'Therapy' → V_WELLNESS.\n"
    "- 'Concert', '演唱會', 'Tour' (concert tour), 'Live Show', 'Live Music', 'Performance', "
    "'Magic Show', '魔術', 'Theatrical', 'Variety Show', '綜藝', 'Cirque', 'Acrobatic', "
    "'Stand-up Comedy', 'Dance Performance' → V_CONCERT (covers ALL artist/talent live "
    "performances, not just music).\n"
    "- 'WTT', 'ITTF', 'Marathon', 'Cup', 'Championship', 'Tournament', 'Olympics', "
    "'運動會', '賽事' → V_SPORT_EVENT.\n"
    "- 'Art', '藝術', 'Exhibition', '展覽', 'Gallery' → V_ART_EXHIBITION.\n"
    "- 'Museum', '博物館' → V_MUSEUM.\n"
    "- 'Restaurant', '餐廳', 'Dining', 'Café', 'Bistro' → V_RESTAURANT (the restaurant venue).\n"
    "- 'Food Festival', '美食節', 'Culinary', 'Wine Tasting', '美食活動' → V_FOOD_EVENT.\n"
    "- 'Donation', 'Charity', 'Tung Sin Tong', '同善堂', '公益', '社區', 'Community' "
    "→ V_COMMUNITY.\n"
    "- 'Cruise', 'Yacht', 'Marina', '海上', 'Ferry' → V_MARITIME.\n"
    "- 'Theme Park', '主題遊樂', 'Amusement', 'Kids', 'Family Entertainment' → V_THEME_PARK.\n"
    "- 'Roadshow', '路演', 'Trade Show' → V_OVERSEAS_ROADSHOW.\n"
    "- 'International Visitor', 'Privilege', 'Foreign Visitor', '外國客人' → V_INVITE_GUEST.\n"
    "- 'Travel Agency', 'OTA', '旅行社', 'Tour Operator' → V_INVITE_AGENCY.\n"
    "- 'Website', 'SEO', 'Search Engine', 'Digital Marketing', '海外網站' → V_OVERSEAS_WEB_SEO.\n"
    "- 'Promotional Video', '宣傳視頻', 'Videography', 'Video Shoot' → V_PROMO_VIDEO.\n"
    "- 'Overseas Office', '海外辦事處' → V_OVERSEAS_OFFICE.\n"
    "- 'Regional Team', '區域代表團隊', 'International Representative' → V_REGIONAL_TEAM.\n"
    "- 'Sales Rep', '區域銷售' → V_REGIONAL_SALES.\n"
    "- 'MICE', 'Conference', 'Convention', '會議展覽' → V_MICE.\n"
    "- 'Convention Center', 'Concert Hall', 'Stadium', 'Arena', 'Theatre/Theater', "
    "'會展中心', '演藝中心', '體育館' (the BUILDING itself) → V_VENUE_PERF_SPORT_MICE.\n\n"
    "CAPEX vs OPEX as a TIE-BREAKER (use when project name alone is ambiguous):\n"
    "- Capex-dominant (≥70% Capex) → favour construction/facility verticals: "
    "V_PROPERTY_UPGRADE, V_GAMING_VENUE, V_GAMING_EQUIP, V_RESTAURANT (new build), "
    "V_VENUE_PERF_SPORT_MICE (new venue), V_THEME_PARK (new attraction), V_MUSEUM.\n"
    "- Opex-dominant (≥70% Opex) → favour event/activity/operational verticals: "
    "V_CONCERT, V_SPORT_EVENT, V_FOOD_EVENT, V_ART_EXHIBITION, V_MICE, "
    "V_INVITE_GUEST, V_INVITE_AGENCY, V_OVERSEAS_ROADSHOW, V_COMMUNITY, V_MARITIME, "
    "V_OVERSEAS_WEB_SEO, V_PROMO_VIDEO, V_REGIONAL_SALES, V_WELLNESS (ongoing ops).\n"
    "- Mixed Capex+Opex → treat as Capex for facility sub-projects, Opex for running costs.\n\n"
    "AVOID V_OTHER unless the project genuinely doesn't fit ANY category. "
    "When uncertain between two specific candidates, pick the more specific one and "
    "list the alternative in alt_candidates.\n\n"
    "You MUST pick a vertical_id that appears in the candidate list. "
    "Always reply in valid JSON only."
)


def build_user_prompt(ctx: dict, candidates: list[dict], ng_label: str) -> str:
    lines = [
        f"Project name: {ctx['project']}",
        f"Primary NG category: {ctx['primary_ng']} ({ng_label})",
        f"Row count: {ctx['row_count']:,}   Total amount (MOP): {ctx['total_amount']:,.0f}",
        f"NG distribution: {ctx['ng_counts']}",
        f"Capex/Opex distribution: {ctx['capex_opex_counts']}",
    ]
    if ctx.get("source_hints"):
        lines.append(f"Data sources (type of spend record): {ctx['source_hints']}")
    if ctx.get("section_hints"):
        lines.append(f"Section / NG subsection codes: {ctx['section_hints']}")
    lines += [
        "",
        "Top account descriptions (most frequent):",
    ]
    for a in ctx["top_account_descs"]:
        lines.append(f"  - {a}")
    lines.append("")
    lines.append("Top description samples:")
    for d in ctx["top_descriptions"]:
        lines.append(f"  - {d}")
    lines.append("")
    lines.append("Candidate verticals (pick exactly ONE id from this list):")
    for c in candidates:
        lines.append(f"  - id: {c['id']}   label: {c['label']}")
    lines.append("")
    lines.append(
        "Reply with JSON only, schema:\n"
        "{\n"
        '  "vertical_id": "<one of candidate ids>",\n'
        '  "confidence": <float 0..1>,\n'
        '  "reasoning": "<one sentence in Chinese explaining the choice>",\n'
        '  "alt_candidates": [<other plausible candidate ids, ranked, may be empty>]\n'
        "}"
    )
    return "\n".join(lines)


def main():
    cfg = load_config()
    cats = load_categories()
    company = cfg["company"]["code"]
    interim = path(cfg, "interim_dir")
    cols = cfg["columns"]

    parquet_path = interim / f"{company}_raw.parquet"
    proj_xlsx = interim / f"{company}_unique_projects.xlsx"
    if not parquet_path.exists() or not proj_xlsx.exists():
        raise FileNotFoundError("Run Step 0 + Step 1 first.")

    proj_df = pd.read_excel(proj_xlsx)

    # Skip re-running the LLM if this file was already tagged on a previous run (V tags live in
    # llm_vertical). A plain `kedro run --pipeline=<ent>` would otherwise re-do every project's LLM
    # call (and load the big raw parquet, which OOM'd). `--from-nodes step4` can't be used because
    # the inter-node datasets are MemoryDatasets. Set KPI_FORCE_STEP2=1 (or delete the xlsx) to re-tag.
    import os as _os
    _tagcol = next((c for c in ("llm_vertical", "manual_vertical", "vertical_id") if c in proj_df.columns), None)
    if (_tagcol and not _os.environ.get("KPI_FORCE_STEP2", "").strip()
            and proj_df[_tagcol].astype("string").fillna("").str.strip().ne("").mean() > 0.5):
        print(f"  [skip] {proj_xlsx.name} already tagged ({_tagcol} populated) — "
              f"set KPI_FORCE_STEP2=1 to re-run the LLM", flush=True)
        return

    # Conf flag `skip_vertical_llm: true` — for entities whose V is fully driven by a project-team
    # column_map applied at step4 (e.g. wynn 項目分類2, mgm Project_code). Don't load the big raw
    # parquet or run the project LLM at all: default every untagged project to V_OTHER here; step4's
    # row_vertical_overrides column_map sets the real V. Mirrors step3 skip_signature_llm. Manual
    # tags (manual_vertical) are untouched and still win at step4.
    _ent = ((cfg.get("_master") or {}).get("companies") or {}).get(company, {}) or {}
    if bool(_ent.get("skip_vertical_llm", False)):
        _dft = _ent.get("vertical_default", "V_OTHER")
        if "llm_vertical" not in proj_df.columns:
            proj_df["llm_vertical"] = ""
        _blank = proj_df["llm_vertical"].astype("string").fillna("").str.strip().eq("")
        proj_df.loc[_blank, "llm_vertical"] = _dft
        proj_df["llm_status"] = "skip_llm"
        proj_df["tag_source"] = "skip_llm"
        proj_df.to_excel(proj_xlsx, index=False)
        print(f"  [skip-llm] vertical LLM disabled — {int(_blank.sum()):,} projects defaulted to {_dft} "
              f"(step4 column_map / row_vertical_overrides set final V); raw parquet NOT loaded", flush=True)
        return

    df = pd.read_parquet(parquet_path)

    vlookup = build_vertical_lookup(cats)
    ng_categories = cats.get("ng_categories", {})

    cache_path = interim / "cache" / f"{company}_step2_projects.jsonl"
    llm = LLMClient(cfg, cache_path=cache_path)

    project_col = cols["project"]
    projects = proj_df[project_col].astype("string").fillna("").tolist()

    print(f"Tagging {len(projects)} unique projects.")

    df = df.assign(_proj=df[project_col].astype("string").fillna(""))
    groups = dict(tuple(df.groupby("_proj", sort=False)))

    payloads_all = []
    for proj in tqdm(projects, desc="step2 build context"):
        sub = groups.get(proj)
        if sub is None or len(sub) == 0:
            continue
        ctx = project_context_from_group(proj, sub, cols)
        primary_ng = normalize_ng_code(ctx["primary_ng"], cats) or "NG11"
        ctx["primary_ng"] = primary_ng  # store canonical code
        cand_ids = candidates_for_ng(primary_ng, cats)
        candidates = [{"id": cid, "label": vlookup[cid]["label"]} for cid in cand_ids if cid in vlookup]
        ng_label = ng_categories.get(primary_ng, {}).get("label", "")
        payloads_all.append({
            "project": proj,
            "primary_ng": primary_ng,
            "ctx": ctx,
            "candidates": candidates,
            "candidate_ids": cand_ids,
            "ng_label": ng_label,
        })

    # Predominant rules from parameters.yml — applied BEFORE LLM. First-matching rule wins.
    # Rule's `then` must be in the candidates list for that project's NG (else fall through to LLM).
    from kpi.lib.rules import load_predominant_rules, find_predominant_match
    pr_rules = load_predominant_rules(cfg, kind="vertical")

    pre_results: list[dict | None] = [None] * len(payloads_all)
    rule_hits = 0
    if pr_rules:
        for i, p in enumerate(payloads_all):
            # Build a flat item for rule matching from the project context
            item = {
                "project": p["project"],
                "vendor": "",      # project-level — no single vendor
                "capex_opex": (max(p["ctx"]["capex_opex_counts"], key=p["ctx"]["capex_opex_counts"].get)
                                if p["ctx"].get("capex_opex_counts") else ""),
                "desc": " ".join(p["ctx"].get("top_descriptions", [])),
                "account_desc": " ".join(p["ctx"].get("top_account_descs", [])),
            }
            target, idx = find_predominant_match(pr_rules, item)
            if target and target in {c["id"] for c in p["candidates"]}:
                pre_results[i] = {
                    "vertical_id": target,
                    "confidence": 0.99,
                    "reasoning": f"Predominant rule #{idx} ({pr_rules[idx].get('if')!r})",
                    "alt_candidates": [],
                    "_tag_source": f"rule:{idx}",
                }
                rule_hits += 1
        print(f"  [rules] {rule_hits} projects resolved by predominant rules ({len(pr_rules)} rules configured)")

    payloads = [p for p, r in zip(payloads_all, pre_results) if r is None]
    print(f"  LLM workload: {len(payloads)} projects")

    heartbeat = Heartbeat(
        path=interim / f"{company}_step2_progress.json",
        step_name="step2_projects",
        total=len(payloads),
    )
    llm.attach_heartbeat(heartbeat)

    def call_one(p):
        result = llm.chat_json(
            system=SYSTEM_PROMPT,
            user=build_user_prompt(p["ctx"], p["candidates"], p["ng_label"]),
        )
        cand_ids = {c["id"] for c in p["candidates"]}
        chosen = result.get("vertical_id")
        if chosen not in cand_ids:
            result["_invalid_choice"] = chosen
            result["vertical_id"] = "V_OTHER"
            result["confidence"] = min(float(result.get("confidence", 0.0)), 0.3)
        result["_tag_source"] = "llm"
        return result

    if not payloads:
        llm_results: list = []
    else:
        llm_results = llm.map_parallel(payloads, call_one, desc="step2 projects")

    # Splice pre-tagged + LLM results back into full result array.
    full_results: list[dict] = []
    llm_iter = iter(llm_results)
    for r in pre_results:
        if r is not None:
            full_results.append(r)
        else:
            try:
                full_results.append(next(llm_iter))
            except StopIteration:
                full_results.append({"_error": "missing LLM result"})
    results = full_results
    payloads = payloads_all  # for downstream zip

    def is_err(r: dict) -> bool:
        return isinstance(r, dict) and "_error" in r

    out_rows = []
    for p, r in zip(payloads, results):
        if is_err(r):
            out_rows.append({
                project_col: p["project"],
                "primary_ng": p["primary_ng"],
                "llm_status": "needs_manual",
                "tag_source": "",
                "llm_vertical": "",
                "llm_vertical_label": "",
                "llm_capex_opex": "",
                "llm_confidence": None,
                "llm_reasoning": "",
                "llm_alt_candidates": "",
                "llm_invalid_choice": "",
                "llm_error": r.get("_error", ""),
            })
            continue
        vid = r.get("vertical_id", "V_OTHER")
        out_rows.append({
            project_col: p["project"],
            "primary_ng": p["primary_ng"],
            "llm_status": "ok",
            "tag_source": r.get("_tag_source", "llm"),
            "llm_vertical": vid,
            "llm_vertical_label": vlookup.get(vid, {}).get("label", ""),
            "llm_capex_opex": "",
            "llm_confidence": r.get("confidence"),
            "llm_reasoning": r.get("reasoning"),
            "llm_alt_candidates": json.dumps(r.get("alt_candidates", []), ensure_ascii=False),
            "llm_invalid_choice": r.get("_invalid_choice", ""),
            "llm_error": "",
        })
    enrich = pd.DataFrame(out_rows)

    drop_cols = [c for c in enrich.columns if c != project_col and c in proj_df.columns]
    proj_df = proj_df.drop(columns=drop_cols, errors="ignore")
    merged = proj_df.merge(enrich, on=project_col, how="left")
    merged.to_excel(proj_xlsx, index=False)

    needs_manual = (merged["llm_status"] == "needs_manual").sum()
    rule_tagged = merged["tag_source"].fillna("").astype(str).str.startswith("rule:").sum()
    rerank_tagged = merged["tag_source"].fillna("").astype(str).str.startswith("rerank:").sum()
    llm_tagged = (merged["tag_source"] == "llm").sum()
    other_count = (merged["llm_vertical"] == "V_OTHER").sum()
    other_pct = (other_count / max(len(merged), 1)) * 100
    low_conf = ((merged["tag_source"] == "llm") & (merged["llm_confidence"] < 0.7)).sum()
    invalid = (merged["llm_invalid_choice"].fillna("") != "").sum()
    print(f"Done. Wrote {proj_xlsx}")
    print(f"  rule-tagged:                {rule_tagged}")
    print(f"  rerank-tagged:              {rerank_tagged}")
    print(f"  LLM-tagged:                 {llm_tagged}")
    print(f"  needs_manual (LLM failed):  {needs_manual}  ← fill manual_vertical for these")
    print(f"  low-confidence LLM (<0.7):  {low_conf}")
    print(f"  invalid choices to V_OTHER: {invalid}")
    print(f"  V_OTHER count:              {other_count} ({other_pct:.1f}%)")
    ok_results = sum(1 for r in results if not is_err(r))
    err_results = sum(1 for r in results if is_err(r))
    heartbeat.finalize(done=len(results), ok=ok_results, err=err_results)


if __name__ == "__main__":
    main()
