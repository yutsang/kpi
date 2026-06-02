"""Rule engines for both step 2 (project → vertical) and step 3 (account → horizontal).

Both classes share the same matching primitives (prefix / contains / equals / regex).
Designed for high-confidence cases only — anything unmapped falls to rerank/LLM.
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import Iterable

import yaml


class AccountRules:
    """Rule engine over a row's signature fields. Supports two formats:

    Simple (matches account_desc only):
      - {id: NAME, match: prefix, value: "AUC-", horizontal_id: H_CONSTRUCTION}

    Compound (multi-field with exclusion):
      - id: AUC_BUT_NOT_LABOR
        horizontal_id: H_CONSTRUCTION
        if_any:
          - {field: account_desc, match: prefix, value: "AUC-"}
        exclude_if_any:
          - {field: description, match: contains, value: "staff"}
          - {field: description, match: contains, value: "payroll"}

    Compound semantics:
      - if_all   → all clauses must match
      - if_any   → at least one clause must match
      - exclude_if_any → if ANY of these match, the rule is skipped
      - exclude_if_all → if ALL of these match, the rule is skipped
      - field defaults to "account_desc" if omitted; valid: account_code, account_desc,
        desc_norm, description, project, etc. (whatever's in the sig dict).
    """

    def __init__(
        self,
        rules: list[dict],
        account_code_overrides: dict[str, str] | None = None,
        row_type_defaults: dict | None = None,
    ):
        self.rules = rules
        self.account_code_overrides = {
            str(k).strip(): str(v).strip() for k, v in (account_code_overrides or {}).items() if k and v
        }
        self.row_type_defaults = row_type_defaults or {}

    @classmethod
    def load(cls, root: Path, company_code: str | None = None) -> "AccountRules":
        """Loads:
          - UNIVERSAL pattern rules from config/account_horizontal_rules.yaml
          - COMPANY-SPECIFIC overrides from config/account_overrides_<company_code>.yaml
        Pattern rules apply across companies (SAP standard names);
        account_code overrides are per-company since chart-of-accounts differs."""
        import sys
        rules: list = []
        overrides: dict = {}
        row_type_defaults: dict = {}

        universal_path = root / "config" / "account_horizontal_rules.yaml"
        if universal_path.exists():
            with universal_path.open("r", encoding="utf-8") as f:
                doc = yaml.safe_load(f) or {}
            rules = doc.get("rules") or []
            row_type_defaults = doc.get("row_type_defaults") or {}
            # account_code_overrides may also live here for backwards compat
            overrides.update(doc.get("account_code_overrides") or {})
        else:
            print(
                f"WARNING: {universal_path} not found — pattern rules disabled.",
                file=sys.stderr,
            )

        if company_code:
            company_path = root / "config" / f"account_overrides_{company_code}.yaml"
            if company_path.exists():
                with company_path.open("r", encoding="utf-8") as f:
                    cdoc = yaml.safe_load(f) or {}
                co = cdoc.get("account_code_overrides") or {}
                # Per-company overrides take precedence over any universal defaults.
                overrides.update(co)
                print(
                    f"  loaded {len(co)} company-specific account overrides from "
                    f"{company_path.name}",
                    file=sys.stderr,
                )
            else:
                print(
                    f"  no company-specific overrides at {company_path.name} "
                    f"(file is optional; rerank/LLM will handle company-specific accounts)",
                    file=sys.stderr,
                )

        return cls(
            rules=rules,
            account_code_overrides=overrides,
            row_type_defaults=row_type_defaults,
        )

    def apply(self, sig: dict | str) -> tuple[str | None, str | None]:
        """Resolution order:
          1. Compound rules in `rules:` (LABOR_OVERRIDE_DESC etc fire first regardless of account)
          2. account_code_overrides (deterministic per-code lookup)
        Returns (horizontal_id, rule_id) on the first match, else (None, None).
        """
        if isinstance(sig, str):
            sig = {"account_desc": sig}
        # Tier 1: pattern rules
        for rule in self.rules:
            if _rule_matches(rule, sig):
                return (rule.get("horizontal_id"), rule.get("id") or rule.get("match", ""))
        # Tier 2: account_code direct overrides
        code = str(sig.get("account_code") or "").strip()
        if code and code in self.account_code_overrides:
            return (self.account_code_overrides[code], f"acct_code:{code}")
        return (None, None)

    def default_row_type(self, has_negative: bool) -> str:
        key = "has_negative_true" if has_negative else "has_negative_false"
        return str(self.row_type_defaults.get(key, "normal"))


def _rule_matches(rule: dict, sig: dict) -> bool:
    """Evaluate a single rule against a signature dict, supporting both simple and
    compound rule formats."""
    # Simple format: {match: ..., value: ..., horizontal_id: ...}
    is_simple = "match" in rule and "value" in rule and "if_any" not in rule and "if_all" not in rule
    if is_simple:
        field = rule.get("field", "account_desc")
        text = str(sig.get(field) or "")
        if not _match(rule, text):
            return False
        # simple format may still have exclude_if_any
    else:
        # Compound: must satisfy if_all (all) and if_any (at least one)
        if_all = rule.get("if_all") or []
        if if_all and not all(_clause_matches(c, sig) for c in if_all):
            return False
        if_any = rule.get("if_any") or []
        if if_any and not any(_clause_matches(c, sig) for c in if_any):
            return False
        if not if_all and not if_any:
            return False  # empty compound — never matches

    # Negative conditions (exclude)
    excl_any = rule.get("exclude_if_any") or []
    if excl_any and any(_clause_matches(c, sig) for c in excl_any):
        return False
    excl_all = rule.get("exclude_if_all") or []
    if excl_all and all(_clause_matches(c, sig) for c in excl_all):
        return False
    return True


def _clause_matches(clause: dict, sig: dict) -> bool:
    field = clause.get("field", "account_desc")
    text = str(sig.get(field) or "")
    return _match(clause, text)


def _match(rule: dict, text: str) -> bool:
    kind = rule.get("match", "").lower()
    value = str(rule.get("value", ""))
    if not value:
        return False
    lower = text.lower()
    v = value.lower()
    if kind == "prefix":
        return lower.startswith(v)
    if kind == "contains":
        return v in lower
    if kind == "equals":
        return lower == v
    if kind == "regex":
        try:
            return bool(re.search(value, text, re.IGNORECASE))
        except re.error:
            return False
    return False


# ─────────────────────────────────────────────────────────────────────────────
# Predominant rules — loaded from parameters.yml `predominant_rules` section.
# Lightweight, applied BEFORE LLM in step2 / step3 to short-circuit obvious cases.
# ─────────────────────────────────────────────────────────────────────────────

import logging as _logging

_pr_log = _logging.getLogger(__name__)


def _to_list(v) -> list[str]:
    if v is None:
        return []
    if isinstance(v, list):
        return [str(x) for x in v]
    return [str(v)]


def _icontains(haystack: str, needles: list[str]) -> bool:
    hl = haystack.lower()
    return any(n and n.lower() in hl for n in needles)


def _predominant_match_one(rule: dict, item: dict) -> bool:
    """Match a single rule's `if` clause (all conditions AND-joined)."""
    conds = rule.get("if") or {}
    if not conds:
        return False
    for ck, cv in conds.items():
        if ck == "account_code_prefix":
            code = str(item.get("account_code", ""))
            prefixes = _to_list(cv)
            if not prefixes or not any(code.startswith(p) for p in prefixes):
                return False
        elif ck == "account_code_equals":
            code = str(item.get("account_code", "")).strip()
            values = [str(v).strip() for v in _to_list(cv)]
            if not values or code not in values:
                return False
        elif ck == "account_desc_equals":
            desc = str(item.get("account_desc", "")).strip().lower()
            values = [str(v).strip().lower() for v in _to_list(cv)]
            if not values or desc not in values:
                return False
        elif ck == "account_desc_prefix":
            desc = str(item.get("account_desc", "")).lower()
            prefixes = [str(p).lower() for p in _to_list(cv)]
            if not prefixes or not any(desc.startswith(p) for p in prefixes):
                return False
        elif ck == "account_desc_contains":
            if not _icontains(str(item.get("account_desc", "")), _to_list(cv)):
                return False
        elif ck == "desc_contains":
            blob = " ".join([
                str(item.get("desc", "")),
                str(item.get("desc_norm", "")),
                str(item.get("description", "")),
                str(item.get("desc_samples", "")),
            ])
            if not _icontains(blob, _to_list(cv)):
                return False
        elif ck == "project_contains":
            # Sig-level dict has 'project_samples' (aggregated) but rarely 'project'.
            # Check both: row-level 'project' OR sig-level 'project_samples' aggregated string.
            blob = " ".join([
                str(item.get("project", "")),
                str(item.get("project_samples", "")),
            ])
            if not _icontains(blob, _to_list(cv)):
                return False
        elif ck == "vendor_contains":
            if not _icontains(str(item.get("vendor", "")), _to_list(cv)):
                return False
        elif ck == "job_code_contains":
            blob = " ".join([
                str(item.get("job_code", "")),
                str(item.get("job_code_samples", "")),
            ])
            if not _icontains(blob, _to_list(cv)):
                return False
        elif ck == "capex_opex_equals":
            if str(item.get("capex_opex", "")).strip().upper() != str(cv).strip().upper():
                return False
        else:
            # ── Generic fallback ────────────────────────────────────────────────
            # Supports arbitrary user-defined fields surfaced via step1's auto-detect.
            # Predicate naming convention: <field>_equals / _contains / _prefix / _not_empty
            # Looks up both <field> (single-row dict) and <field>_samples (aggregated sig dict).
            handled = False
            EMPTY_TOKENS = {"", "0", "0.0", "(blanks)", "nan", "none", "n/a", "na"}

            def _field_blob(field_name: str) -> str:
                a = str(item.get(field_name, "") or "")
                b = str(item.get(f"{field_name}_samples", "") or "")
                return f"{a} || {b}".strip()

            def _split_tokens(blob: str) -> list[str]:
                return [t.strip() for t in blob.split("||") if t.strip()]

            for suffix in ("_equals", "_contains", "_prefix", "_not_empty"):
                if not ck.endswith(suffix):
                    continue
                field = ck[: -len(suffix)]
                blob = _field_blob(field)

                if suffix == "_equals":
                    tokens = [t.lower() for t in _split_tokens(blob)]
                    targets = [str(v).strip().lower() for v in _to_list(cv)]
                    if not targets or not any(t in targets for t in tokens):
                        return False
                elif suffix == "_contains":
                    if not _icontains(blob, _to_list(cv)):
                        return False
                elif suffix == "_prefix":
                    tokens = [t.lower() for t in _split_tokens(blob)]
                    prefixes = [str(p).lower() for p in _to_list(cv)]
                    if not prefixes or not any(any(t.startswith(p) for p in prefixes) for t in tokens):
                        return False
                elif suffix == "_not_empty":
                    tokens = [t.lower() for t in _split_tokens(blob)]
                    non_empty = [t for t in tokens if t not in EMPTY_TOKENS]
                    want_non_empty = bool(cv)
                    if want_non_empty and not non_empty:
                        return False
                    if not want_non_empty and non_empty:
                        return False

                handled = True
                break

            if not handled:
                _pr_log.warning("predominant rule: unknown condition '%s' — rule skipped", ck)
                return False
    return True


def find_predominant_match(rules: list[dict], item: dict) -> tuple[str | None, int]:
    """Legacy: return (rule['then'] value, rule_index) for first matching rule.
    Use find_predominant_rule for richer rules (then_scope, then_any_of)."""
    for i, rule in enumerate(rules or []):
        if _predominant_match_one(rule, item):
            return rule.get("then"), i
    return None, -1


def find_predominant_rule(rules: list[dict], item: dict) -> tuple[dict | None, int]:
    """Return (matched_rule_dict, rule_index) for first matching rule, or (None, -1)."""
    for i, rule in enumerate(rules or []):
        if _predominant_match_one(rule, item):
            return rule, i
    return None, -1


def resolve_then(rule: dict, all_horizontals: list[dict]) -> dict:
    """Inspect a matched rule's target. Returns one of:
      {"mode": "fixed",   "id": "<H_ID>"}                 — single ID, skip LLM
      {"mode": "subset",  "ids": ["<H_A>", "<H_B>", ...]}  — narrow LLM candidate list
      {"mode": "invalid"}                                   — rule has no valid target
    """
    if "then" in rule:
        return {"mode": "fixed", "id": rule["then"]}
    if "then_scope" in rule:
        scope = str(rule["then_scope"]).lower().strip()
        ids = [
            h["id"] for h in all_horizontals
            if h.get("id") != "H_COUNT" and str(h.get("scope", "")).lower() == scope
        ]
        return {"mode": "subset", "ids": ids}
    if "then_any_of" in rule:
        ids = list(rule["then_any_of"]) if isinstance(rule["then_any_of"], list) else [rule["then_any_of"]]
        return {"mode": "subset", "ids": [str(x) for x in ids]}
    return {"mode": "invalid"}


def load_predominant_rules(cfg: dict, kind: str) -> list[dict]:
    """Load rules of `kind` ('horizontal' or 'vertical') from entity parameters.yml.
    Returns [] if not configured."""
    master = cfg.get("_master") or {}
    entity_key = cfg.get("_entity") or cfg.get("company", {}).get("code")
    ent = (master.get("companies") or {}).get(entity_key) or {}
    pr = ent.get("predominant_rules") or {}
    return list(pr.get(kind) or [])


class ProjectRules:
    """Project_name → vertical_id, constrained to NG-eligible verticals.

    Eligibility comes from categories.yaml ng_categories[NG].eligible_verticals.
    Within that NG, verticals are checked in eligible-list order; the first whose
    pattern matches wins. This means more specific verticals should appear earlier
    in the eligible list (e.g. V_GAMING_VENUE before V_PROPERTY_UPGRADE in NG0).
    """

    def __init__(self, vertical_rules: dict[str, list[dict]]):
        self.vertical_rules = vertical_rules

    @classmethod
    def load(cls, root: Path) -> "ProjectRules":
        import sys
        path = root / "config" / "project_vertical_rules.yaml"
        if not path.exists():
            print(
                f"WARNING: {path} not found — project rule pre-tagging disabled. "
                "Step 2 will rely entirely on rerank+LLM for vertical classification.",
                file=sys.stderr,
            )
            return cls(vertical_rules={})
        with path.open("r", encoding="utf-8") as f:
            doc = yaml.safe_load(f) or {}
        return cls(vertical_rules=doc.get("vertical_rules") or {})

    def apply(
        self,
        project_name: str,
        eligible_verticals: list[str],
    ) -> tuple[str | None, str | None]:
        text = str(project_name or "").strip()
        if not text:
            return (None, None)
        for vid in eligible_verticals:
            rules = self.vertical_rules.get(vid) or []
            for rule in rules:
                if _match(rule, text):
                    sig = rule.get("value") or rule.get("match", "")
                    return (vid, f"{vid}:{sig}")
        return (None, None)
