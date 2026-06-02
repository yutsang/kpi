"""Master config loader (kedro envs layout).

Layout:
  conf/base/parameters.yml         — COMMON defaults (concurrency, batch_size, etc.)
  conf/base/categories.yml         — COMMON Macau NG taxonomy
  conf/<entity>/parameters.yml     — SPECIFIC per-entity (alias, columns, filename)
  conf/local/credentials.yml       — LLM api credentials (gitignored)

Selecting an entity:
  KPI_ENTITY=company_1 python ...
  python ... --entity company_1
  python ... --entity galaxy        # if conf/company_1/parameters.yml has alias='galaxy'
"""
from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

import yaml

ROOT = Path(__file__).resolve().parents[3]
CONF_DIR = ROOT / "conf"
RESERVED_ENV_DIRS = {"base", "local"}


def _read_yaml(p: Path) -> dict:
    if not p.exists():
        return {}
    with p.open("r", encoding="utf-8-sig") as f:  # utf-8-sig strips Windows BOM
        return yaml.safe_load(f) or {}


def _scan_companies() -> dict[str, dict]:
    """Find every conf/<dir>/parameters.yml and treat each as one company."""
    out: dict[str, dict] = {}
    if not CONF_DIR.exists():
        return out
    for d in sorted(CONF_DIR.iterdir()):
        if not d.is_dir() or d.name in RESERVED_ENV_DIRS:
            continue
        yml = d / "parameters.yml"
        try:
            params = _read_yaml(yml)
            if params:
                out[d.name] = params
        except Exception as e:
            print(f"[conf] WARNING: skipping {yml} — {e}", flush=True)
    return out


def load_master() -> dict:
    """Merge base defaults + credentials + auto-discovered per-entity configs."""
    base = _read_yaml(CONF_DIR / "base" / "parameters.yml")
    creds = _read_yaml(CONF_DIR / "local" / "credentials.yml")
    local_overrides = _read_yaml(CONF_DIR / "local" / "parameters.yml")
    companies = _scan_companies()

    master = {
        "defaults": base.get("defaults") or {},
        "llm": creds.get("llm") or {},
        "companies": companies,
    }
    for k, v in local_overrides.items():
        if isinstance(v, dict) and isinstance(master.get(k), dict):
            master[k] = {**master[k], **v}
        else:
            master[k] = v
    return master


def _resolve_entity_arg() -> str:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--entity", default=None)
    parser.add_argument("--config", default=None)
    args, _ = parser.parse_known_args(sys.argv[1:])
    if args.entity:
        return args.entity
    env = os.environ.get("KPI_ENTITY")
    if env:
        return env
    raise ValueError("No entity selected. Set KPI_ENTITY or pass --entity.")


def resolve_entity_key(master: dict, entity_or_alias: str) -> str:
    companies = master.get("companies") or {}
    if entity_or_alias in companies:
        return entity_or_alias
    for k, v in companies.items():
        if v and v.get("alias") == entity_or_alias:
            return k
    raise KeyError(
        f"'{entity_or_alias}' is neither a company key nor an alias. "
        f"Known keys: {list(companies)}"
    )


def folder_for(entity_key: str, ent_cfg: dict) -> str:
    return (ent_cfg.get("alias") or entity_key).strip()


def load_config(entity_or_alias: str | None = None) -> dict:
    master = load_master()
    if entity_or_alias is None:
        entity_or_alias = _resolve_entity_arg()
    entity_key = resolve_entity_key(master, entity_or_alias)
    ent = (master.get("companies") or {}).get(entity_key, {})
    if not ent.get("name") or not ent.get("raw_filename"):
        raise ValueError(
            f"conf/{entity_key}/parameters.yml missing name or raw_filename. "
            f"Edit it or copy from conf/{entity_key}/parameters.yml.template."
        )
    folder = folder_for(entity_key, ent)
    cfg = {
        "company": {
            "code": entity_key,
            "name": ent.get("name", ""),
            "alias": ent.get("alias", ""),
            "folder": folder,
        },
        "llm": {**(master.get("defaults") or {}), **(master.get("llm") or {})},
        "rerank": {"enabled": False},
        "columns": ent.get("columns", {}) or {},
        "paths": {
            "base_dir": f"data/{folder}",
            "raw_filename": ent["raw_filename"],
            "raw_sheet": ent.get("raw_sheet", 0),
        },
        "header_row": ent.get("header_row", 0),
        "_root": ROOT,
        "_master": master,
        "_entity": entity_key,
    }
    print(
        f"Loaded entity={entity_key}"
        + (f" (alias={ent.get('alias')})" if ent.get("alias") else "")
        + f" name={ent.get('name')}",
        file=sys.stderr,
    )
    return cfg


def load_categories() -> dict:
    p = CONF_DIR / "base" / "categories.yml"
    if not p.exists():
        return {"verticals": [], "horizontals": [], "ng_categories": {}, "row_types": []}
    return _read_yaml(p)


def path(cfg: dict, key: str) -> Path:
    paths = cfg.get("paths", {})
    base = paths.get("base_dir")
    derived = {
        "raw_xlsx":     f"{base}/raw/{paths.get('raw_filename', cfg['company']['code'] + '.xlsx')}",
        "interim_dir":  f"{base}/interim",
        "output_dir":   f"{base}/output",
        "cache_dir":    f"{base}/interim/cache",
        "mappings_dir": f"{base}/mappings",
    }
    if key in derived:
        p = ROOT / derived[key]
    else:
        p = ROOT / paths[key]
    if key.endswith("_dir"):
        p.mkdir(parents=True, exist_ok=True)
    return p
