"""Shared helpers for all pipeline nodes (progress logging, FORCE_REDO, folder resolution).

Each step's nodes.py imports `wrap_with_progress` and uses it as:

    from kpi.pipelines.common import wrap_with_progress
    from . import _logic

    def step0_convert(entity_key: str, _previous=None) -> str:
        return wrap_with_progress("step0_convert", entity_key, _logic.main)
"""
from __future__ import annotations

import json
import logging
import os
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Callable

ROOT = Path(__file__).resolve().parents[3]
log = logging.getLogger(__name__)


def _progress_path(folder: str) -> Path:
    p = ROOT / "data" / folder / "interim" / "_pipeline_progress.json"
    p.parent.mkdir(parents=True, exist_ok=True)
    return p


def _load_progress(p: Path) -> dict:
    if not p.exists():
        return {}
    try:
        with p.open("r", encoding="utf-8") as f:
            return json.load(f) or {}
    except Exception:
        return {}


def _save_progress(p: Path, data: dict) -> None:
    tmp = p.with_suffix(p.suffix + ".tmp")
    with tmp.open("w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, p)


def _iso(ts: float) -> str:
    return datetime.fromtimestamp(ts, tz=timezone.utc).isoformat()


def _force_redo_targets() -> set[str]:
    raw = os.environ.get("FORCE_REDO", "").strip()
    if not raw:
        return set()
    return {x.strip() for x in raw.split(",") if x.strip()}


def _clear_caches_for(module_name: str, folder: str) -> None:
    cache_dir = ROOT / "data" / folder / "interim" / "cache"
    if not cache_dir.exists():
        return
    deleted = []
    for f in cache_dir.iterdir():
        if module_name in f.name:
            f.unlink()
            deleted.append(f.name)
    if deleted:
        log.info("FORCE_REDO: cleared %s cache files", len(deleted))


def _resolve_folder(entity_key: str) -> str:
    from kpi.lib.conf import load_master, folder_for
    master = load_master()
    ent = (master.get("companies") or {}).get(entity_key) or {}
    return folder_for(entity_key, ent)


def wrap_with_progress(module_name: str, entity_key: str, logic_fn: Callable) -> str:
    """Set KPI_ENTITY, run logic_fn(), record progress JSON with timing."""
    os.environ["KPI_ENTITY"] = entity_key
    folder = _resolve_folder(entity_key)
    progress_p = _progress_path(folder)
    progress = _load_progress(progress_p)

    if module_name in _force_redo_targets() or "all" in _force_redo_targets():
        log.info("[FORCE_REDO] %s", module_name)
        _clear_caches_for(module_name, folder)
        progress.pop(module_name, None)
        _save_progress(progress_p, progress)

    started = time.time()
    progress[module_name] = {"status": "running", "started_at": _iso(started)}
    _save_progress(progress_p, progress)

    log.info("Running %s  (entity=%s / folder=%s)", module_name, entity_key, folder)
    try:
        logic_fn()
    except Exception as e:
        ended = time.time()
        progress[module_name] = {
            "status": "failed",
            "started_at": _iso(started),
            "ended_at": _iso(ended),
            "elapsed_seconds": round(ended - started, 2),
            "error": str(e)[:500],
        }
        _save_progress(progress_p, progress)
        raise
    ended = time.time()
    progress[module_name] = {
        "status": "done",
        "started_at": _iso(started),
        "ended_at": _iso(ended),
        "elapsed_seconds": round(ended - started, 2),
        "elapsed_minutes": round((ended - started) / 60, 2),
    }
    _save_progress(progress_p, progress)
    log.info("%s done in %.1fs", module_name, ended - started)
    return f"{module_name}_done"
