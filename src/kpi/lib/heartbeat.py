"""Heartbeat: periodically write progress JSON so an external 'cat' can monitor a long run.

Atomic write (temp file + rename) so the file is never half-written.
"""
from __future__ import annotations

import json
import os
import threading
import time
from datetime import datetime, timezone
from pathlib import Path


class Heartbeat:
    def __init__(self, path: Path, step_name: str, total: int):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.step_name = step_name
        self.total = total
        self.start_time = time.perf_counter()
        self._lock = threading.Lock()
        self._last_write = 0.0
        self.update(done=0, ok=0, err=0, force=True)

    def update(
        self,
        *,
        done: int,
        ok: int,
        err: int,
        extra: dict | None = None,
        force: bool = False,
    ) -> None:
        now = time.perf_counter()
        if not force and (now - self._last_write) < 1.0:
            return
        with self._lock:
            self._last_write = now
            elapsed = now - self.start_time
            rate = done / elapsed if elapsed > 0 else 0.0
            remaining = max(self.total - done, 0)
            eta_seconds = remaining / rate if rate > 0 else None
            payload = {
                "step": self.step_name,
                "timestamp": datetime.now(timezone.utc).isoformat(),
                "total": self.total,
                "done": done,
                "left": remaining,
                "ok": ok,
                "err": err,
                "pct": round((done / self.total) * 100, 2) if self.total else 0,
                "elapsed_seconds": round(elapsed, 1),
                "rate_per_second": round(rate, 3),
                "eta_seconds": round(eta_seconds, 1) if eta_seconds is not None else None,
                "eta_minutes": round(eta_seconds / 60, 1) if eta_seconds is not None else None,
                "eta_hours": round(eta_seconds / 3600, 2) if eta_seconds is not None else None,
            }
            if extra:
                payload.update(extra)
            tmp = self.path.with_suffix(self.path.suffix + ".tmp")
            with tmp.open("w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)
            os.replace(tmp, self.path)

    def finalize(self, *, done: int, ok: int, err: int) -> None:
        self.update(done=done, ok=ok, err=err, extra={"finished": True}, force=True)
