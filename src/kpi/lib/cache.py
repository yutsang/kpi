"""Disk-based jsonl cache for idempotent LLM calls.

Crash-safe: each (key, value) is appended on success, so re-running picks up
where you left off without redoing work.
"""
from __future__ import annotations

import hashlib
import json
import threading
from pathlib import Path
from typing import Any


class JsonlCache:
    def __init__(self, path: Path):
        self.path = Path(path)
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._lock = threading.Lock()
        self._mem: dict[str, Any] = {}
        if self.path.exists():
            with self.path.open("r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if not line:
                        continue
                    try:
                        obj = json.loads(line)
                        self._mem[obj["key"]] = obj["value"]
                    except Exception:
                        continue

    @staticmethod
    def make_key(*parts: Any) -> str:
        h = hashlib.sha256()
        for p in parts:
            h.update(repr(p).encode("utf-8"))
        return h.hexdigest()[:40]

    def get(self, key: str):
        return self._mem.get(key)

    def put(self, key: str, value: Any) -> None:
        with self._lock:
            if key in self._mem:
                return
            self._mem[key] = value
            with self.path.open("a", encoding="utf-8") as f:
                f.write(json.dumps({"key": key, "value": value}, ensure_ascii=False) + "\n")

    def __len__(self) -> int:
        return len(self._mem)
