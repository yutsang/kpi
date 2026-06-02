"""Thread-safe leaky-bucket rate limiter.

Use to keep client request rate just under the server's per-second limit so concurrent
threads don't trigger 'REQUEST TOO FREQUENT' (code=30001) responses.
"""
from __future__ import annotations

import threading
import time


class RateLimiter:
    def __init__(self, max_per_second: float):
        if max_per_second is None or max_per_second <= 0:
            self.disabled = True
            self.interval = 0.0
            self._next_slot = 0.0
            self._lock = threading.Lock()
            return
        self.disabled = False
        self.interval = 1.0 / float(max_per_second)
        self._next_slot = time.perf_counter()
        self._lock = threading.Lock()

    def acquire(self, stop_event: threading.Event | None = None) -> None:
        """Block until a slot is available. Reserves the slot atomically before sleeping
        so concurrent callers each get their own staggered slot."""
        if self.disabled:
            return
        wait = 0.0
        with self._lock:
            now = time.perf_counter()
            if self._next_slot <= now:
                self._next_slot = now + self.interval
            else:
                wait = self._next_slot - now
                self._next_slot += self.interval
        if wait > 0:
            if stop_event is not None:
                stop_event.wait(timeout=wait)
            else:
                time.sleep(wait)
