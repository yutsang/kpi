"""Force stdout/stderr unbuffered so tqdm progress is visible in real time on
Windows terminals and IDE consoles (which otherwise fully-buffer streams when
they don't detect a TTY).

Also exposes ProgressRefresher: a background thread that periodically calls
bar.refresh() so the tqdm display stays current even when the main thread is
blocked inside a slow operation (HTTP call etc.). Without this, the bar visibly
freezes during long iterations and users have to Ctrl+C to see progress."""
from __future__ import annotations

import sys
import threading


def force_unbuffered_io() -> None:
    for stream in (sys.stdout, sys.stderr):
        try:
            # Force UTF-8 + replace-on-error so Chinese / non-ASCII prints don't crash on
            # Windows consoles (default cp950/cp1252 chokes with 'charmap' codec errors).
            stream.reconfigure(
                write_through=True,
                line_buffering=True,
                encoding="utf-8",
                errors="replace",
            )
        except Exception:
            try:
                stream.reconfigure(write_through=True, line_buffering=True)
            except Exception:
                pass
        try:
            stream.flush()
        except Exception:
            pass


class ProgressRefresher:
    """Background-thread tqdm refresher. Use as a context manager:

        bar = tqdm(...)
        with ProgressRefresher(bar):
            for item in bar:
                slow_blocking_call()   # bar still visibly refreshes every 0.5s
    """

    def __init__(self, bar, interval: float = 0.5):
        self.bar = bar
        self.interval = interval
        self._stop = threading.Event()
        self._thread: threading.Thread | None = None

    def start(self) -> "ProgressRefresher":
        if self._thread is not None:
            return self
        self._thread = threading.Thread(target=self._loop, daemon=True, name="tqdm-refresher")
        self._thread.start()
        return self

    def stop(self) -> None:
        self._stop.set()
        if self._thread is not None:
            self._thread.join(timeout=1.0)
            self._thread = None

    def __enter__(self) -> "ProgressRefresher":
        return self.start()

    def __exit__(self, *exc) -> None:
        self.stop()

    def _loop(self) -> None:
        while not self._stop.is_set():
            try:
                bar = self.bar
                if bar is not None:
                    bar.refresh()
                    sys.stderr.flush()
            except Exception:
                pass
            self._stop.wait(self.interval)
