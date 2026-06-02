"""Rerank client for company endpoints (e.g. bgem3v2_rerank).

Given a query and a list of candidate documents, returns relevance scores per
document (higher = more relevant). Used as a middle layer between explicit rules
and full LLM tagging.
"""
from __future__ import annotations

import json
from pathlib import Path
from typing import Optional

import httpx

from .cache import JsonlCache


class RerankClient:
    def __init__(self, cfg: dict, cache_path: Path | str | None = None):
        rcfg = cfg.get("rerank") or {}
        if not rcfg.get("enabled", False):
            self.enabled = False
            return
        self.enabled = True
        api_base = rcfg.get("api_base") or ""
        if not api_base:
            raise ValueError("rerank.api_base is required when rerank.enabled is true")
        self.api_base = api_base
        self.api_key = rcfg.get("api_key") or cfg.get("llm", {}).get("api_key", "")
        self.model = rcfg.get("model") or "bgem3v2_rerank"
        self.timeout = int(rcfg.get("request_timeout", 30) or 30)
        self.score_threshold = float(rcfg.get("score_threshold", 0.6) or 0.6)
        verify = bool(rcfg.get("verify_ssl", False))
        self._client = httpx.Client(verify=verify, timeout=self.timeout)
        self.cache = JsonlCache(Path(cache_path)) if cache_path else None

    def rerank(self, query: str, documents: list[str]) -> list[float]:
        """Return scores in same order as documents (always len == len(documents))."""
        if not self.enabled or not documents:
            return [0.0] * len(documents)
        cache_key = None
        if self.cache is not None:
            cache_key = JsonlCache.make_key(self.model, query, documents)
            hit = self.cache.get(cache_key)
            if hit is not None:
                return list(hit)
        headers = {
            "authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }
        body = {"documents": documents, "query": query, "model": self.model}
        data = json.dumps(body, ensure_ascii=False).encode("utf-8")
        resp = self._client.post(self.api_base, headers=headers, content=data)
        resp.raise_for_status()
        result = resp.json()
        scores = self._parse_scores(result, len(documents))
        if self.cache is not None and cache_key:
            self.cache.put(cache_key, scores)
        return scores

    @staticmethod
    def _parse_scores(payload: dict, n: int) -> list[float]:
        items = (
            payload.get("results")
            or payload.get("data")
            or (payload.get("response") or {}).get("results")
            or []
        )
        scores = [0.0] * n
        for item in items:
            idx = item.get("index")
            if idx is None:
                continue
            score = (
                item.get("relevance_score")
                if item.get("relevance_score") is not None
                else item.get("score")
            )
            if score is None:
                continue
            try:
                idx = int(idx)
                if 0 <= idx < n:
                    scores[idx] = float(score)
            except (TypeError, ValueError):
                continue
        return scores

    def best_match(
        self,
        query: str,
        candidates: list[dict],
        label_key: str = "label",
        id_key: str = "id",
    ) -> Optional[tuple[str, float]]:
        """Convenience: rerank candidate dicts and return (id, score) if above threshold."""
        if not self.enabled or not candidates:
            return None
        labels = [str(c.get(label_key, "")) for c in candidates]
        scores = self.rerank(query, labels)
        if not scores:
            return None
        top = max(range(len(scores)), key=lambda i: scores[i])
        if scores[top] >= self.score_threshold:
            return (candidates[top].get(id_key), scores[top])
        return None
