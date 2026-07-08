"""LLM client（databook 註釋 / headlines 用）— 完全 config-driven，方便升級/轉 model。

專用於 databook 自動化（低量、高質 reasoning），同 pipeline 個 tagging LLMClient 分開。

一切都可由 config 控制（唔使改 code）：provider / base_url / api_version / model / models 對照 / headers。
解析次序：env 覆蓋 > conf/local/credentials.yml『workbench:』section > code 內置預設。

金鑰同 config 位（repo PUBLIC，key 永不 hardcode/入 git）：
  1. env  WB_API_KEY / WB_BASE_URL / WB_API_VERSION / WB_MODEL / WB_PROVIDER / WB_CHARGE_CODE / WB_REGION
  2. conf/local/credentials.yml → workbench.*   （gitignored，手動 Mac→Windows copy）

conf/local/credentials.yml 例：
  workbench:
    provider: "azure"          # azure（KPMG Workbench）| openai（OpenAI-compatible，如 qwen）
    api_key: "..."
    base_url: "https://api.workbench.kpmg/genai/azure/openai"
    api_version: "2024-12-01-preview"
    charge_code: "0000"
    region: "westeurope"
    model: "5.5"               # 預設；填 alias 或完整 deployment/model 名
    models:                    # alias → 實際名（可加 qwen 之類）
      "5.5": "gpt-5-5-2026-04-24-gs-sdc"
      "5.4": "gpt-5-4-2026-03-05-gs-sdc"
      # "qwen": "qwen2.5-72b-instruct"

Client 係 lazy：import 唔會出網；真正 chat() 先建 SDK client。
"""
from __future__ import annotations

import json
import os
import re
from pathlib import Path
from typing import Any

# code 內置預設（config 冇填先用）─────────────────────────────────────
DEFAULT_PROVIDER = "azure"
DEFAULT_BASE_URL = "https://api.workbench.kpmg/genai/azure/openai"
DEFAULT_API_VERSION = "2024-12-01-preview"
DEFAULT_MODEL = "5.5"
DEFAULT_MODELS = {
    "5.5": "gpt-5-5-2026-04-24-gs-sdc",
    "5.4": "gpt-5-4-2026-03-05-gs-sdc",
}
_CRED = Path("conf/local/credentials.yml")


def _cred_section() -> dict:
    if not _CRED.exists():
        return {}
    try:
        import yaml
        d = yaml.safe_load(_CRED.read_text(encoding="utf-8")) or {}
        return dict(d.get("workbench") or {})
    except Exception:
        return {}


def _resolve(name: str, env: str, default: str = "") -> str:
    """env > config > default（字串）。"""
    v = os.environ.get(env, "").strip()
    if v:
        return v
    v = str(_cred_section().get(name, "") or "").strip()
    return v or default


def _parse_json_lenient(text: str) -> dict:
    s = (text or "").strip()
    if s.startswith("```"):
        s = re.sub(r"^```[a-zA-Z]*\s*", "", s)
        s = re.sub(r"\s*```$", "", s).strip()
    try:
        return json.loads(s)
    except json.JSONDecodeError:
        m = re.search(r"\{[\s\S]*\}", s)
        if not m:
            raise
        return json.loads(m.group(0))


class Workbench:
    def __init__(self, model: str | None = None, *, provider: str | None = None,
                 charge_code: str | None = None, region: str | None = None,
                 base_url: str | None = None, api_version: str | None = None,
                 api_key: str | None = None, models: dict | None = None):
        sec = _cred_section()
        # alias → 實際名 對照表：內置 ← config.models 覆蓋 ← 呼叫者覆蓋
        self.models = {**DEFAULT_MODELS, **(sec.get("models") or {}), **(models or {})}
        self.provider = (provider or _resolve("provider", "WB_PROVIDER", DEFAULT_PROVIDER)).lower()
        self.model_alias = model or _resolve("model", "WB_MODEL", DEFAULT_MODEL)
        self.base_url = base_url or _resolve("base_url", "WB_BASE_URL", DEFAULT_BASE_URL)
        self.api_version = api_version or _resolve("api_version", "WB_API_VERSION", DEFAULT_API_VERSION)
        self.charge_code = charge_code or _resolve("charge_code", "WB_CHARGE_CODE", "0000")
        self.region = region or _resolve("region", "WB_REGION", "westeurope")
        self._api_key = (api_key or "").strip() or None
        self._client = None

    def resolve_model(self, model: str | None = None) -> str:
        """alias → 實際 deployment/model 名（唔喺對照表就當已經係實名）。"""
        m = model or self.model_alias
        return self.models.get(m, m)

    @property
    def model(self) -> str:
        return self.resolve_model()

    # ── key / headers ────────────────────────────────────────────────
    @property
    def api_key(self) -> str:
        if not self._api_key:
            k = _resolve("api_key", "WB_API_KEY")
            if not k:
                raise RuntimeError(
                    "冇 API key：set env WB_API_KEY，"
                    "或 conf/local/credentials.yml 加 workbench.api_key（gitignored）")
            self._api_key = k
        return self._api_key

    def _headers(self) -> dict:
        # KPMG Workbench(azure) 專用 header；openai-compatible（qwen）唔加，SDK 自己用 api_key。
        if self.provider == "azure":
            return {
                "Ocp-Apim-Subscription-Key": self.api_key,
                "x-kpmg-charge-code": self.charge_code,
                "x-kpmg-region-override": self.region,
            }
        return {}

    def _client_obj(self):
        if self._client is None:
            if self.provider == "azure":
                from openai import AzureOpenAI
                self._client = AzureOpenAI(
                    api_key=self.api_key, base_url=self.base_url,
                    api_version=self.api_version, default_headers=self._headers() or None)
            else:  # openai-compatible（qwen 等）
                from openai import OpenAI
                self._client = OpenAI(
                    api_key=self.api_key, base_url=self.base_url,
                    default_headers=self._headers() or None)
        return self._client

    # ── chat ─────────────────────────────────────────────────────────
    def chat(self, user: str, system: str = "You are a helpful assistant.", *,
             model: str | None = None, reasoning_effort: str | None = "high",
             temperature: float | None = 1, max_tokens: int | None = None,
             json_mode: bool = False) -> str:
        """一問一答，回 content。對唔支持嘅參數自動 drop 再試，唔會因 model/provider 差異炒。"""
        cli = self._client_obj()
        kw: dict[str, Any] = {
            "model": self.resolve_model(model),
            "messages": [
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
        }
        hdr = self._headers()
        if hdr:
            kw["extra_headers"] = {"Content-Type": "application/json", **hdr}
        if temperature is not None:
            kw["temperature"] = temperature
        if reasoning_effort:
            kw["reasoning_effort"] = reasoning_effort
        if max_tokens:
            kw["max_tokens"] = max_tokens
        if json_mode:
            kw["response_format"] = {"type": "json_object"}

        droppable = ["reasoning_effort", "temperature", "response_format", "max_tokens"]
        for _ in range(len(droppable) + 1):
            try:
                resp = cli.chat.completions.create(**kw)
                return resp.choices[0].message.content or ""
            except Exception as e:
                msg = str(e).lower().replace("_", "")
                hit = next((p for p in droppable if p in kw and p.replace("_", "") in msg), None)
                if hit is None:
                    raise
                kw.pop(hit, None)
        raise RuntimeError("chat 失敗：所有可 drop 參數都試過")

    def chat_json(self, user: str, system: str = "You are a helpful assistant. Reply with JSON only.",
                  **kw) -> dict:
        kw.setdefault("json_mode", True)
        return _parse_json_lenient(self.chat(user, system, **kw))

    def ping(self, model: str | None = None) -> str:
        return self.chat("hi", model=model)

    # ── config 檢視（key 遮蔽）────────────────────────────────────────
    def config_masked(self) -> dict:
        try:
            k = self.api_key
            km = f"{k[:4]}…({len(k)} chars)" if k else "(缺)"
            key_ok = True
        except Exception as e:
            km, key_ok = f"⚠ {e}", False
        return {
            "provider": self.provider,
            "model_alias": self.model_alias, "model(resolved)": self.model,
            "models_map": self.models,
            "base_url": self.base_url, "api_version": self.api_version,
            "charge_code": self.charge_code, "region": self.region,
            "api_key": km, "key_ok": key_ok,
        }


def _cli() -> None:
    import argparse
    ap = argparse.ArgumentParser(description="LLM client（databook 用，config-driven）")
    ap.add_argument("--model", default=None, help="alias（5.5/5.4/qwen…）或完整名；預設由 config")
    ap.add_argument("--provider", default=None, help="azure / openai（覆蓋 config）")
    ap.add_argument("--ping", action="store_true", help="送 'hi' 驗連線（要喺 KPMG 網內跑）")
    ap.add_argument("--config", action="store_true", help="offline：只印解析到嘅 config（key 遮蔽）")
    ap.add_argument("--ask", help="送一句 user prompt 睇回覆")
    a = ap.parse_args()
    wb = Workbench(model=a.model, provider=a.provider)
    if a.config or not (a.ping or a.ask):
        print("LLM config（offline，唔出網）:")
        for k, v in wb.config_masked().items():
            print(f"  {k}: {v}")
        if not (a.ping or a.ask):
            print("\n→ 連線測試喺 KPMG 網內跑： python -m kpi.lib.workbench --ping")
            return
    if a.ping:
        print("\n--ping →", wb.ping())
    if a.ask:
        print("\n--ask →", wb.chat(a.ask))


if __name__ == "__main__":
    _cli()
