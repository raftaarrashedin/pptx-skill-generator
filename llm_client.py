"""
llm_client.py

Shared LLM configuration + invocation helpers for the presentation builder.

This keeps model selection, provider wiring, and deepagents invocation in one
place so generation, revision, and orchestration all use the same live model.
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from typing import Any, Sequence

DEFAULT_LLM_ID = "openai/qwen3.6:35b"

MODEL_REGISTRY = {
    DEFAULT_LLM_ID: {
        "MODEL_NAME": "qwen",
        "INFERENCE_ENGINE": "openai",
        "API_KEY": "none",
        "API_URL": "",
        "DESCRIPTION": "A high-performance model optimized for coding and complex reasoning.",
    },
}


@dataclass(frozen=True)
class LLMConfig:
    registry_id: str
    model_name: str
    inference_engine: str
    api_key: str
    api_url: str
    supports_vision: bool
    supports_thinking: bool
    active: bool
    supports_deep_research: bool
    description: str
    description_ar: str


def get_llm_config(model_id: str = DEFAULT_LLM_ID) -> LLMConfig:
    raw = MODEL_REGISTRY.get(model_id)
    if raw is None:
        raise KeyError(f"Unknown LLM config: {model_id}")

    return LLMConfig(
        registry_id=model_id,
        model_name=raw["MODEL_NAME"],
        inference_engine=raw["INFERENCE_ENGINE"],
        api_key=raw["API_KEY"],
    )


def llm_enabled(model_id: str = DEFAULT_LLM_ID) -> bool:
    config = get_llm_config(model_id)
    return bool(
        config.active
        and config.model_name
        and config.inference_engine
        and config.api_url
    )


def active_llm_summary(model_id: str = DEFAULT_LLM_ID) -> str:
    config = get_llm_config(model_id)
    return f"{config.registry_id} via {config.inference_engine} @ {config.api_url}"


def _configure_openai_compatible_env(config: LLMConfig) -> None:
    # Keep any user-provided OPENAI_API_KEY override, otherwise use the
    # registry default for this OpenAI-compatible endpoint.
    os.environ["OPENAI_API_KEY"] = os.environ.get("OPENAI_API_KEY", config.api_key or "none")
    os.environ["OPENAI_BASE_URL"] = config.api_url
    os.environ["OPENAI_API_BASE"] = config.api_url


def _deepagents_model_selector(config: LLMConfig) -> str:
    # deepagents uses provider-prefixed selectors such as anthropic:... .
    return f"{config.inference_engine}:{config.model_name}"


def create_llm_agent(
    *,
    system_prompt: str,
    response_format: Any | None = None,
    tools: Sequence[Any] | None = None,
    model_id: str = DEFAULT_LLM_ID,
):
    if not llm_enabled(model_id):
        raise RuntimeError(f"Live LLM is not enabled for config {model_id}")

    config = get_llm_config(model_id)
    if config.inference_engine != "openai":
        raise ValueError(f"Unsupported inference engine: {config.inference_engine}")

    _configure_openai_compatible_env(config)

    from deepagents import create_deep_agent

    agent_kwargs: dict[str, Any] = {
        "model": _deepagents_model_selector(config),
        "system_prompt": system_prompt,
    }
    if response_format is not None:
        agent_kwargs["response_format"] = response_format
    if tools is not None:
        agent_kwargs["tools"] = list(tools)

    return create_deep_agent(**agent_kwargs)


def invoke_llm(
    *,
    system_prompt: str,
    user_message: str | None = None,
    messages: Sequence[dict[str, Any]] | None = None,
    response_format: Any | None = None,
    tools: Sequence[Any] | None = None,
    model_id: str = DEFAULT_LLM_ID,
) -> dict[str, Any]:
    if messages is None:
        if user_message is None:
            raise ValueError("Either user_message or messages must be provided")
        messages = [{"role": "user", "content": user_message}]

    agent = create_llm_agent(
        system_prompt=system_prompt,
        response_format=response_format,
        tools=tools,
        model_id=model_id,
    )
    return agent.invoke({"messages": list(messages)})
