"""
run_presentation_orchestrator.py

The coordinator. This is the first real "deepagents" layer in the pipeline --
everything before this (template context, request capture) and after it
(render, export) is deliberately deterministic. Orchestration is where
agent-based reasoning actually earns its keep: deciding whether enough
information exists, whether to ask a follow-up question, whether to produce
a fresh draft, or revise an existing one.

It does this by giving a deepagents agent three thin tools that wrap the
plain functions in generate_presentation_draft.py and revise_presentation_draft.py
-- the agent decides which to call, the scripts still do the actual work.

Usage:
    # No request yet captured -> orchestrator will tell you to run
    # capture_presentation_request.py first.
    python run_presentation_orchestrator.py

    # Request exists, no draft yet -> generates one.
    python run_presentation_orchestrator.py

    # Draft exists, feedback given -> revises it.
    python run_presentation_orchestrator.py --feedback "Trim slide 4, add a metrics slide after it"

    # Test the coordination logic without calling a real model.
    python run_presentation_orchestrator.py --mock
"""

import argparse
import json
import os
from datetime import datetime, timezone

from capture_presentation_request import missing_fields
from generate_presentation_draft import generate_draft
from llm_client import invoke_llm, llm_enabled
from revise_presentation_draft import revise_draft


def log(output_dir, message):
    path = os.path.join(output_dir, "run.log")
    with open(path, "a", encoding="utf-8") as f:
        f.write(
            f"{datetime.now(timezone.utc).isoformat()} | run_presentation_orchestrator | {message}\n"
        )


ORCHESTRATOR_SYSTEM_PROMPT = """You are the orchestrator for a presentation-building pipeline.

Given the current state (whether a presentation request exists, whether a
draft already exists, and any feedback the user just gave), decide exactly
ONE next action and call exactly one tool to carry it out:

- If the request is missing required fields, call `ask_followup` explaining
  what's missing in one short, specific question. Do not call any other tool.
- If there is no draft yet, call `generate_draft_tool`.
- If a draft exists and the user gave feedback, call `revise_draft_tool`.
- If a draft exists and there is no feedback, call `ask_followup` and ask
  whether the user wants to revise it or approve it (approval itself happens
  in review_presentation_draft.py, not here).

Always call exactly one tool. Do not ask more than one question at a time.
"""


def _build_tools(output_dir, mock):
    from langchain_core.tools import tool

    @tool
    def ask_followup(question: str) -> str:
        """Ask the user a single, specific follow-up question instead of generating or revising."""
        return f"FOLLOWUP_QUESTION: {question}"

    @tool
    def generate_draft_tool() -> str:
        """Generate a brand-new presentation draft from template_context.json + presentation_request.json."""
        ctx_path = os.path.join(output_dir, "template_context.json")
        req_path = os.path.join(output_dir, "presentation_request.json")
        with open(ctx_path, encoding="utf-8") as f:
            template_context = json.load(f)
        with open(req_path, encoding="utf-8") as f:
            request = json.load(f)
        draft = generate_draft(template_context, request, mock=mock)
        draft_path = os.path.join(output_dir, "presentation_draft.json")
        with open(draft_path, "w", encoding="utf-8") as f:
            json.dump(draft, f, indent=2)
        log(output_dir, f"output file written: {draft_path} (via generate_draft_tool)")
        return f"Draft generated with {len(draft.get('slides', []))} slides and written to {draft_path}."

    @tool
    def revise_draft_tool(feedback: str) -> str:
        """Revise the existing presentation_draft.json based on user feedback."""
        draft_path = os.path.join(output_dir, "presentation_draft.json")
        with open(draft_path, encoding="utf-8") as f:
            draft = json.load(f)
        revised = revise_draft(draft, feedback, mock=mock)
        with open(draft_path, "w", encoding="utf-8") as f:
            json.dump(revised, f, indent=2)
        log(output_dir, f"output file written: {draft_path} (via revise_draft_tool)")
        return f"Draft revised and written to {draft_path}."

    return [ask_followup, generate_draft_tool, revise_draft_tool]


def _mock_decide(output_dir, request_exists, request, draft_exists, feedback):
    """Deterministic stand-in for the orchestrator's decision, used for --mock
    and as an offline fallback -- same priority order as the real agent."""
    if not request_exists:
        return "FOLLOWUP_QUESTION: No presentation request found yet. Run capture_presentation_request.py first."

    missing = missing_fields(request)
    if missing:
        return f"FOLLOWUP_QUESTION: Missing required field(s): {', '.join(missing)}. Please provide them."

    if not draft_exists:
        with open(
            os.path.join(output_dir, "template_context.json"), encoding="utf-8"
        ) as f:
            template_context = json.load(f)
        draft = generate_draft(template_context, request, mock=True)
        with open(
            os.path.join(output_dir, "presentation_draft.json"), "w", encoding="utf-8"
        ) as f:
            json.dump(draft, f, indent=2)
        log(
            output_dir,
            "output file written: presentation_draft.json (mock decide: generate)",
        )
        return f"Draft generated with {len(draft.get('slides', []))} slides."

    if feedback:
        with open(
            os.path.join(output_dir, "presentation_draft.json"), encoding="utf-8"
        ) as f:
            draft = json.load(f)
        revised = revise_draft(draft, feedback, mock=True)
        with open(
            os.path.join(output_dir, "presentation_draft.json"), "w", encoding="utf-8"
        ) as f:
            json.dump(revised, f, indent=2)
        log(
            output_dir,
            "output file written: presentation_draft.json (mock decide: revise)",
        )
        return "Draft revised."

    return "FOLLOWUP_QUESTION: A draft already exists. Revise it (--feedback) or approve it via review_presentation_draft.py --approve."


def orchestrate(output_dir: str, feedback: str = None, mock: bool = False) -> str:
    req_path = os.path.join(output_dir, "presentation_request.json")
    draft_path = os.path.join(output_dir, "presentation_draft.json")
    request_exists = os.path.exists(req_path)
    draft_exists = os.path.exists(draft_path)
    request = None
    if request_exists:
        with open(req_path, encoding="utf-8") as f:
            request = json.load(f)

    if mock or not llm_enabled():
        return _mock_decide(output_dir, request_exists, request, draft_exists, feedback)

    tools = _build_tools(output_dir, mock=False)
    state_summary = {
        "request_exists": request_exists,
        "request_missing_fields": (
            missing_fields(request) if request else ["<no request captured>"]
        ),
        "draft_exists": draft_exists,
        "user_feedback": feedback,
    }
    user_message = (
        f"CURRENT STATE:\n{json.dumps(state_summary, indent=2)}\n\nDecide and act."
    )

    result = invoke_llm(
        system_prompt=ORCHESTRATOR_SYSTEM_PROMPT,
        user_message=user_message,
        tools=tools,
    )
    final_message = result["messages"][-1]
    return getattr(final_message, "content", str(final_message))


def main():
    parser = argparse.ArgumentParser(
        description="Coordinate generate/revise/ask-followup."
    )
    parser.add_argument(
        "--feedback", default=None, help="User feedback on the current draft, if any"
    )
    parser.add_argument("--output-dir", default="output")
    parser.add_argument("--mock", action="store_true")
    args = parser.parse_args()

    os.makedirs(args.output_dir, exist_ok=True)
    log(args.output_dir, "started")
    if args.feedback:
        log(args.output_dir, f"feedback: {args.feedback}")

    try:
        result = orchestrate(args.output_dir, feedback=args.feedback, mock=args.mock)
    except Exception as e:
        log(args.output_dir, f"ERROR: {e}")
        raise

    log(args.output_dir, f"decision result: {result}")
    print(result)


if __name__ == "__main__":
    main()
