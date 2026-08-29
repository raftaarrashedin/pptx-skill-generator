"""
revise_presentation_draft.py

Accepts user feedback and updates the existing presentation_draft.json
in place, rather than regenerating from scratch. v1 scope, on purpose:
- changing text
- adding/removing bullets
- changing slide order
- adding/removing slides
- changing tone or emphasis
No advanced structural transformations yet.

Reuses the same PresentationDraft schema as generate_presentation_draft.py
so the file shape never drifts between the two.

Output: output/presentation_draft.json (overwritten)

Usage:
    python revise_presentation_draft.py --feedback "Make slide 2 punchier, drop slide 5, swap order of 3 and 4"
    python revise_presentation_draft.py --feedback "..." --mock
"""

import argparse
import json
import os
from datetime import datetime, timezone

from pydantic import BaseModel

from generate_presentation_draft import PresentationDraft
from llm_client import invoke_llm, llm_enabled

REVISION_SYSTEM_PROMPT = """You are a presentation revision agent.

You will be given the CURRENT presentation draft (full JSON) and FEEDBACK
from the user. Apply the feedback as targeted edits to the existing draft --
do not regenerate the presentation from scratch, and do not change slides
the feedback didn't touch.

Scope for v1: you may change text, add/remove bullets, reorder slides,
add/remove slides, and adjust tone/emphasis. Avoid major structural
transformations (e.g. don't invent new slide types not already in this
draft's vocabulary: title, section, bullets, metrics, comparison, chart,
image_text).

Keep slide `id` values sequential and unique after any reordering or
add/remove operations. Return the complete, updated draft as your
structured response.
"""


def log(output_dir, message):
    path = os.path.join(output_dir, "run.log")
    with open(path, "a", encoding="utf-8") as f:
        f.write(f"{datetime.now(timezone.utc).isoformat()} | revise_presentation_draft | {message}\n")


def _mock_revise(draft: dict, feedback: str) -> dict:
    """Deterministic stand-in: appends a note recording the feedback was
    received, without attempting to interpret it. Good enough to test
    plumbing in --mock mode; real interpretation needs the LLM."""
    revised = json.loads(json.dumps(draft))  # deep copy
    for slide in revised.get("slides", []):
        notes = slide.get("notes") or ""
        slide["notes"] = (notes + f" [mock-revision noted feedback: {feedback}]").strip()
    return revised


def revise_draft(draft: dict, feedback: str, mock: bool = False) -> dict:
    if mock or not llm_enabled():
        raw = _mock_revise(draft, feedback)
        return PresentationDraft(**raw).model_dump()

    user_message = (
        "CURRENT DRAFT:\n"
        f"{json.dumps(draft, indent=2)}\n\n"
        "FEEDBACK:\n"
        f"{feedback}\n\n"
        "Apply this feedback and return the complete updated draft."
    )

    result = invoke_llm(
        system_prompt=REVISION_SYSTEM_PROMPT,
        user_message=user_message,
        response_format=PresentationDraft,
    )
    structured = result["structured_response"]
    if isinstance(structured, BaseModel):
        return structured.model_dump()
    return structured


def main():
    parser = argparse.ArgumentParser(description="Revise presentation_draft.json based on feedback.")
    parser.add_argument("--feedback", required=True, help="Free-text revision feedback")
    parser.add_argument("--output-dir", default="output")
    parser.add_argument("--mock", action="store_true")
    args = parser.parse_args()

    os.makedirs(args.output_dir, exist_ok=True)
    log(args.output_dir, "started")

    draft_path = os.path.join(args.output_dir, "presentation_draft.json")
    if not os.path.exists(draft_path):
        log(args.output_dir, f"ERROR: missing input {draft_path}")
        raise SystemExit(f"Missing input: {draft_path}. Run generate_presentation_draft.py first.")

    with open(draft_path, encoding="utf-8") as f:
        draft = json.load(f)
    log(args.output_dir, f"input file used: {draft_path}")
    log(args.output_dir, f"feedback: {args.feedback}")

    try:
        revised = revise_draft(draft, args.feedback, mock=args.mock)
    except Exception as e:
        log(args.output_dir, f"ERROR: {e}")
        raise

    with open(draft_path, "w", encoding="utf-8") as f:
        json.dump(revised, f, indent=2)

    log(args.output_dir, f"output file written: {draft_path}")
    print(f"Wrote {draft_path}")
    print(json.dumps(revised, indent=2)[:2000])


if __name__ == "__main__":
    main()
