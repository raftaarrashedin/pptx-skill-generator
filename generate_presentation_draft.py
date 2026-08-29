"""
generate_presentation_draft.py

The content-generation agent. Takes template_context.json + presentation_request.json
and produces ONE structured file: presentation_draft.json. This JSON is the
source of truth for the whole pipeline -- not the eventual .pptx.

Uses a deepagents agent constrained to a pydantic response_format, so the
output shape is guaranteed (no hand-parsing of LLM text).

Run with --mock to exercise the full file plumbing without calling a real
model -- useful for testing the pipeline before spending tokens.

Output: output/presentation_draft.json

Usage:
    python generate_presentation_draft.py
    python generate_presentation_draft.py --mock
"""

import argparse
import json
import os
from datetime import datetime, timezone
from typing import Literal, Optional

from pydantic import BaseModel, Field

from llm_client import invoke_llm, llm_enabled

SLIDE_TYPES = ("title", "section", "bullets", "metrics", "comparison", "chart", "image_text")


# ---------------------------------------------------------------------------
# Schema -- this is the contract every later script (review/render/export)
# relies on. Keep it small and stable.
# ---------------------------------------------------------------------------

class Metric(BaseModel):
    label: str
    value: str


class ComparisonSide(BaseModel):
    heading: str
    items: list[str] = Field(default_factory=list)


class Comparison(BaseModel):
    left: ComparisonSide
    right: ComparisonSide


class ChartSeries(BaseModel):
    name: str
    values: list[float]


class Chart(BaseModel):
    chart_type: Literal["bar", "line", "pie"]
    categories: list[str]
    series: list[ChartSeries]


class ImageSpec(BaseModel):
    placeholder_label: str
    path: Optional[str] = None


class Slide(BaseModel):
    id: int
    type: Literal[SLIDE_TYPES]
    title: Optional[str] = None
    subtitle: Optional[str] = None
    bullets: list[str] = Field(default_factory=list)
    metrics: list[Metric] = Field(default_factory=list)
    comparison: Optional[Comparison] = None
    chart: Optional[Chart] = None
    image: Optional[ImageSpec] = None
    notes: Optional[str] = None
    layout_intent: Optional[str] = None


class PresentationDraft(BaseModel):
    title: str
    subtitle: Optional[str] = None
    language: str = "en"
    audience: Optional[str] = None
    slides: list[Slide]


# ---------------------------------------------------------------------------

def log(output_dir, message):
    path = os.path.join(output_dir, "run.log")
    with open(path, "a", encoding="utf-8") as f:
        f.write(f"{datetime.now(timezone.utc).isoformat()} | generate_presentation_draft | {message}\n")


SYSTEM_PROMPT = """You are a presentation content-generation agent.

You will be given:
1. A template design context (fonts, colors, slide size) -- shape your content
   to fit the spirit of this design, you do not need to repeat it in the output.
2. A structured presentation request (title, topic, slide count, audience,
   detail level, language, optional source content).

Produce a complete slide-by-slide presentation draft as your structured
response. Rules:
- Use only these slide types: title, section, bullets, metrics, comparison,
  chart, image_text.
- The first slide must be type "title". If the request implies distinct
  sections, use "section" slides as dividers.
- Bullets: 3-5 per slide, each a short phrase, not a full paragraph.
- Match the requested slide_count as closely as possible.
- Match the requested language and detail_level (low = sparse, high = more
  detailed bullets/notes).
- Write a one-line speaker note for each slide.
- Do not invent specific factual claims, numbers, or statistics if no source
  content was provided -- use clearly illustrative/placeholder language
  instead (e.g. "X% improvement" framed as illustrative, not asserted fact).
"""


def _build_ai_mock_draft(title: str, topic: str, audience: Optional[str], language: str, image_path: Optional[str]) -> dict:
    slides = [
        {
            "id": 1,
            "type": "title",
            "title": title,
            "subtitle": "Executive overview and adoption roadmap",
            "notes": "Open with the strategic importance of AI and the need for disciplined adoption.",
        },
        {
            "id": 2,
            "type": "bullets",
            "title": "Executive Summary",
            "bullets": [
                "AI is shifting from experimentation to operating model change",
                "The strongest value comes from focused, repeatable use cases",
                "Governance and data readiness matter as much as the model choice",
                "Leadership decisions should balance speed, risk, and measurable outcomes",
            ],
            "notes": "Frame AI as a business transformation topic, not just a tooling discussion.",
        },
        {
            "id": 3,
            "type": "section",
            "title": "Why AI Matters Now",
            "subtitle": f"Implications for {audience or 'leadership teams'}",
            "notes": "Use this divider to move from context into opportunity.",
        },
        {
            "id": 4,
            "type": "bullets",
            "title": "Where AI Creates Value",
            "bullets": [
                "Automates repetitive knowledge work and speeds up decision support",
                "Improves customer and employee experiences through faster responses",
                "Expands analytical capacity across documents, data, and workflows",
                "Raises the productivity ceiling for teams that pair AI with human review",
            ],
            "notes": "Keep the examples practical and tied to operating leverage.",
        },
        {
            "id": 5,
            "type": "metrics",
            "title": "Leadership Priorities",
            "metrics": [
                {"label": "Initial focus areas", "value": "3"},
                {"label": "Planning horizon", "value": "12 mo"},
                {"label": "Core governance owner", "value": "1"},
            ],
            "notes": "These are planning anchors, not market statistics.",
        },
        {
            "id": 6,
            "type": "comparison",
            "title": "Traditional Automation vs AI Systems",
            "comparison": {
                "left": {
                    "heading": "Traditional Automation",
                    "items": [
                        "Rule-based and predictable",
                        "Best for stable, repetitive processes",
                        "Limited flexibility with unstructured inputs",
                    ],
                },
                "right": {
                    "heading": "AI Systems",
                    "items": [
                        "Probabilistic and context-aware",
                        "Best for language, search, and synthesis tasks",
                        "Requires oversight, evaluation, and governance",
                    ],
                },
            },
            "notes": "Highlight that AI complements rather than replaces standard automation.",
        },
        {
            "id": 7,
            "type": "chart",
            "title": "Illustrative AI Adoption Roadmap",
            "chart": {
                "chart_type": "bar",
                "categories": ["Discover", "Pilot", "Scale", "Govern"],
                "series": [{"name": "Program maturity", "values": [1, 2, 3, 4]}],
            },
            "notes": "Explain this as a sequencing model rather than a measured benchmark.",
        },
        {
            "id": 8,
            "type": "image_text",
            "title": "AI Operating Model",
            "bullets": [
                "Align business sponsors, domain experts, and technical teams",
                "Standardize data access, model evaluation, and human review",
                "Create a repeatable path from prototype to production use",
            ],
            "image": {
                "placeholder_label": "AI workflow / governance visual",
                "path": image_path,
            },
            "notes": "Use the visual to show how strategy, execution, and controls connect.",
        },
        {
            "id": 9,
            "type": "section",
            "title": "Risk and Governance",
            "subtitle": "What leadership should manage explicitly",
            "notes": "Transition from opportunity to control points.",
        },
        {
            "id": 10,
            "type": "bullets",
            "title": "Recommended Next Steps",
            "bullets": [
                "Select a small number of high-value AI use cases",
                "Define success metrics, owners, and review checkpoints",
                "Establish governance for data, security, and model behavior",
                "Scale only after measurable wins and clear operating controls",
            ],
            "notes": "Close with a concrete, staged adoption plan.",
        },
    ]

    return {
        "title": title,
        "subtitle": topic,
        "language": language,
        "audience": audience,
        "slides": slides,
    }


def _mock_draft(request: dict, template_context: Optional[dict] = None) -> dict:
    """Deterministic stand-in draft, used for --mock and as an offline fallback.

    It stays stable for testing, but still adapts content to the request topic
    so the generated deck is usable when a live LLM call is unavailable.
    """
    slide_count = max(3, int(request.get("slide_count", 8)))
    title = request.get("title", "Untitled Presentation")
    topic = request.get("topic", "")
    audience = request.get("audience")
    language = request.get("language", "en")
    template_images = (template_context or {}).get("images", [])
    image_path = template_images[0] if template_images else None

    normalized_topic = f"{title} {topic}".lower()
    if "artificial intelligence" in normalized_topic or normalized_topic.strip() == "ai" or " ai" in normalized_topic:
        return _build_ai_mock_draft(title, topic or title, audience, language, image_path)

    pool = [
        {
            "type": "section",
            "title": f"Context for {topic or title}",
            "subtitle": "Topic framing and background",
        },
        {
            "type": "bullets",
            "title": "Key Themes",
            "bullets": [
                f"Core idea shaping {topic or title}".strip() or "Core idea",
                "Primary implication for the audience",
                "Key decision or takeaway",
            ],
        },
        {
            "type": "metrics",
            "title": "Planning Snapshot",
            "metrics": [
                {"label": "Priority areas", "value": "3"},
                {"label": "Review cadence", "value": "Monthly"},
                {"label": "Owner groups", "value": "2"},
            ],
        },
        {
            "type": "comparison",
            "title": "Current State / Future State",
            "comparison": {
                "left": {"heading": "Current", "items": ["Existing challenge 1", "Existing challenge 2"]},
                "right": {"heading": "Future", "items": ["Improved outcome 1", "Improved outcome 2"]},
            },
        },
        {
            "type": "chart",
            "title": "Illustrative Roadmap",
            "chart": {
                "chart_type": "bar",
                "categories": ["Plan", "Build", "Launch", "Improve"],
                "series": [{"name": "Program progress", "values": [1, 2, 3, 4]}],
            },
        },
        {
            "type": "image_text",
            "title": "Visual Reference",
            "bullets": ["Supporting note next to the visual", "Context for the supporting asset"],
            "image": {"placeholder_label": "Diagram / screenshot goes here", "path": image_path},
        },
    ]

    slides = [
        {
            "id": 1,
            "type": "title",
            "title": title,
            "subtitle": topic,
            "notes": "Open with a one-line framing of why this matters to the audience.",
        },
        {
            "id": 2,
            "type": "bullets",
            "title": "Overview",
            "bullets": [
                f"What {topic or title} is",
                f"Why {topic or title} matters now",
                "What this presentation will cover",
            ],
            "notes": "Set expectations for the rest of the deck.",
        },
    ]

    next_id = 3
    pool_idx = 0
    while len(slides) < slide_count - 1:
        template = dict(pool[pool_idx % len(pool)])
        template["id"] = next_id
        template.setdefault("notes", "Placeholder content -- generated in --mock mode.")
        slides.append(template)
        next_id += 1
        pool_idx += 1

    slides.append(
        {
            "id": next_id,
            "type": "bullets",
            "title": "Next Steps",
            "bullets": ["Decision needed", "Owner and timeline", "How success will be measured"],
            "notes": "Close with a clear ask.",
        }
    )

    return {
        "title": title,
        "subtitle": topic,
        "language": language,
        "audience": audience,
        "slides": slides,
    }


def generate_draft(template_context: dict, request: dict, mock: bool = False) -> dict:
    if mock or not llm_enabled():
        raw = _mock_draft(request, template_context)
        # Validate/normalize through the schema so downstream scripts always
        # see the same set of keys, mock or real.
        return PresentationDraft(**raw).model_dump()

    user_message = (
        "TEMPLATE CONTEXT:\n"
        f"{json.dumps(template_context, indent=2)}\n\n"
        "PRESENTATION REQUEST:\n"
        f"{json.dumps(request, indent=2)}\n\n"
        "Generate the presentation draft now."
    )

    result = invoke_llm(
        system_prompt=SYSTEM_PROMPT,
        user_message=user_message,
        response_format=PresentationDraft,
    )
    structured = result["structured_response"]
    if isinstance(structured, BaseModel):
        return structured.model_dump()
    return structured


def main():
    parser = argparse.ArgumentParser(description="Generate presentation_draft.json")
    parser.add_argument("--output-dir", default="output")
    parser.add_argument("--mock", action="store_true", help="Skip the LLM call, use deterministic stub content")
    args = parser.parse_args()

    os.makedirs(args.output_dir, exist_ok=True)
    log(args.output_dir, "started")

    ctx_path = os.path.join(args.output_dir, "template_context.json")
    req_path = os.path.join(args.output_dir, "presentation_request.json")

    for p in (ctx_path, req_path):
        if not os.path.exists(p):
            log(args.output_dir, f"ERROR: missing required input {p}")
            raise SystemExit(f"Missing required input: {p}. Run the earlier pipeline steps first.")

    with open(ctx_path, encoding="utf-8") as f:
        template_context = json.load(f)
    with open(req_path, encoding="utf-8") as f:
        request = json.load(f)

    log(args.output_dir, f"input file used: {ctx_path}, {req_path}")

    try:
        draft = generate_draft(template_context, request, mock=args.mock)
    except Exception as e:
        log(args.output_dir, f"ERROR: {e}")
        raise

    out_path = os.path.join(args.output_dir, "presentation_draft.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(draft, f, indent=2)

    log(args.output_dir, f"output file written: {out_path}")
    print(f"Wrote {out_path}")
    print(json.dumps(draft, indent=2)[:2000])


if __name__ == "__main__":
    main()
