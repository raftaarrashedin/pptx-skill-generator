"""
review_presentation_draft.py

Deterministic (no LLM) step. Reads presentation_draft.json and renders a
readable slide-by-slide outline so a human can approve before any expensive
rendering/export happens. Also writes a compact preview.json a UI could
consume later.

Pass --approve to copy the current draft into presentation_final.json once
you're happy with it -- this is the only thing that unlocks render/export.

Outputs: output/review.txt, output/preview.json, (optionally) output/presentation_final.json

Usage:
    python review_presentation_draft.py
    python review_presentation_draft.py --approve
"""

import argparse
import json
import os
import shutil
from datetime import datetime, timezone


def log(output_dir, message):
    path = os.path.join(output_dir, "run.log")
    with open(path, "a", encoding="utf-8") as f:
        f.write(f"{datetime.now(timezone.utc).isoformat()} | review_presentation_draft | {message}\n")


def _slide_outline(slide: dict) -> str:
    lines = [f"Slide {slide.get('id')}: [{slide.get('type')}] {slide.get('title') or ''}".rstrip()]
    if slide.get("subtitle"):
        lines.append(f"    subtitle: {slide['subtitle']}")
    for b in slide.get("bullets", []):
        lines.append(f"    - {b}")
    for m in slide.get("metrics", []):
        lines.append(f"    metric: {m.get('label')} = {m.get('value')}")
    if slide.get("comparison"):
        comp = slide["comparison"]
        left = comp.get("left", {})
        right = comp.get("right", {})
        lines.append(f"    left  [{left.get('heading')}]: {', '.join(left.get('items', []))}")
        lines.append(f"    right [{right.get('heading')}]: {', '.join(right.get('items', []))}")
    if slide.get("chart"):
        chart = slide["chart"]
        lines.append(f"    chart ({chart.get('chart_type')}): categories={chart.get('categories')}")
        for s in chart.get("series", []):
            lines.append(f"      series '{s.get('name')}': {s.get('values')}")
    if slide.get("image"):
        img = slide["image"]
        lines.append(f"    image: {img.get('placeholder_label')} ({img.get('path') or 'no path -- placeholder'})")
    if slide.get("notes"):
        lines.append(f"    notes: {slide['notes']}")
    return "\n".join(lines)


def build_review(draft: dict) -> str:
    header = [
        f"Title:    {draft.get('title')}",
        f"Subtitle: {draft.get('subtitle') or ''}",
        f"Language: {draft.get('language')}",
        f"Audience: {draft.get('audience') or ''}",
        f"Slides:   {len(draft.get('slides', []))}",
        "-" * 60,
    ]
    body = [_slide_outline(s) for s in draft.get("slides", [])]
    return "\n".join(header) + "\n\n" + "\n\n".join(body) + "\n"


def build_preview(draft: dict) -> dict:
    slides = draft.get("slides", [])
    type_counts = {}
    flags = []
    for s in slides:
        type_counts[s.get("type")] = type_counts.get(s.get("type"), 0) + 1
        if s.get("type") == "bullets" and not s.get("bullets"):
            flags.append(f"Slide {s.get('id')}: bullets slide has no bullets")
        if s.get("type") == "title" and not s.get("title"):
            flags.append(f"Slide {s.get('id')}: title slide missing a title")
        if s.get("type") == "image_text" and not s.get("image"):
            flags.append(f"Slide {s.get('id')}: image_text slide missing image spec")

    return {
        "title": draft.get("title"),
        "slide_count": len(slides),
        "slide_type_counts": type_counts,
        "flags": flags,
    }


def main():
    parser = argparse.ArgumentParser(description="Review presentation_draft.json")
    parser.add_argument("--output-dir", default="output")
    parser.add_argument("--approve", action="store_true", help="Copy draft -> presentation_final.json")
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

    review_text = build_review(draft)
    preview = build_preview(draft)

    review_path = os.path.join(args.output_dir, "review.txt")
    preview_path = os.path.join(args.output_dir, "preview.json")

    with open(review_path, "w", encoding="utf-8") as f:
        f.write(review_text)
    with open(preview_path, "w", encoding="utf-8") as f:
        json.dump(preview, f, indent=2)

    log(args.output_dir, f"output file written: {review_path}")
    log(args.output_dir, f"output file written: {preview_path}")

    print(review_text)
    if preview["flags"]:
        print("FLAGS:")
        for flag in preview["flags"]:
            print(f"  - {flag}")

    if args.approve:
        final_path = os.path.join(args.output_dir, "presentation_final.json")
        shutil.copyfile(draft_path, final_path)
        log(args.output_dir, f"approved: copied {draft_path} -> {final_path}")
        print(f"\nApproved. Wrote {final_path}")
    else:
        print("\nNot approved yet. Re-run with --approve once you're happy with this draft,")
        print("or use revise_presentation_draft.py to request changes first.")


if __name__ == "__main__":
    main()
