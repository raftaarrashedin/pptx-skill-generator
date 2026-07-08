"""
capture_presentation_request.py

Collects the user's presentation intent as structured fields (not one raw
prompt) so generation quality and later debugging don't depend on parsing
free text. Runs interactively by default; pass flags for non-interactive use
(scripting, the orchestrator, CI).

Output: output/presentation_request.json

Usage:
    python capture_presentation_request.py                     # interactive
    python capture_presentation_request.py \
        --title "Q3 Board Update" --topic "Quarterly performance review" \
        --slides 10 --language en --detail medium \
        --audience "Executive leadership" \
        --content-file notes.txt
"""

import argparse
import json
import os
from datetime import datetime, timezone

DETAIL_LEVELS = ("low", "medium", "high")


def log(output_dir, message):
    path = os.path.join(output_dir, "run.log")
    with open(path, "a", encoding="utf-8") as f:
        f.write(f"{datetime.now(timezone.utc).isoformat()} | capture_presentation_request | {message}\n")


def _prompt(label, default=None, required=False):
    suffix = f" [{default}]" if default is not None else ""
    while True:
        value = input(f"{label}{suffix}: ").strip()
        if not value and default is not None:
            return default
        if not value and required:
            print("  This field is required.")
            continue
        return value


def interactive_capture():
    print("=== Presentation Request ===")
    title = _prompt("Presentation title", required=True)
    topic = _prompt("Topic / what it's about", required=True)
    slide_count = _prompt("Approx. slide count", default="10")
    language = _prompt("Language", default="en")
    detail_level = _prompt(f"Detail level {DETAIL_LEVELS}", default="medium")
    audience = _prompt("Target audience", default="General audience")
    raw_content = _prompt("Paste any source content to base it on (optional, blank to skip)", default="")

    return {
        "title": title,
        "topic": topic,
        "slide_count": int(slide_count) if str(slide_count).isdigit() else 10,
        "language": language,
        "detail_level": detail_level if detail_level in DETAIL_LEVELS else "medium",
        "audience": audience,
        "raw_content": raw_content,
    }


def flags_capture(args):
    raw_content = args.content or ""
    if args.content_file:
        with open(args.content_file, "r", encoding="utf-8") as f:
            raw_content = f.read()

    return {
        "title": args.title,
        "topic": args.topic,
        "slide_count": args.slides,
        "language": args.language,
        "detail_level": args.detail if args.detail in DETAIL_LEVELS else "medium",
        "audience": args.audience,
        "raw_content": raw_content,
    }


def missing_fields(request):
    """Used by the orchestrator to decide whether to ask a follow-up question."""
    required = ["title", "topic", "slide_count"]
    return [f for f in required if not request.get(f)]


def main():
    parser = argparse.ArgumentParser(description="Capture structured presentation request.")
    parser.add_argument("--title")
    parser.add_argument("--topic")
    parser.add_argument("--slides", type=int, default=10)
    parser.add_argument("--language", default="en")
    parser.add_argument("--detail", default="medium", choices=DETAIL_LEVELS)
    parser.add_argument("--audience", default="General audience")
    parser.add_argument("--content", default=None, help="Inline source content")
    parser.add_argument("--content-file", default=None, help="Path to a text file with source content")
    parser.add_argument("--output-dir", default="output")
    args = parser.parse_args()

    os.makedirs(args.output_dir, exist_ok=True)
    log(args.output_dir, "started")

    if args.title and args.topic:
        request = flags_capture(args)
        log(args.output_dir, "captured via flags (non-interactive)")
    else:
        request = interactive_capture()
        log(args.output_dir, "captured via interactive prompts")

    missing = missing_fields(request)
    if missing:
        log(args.output_dir, f"WARNING: missing required fields: {missing}")

    out_path = os.path.join(args.output_dir, "presentation_request.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(request, f, indent=2)

    log(args.output_dir, f"output file written: {out_path}")
    print(f"Wrote {out_path}")
    print(json.dumps(request, indent=2))


if __name__ == "__main__":
    main()
