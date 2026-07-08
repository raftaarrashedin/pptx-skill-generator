"""
edit_presentation.py

Targeted post-generation edits. Operates directly on presentation_final.json
(or presentation_draft.json with --target draft) -- not a rerun of the whole
pipeline. v1 operations only:

    move-slide      --id N --before M      (move slide N to just before slide M)
    edit-title      --id N --title "..."
    remove-bullet   --id N --index I
    duplicate-slide --id N
    delete-slide    --id N

After any edit, slide ids are renumbered sequentially starting at 1.
Re-run render_presentation_svg.py + export_presentation_pptx.py afterwards
to see the change reflected in the .pptx -- this script does not re-render.

Usage:
    python edit_presentation.py move-slide --id 4 --before 2
    python edit_presentation.py edit-title --id 3 --title "New Title"
    python edit_presentation.py remove-bullet --id 2 --index 1
    python edit_presentation.py duplicate-slide --id 3
    python edit_presentation.py delete-slide --id 5
    python edit_presentation.py move-slide --id 4 --before 2 --target draft
"""

import argparse
import json
import os
from datetime import datetime, timezone


def log(output_dir, message):
    path = os.path.join(output_dir, "run.log")
    with open(path, "a", encoding="utf-8") as f:
        f.write(f"{datetime.now(timezone.utc).isoformat()} | edit_presentation | {message}\n")


def _renumber(slides):
    for i, slide in enumerate(slides, start=1):
        slide["id"] = i
    return slides


def _find_index(slides, slide_id):
    for i, s in enumerate(slides):
        if s["id"] == slide_id:
            return i
    raise SystemExit(f"No slide with id {slide_id}")


def op_move_slide(draft, slide_id, before_id):
    slides = draft["slides"]
    src_idx = _find_index(slides, slide_id)
    slide = slides.pop(src_idx)
    dest_idx = _find_index(slides, before_id)
    slides.insert(dest_idx, slide)
    draft["slides"] = _renumber(slides)
    return f"Moved slide {slide_id} to before slide {before_id}."


def op_edit_title(draft, slide_id, title):
    slides = draft["slides"]
    idx = _find_index(slides, slide_id)
    slides[idx]["title"] = title
    return f"Slide {slide_id} title set to '{title}'."


def op_remove_bullet(draft, slide_id, index):
    slides = draft["slides"]
    idx = _find_index(slides, slide_id)
    bullets = slides[idx].get("bullets", [])
    if not (0 <= index < len(bullets)):
        raise SystemExit(f"Slide {slide_id} has no bullet at index {index} (has {len(bullets)} bullets).")
    removed = bullets.pop(index)
    return f"Removed bullet {index} ('{removed}') from slide {slide_id}."


def op_duplicate_slide(draft, slide_id):
    slides = draft["slides"]
    idx = _find_index(slides, slide_id)
    duplicate = json.loads(json.dumps(slides[idx]))
    slides.insert(idx + 1, duplicate)
    draft["slides"] = _renumber(slides)
    return f"Duplicated slide {slide_id}; copy inserted right after it."


def op_delete_slide(draft, slide_id):
    slides = draft["slides"]
    idx = _find_index(slides, slide_id)
    slides.pop(idx)
    draft["slides"] = _renumber(slides)
    return f"Deleted slide {slide_id}."


def main():
    parser = argparse.ArgumentParser(description="Targeted edits on a presentation JSON file.")
    parser.add_argument("--output-dir", default="output")
    parser.add_argument("--target", choices=["final", "draft"], default="final",
                         help="Edit presentation_final.json (default) or presentation_draft.json")
    sub = parser.add_subparsers(dest="op", required=True)

    p_move = sub.add_parser("move-slide")
    p_move.add_argument("--id", type=int, required=True)
    p_move.add_argument("--before", type=int, required=True)

    p_title = sub.add_parser("edit-title")
    p_title.add_argument("--id", type=int, required=True)
    p_title.add_argument("--title", required=True)

    p_bullet = sub.add_parser("remove-bullet")
    p_bullet.add_argument("--id", type=int, required=True)
    p_bullet.add_argument("--index", type=int, required=True)

    p_dup = sub.add_parser("duplicate-slide")
    p_dup.add_argument("--id", type=int, required=True)

    p_del = sub.add_parser("delete-slide")
    p_del.add_argument("--id", type=int, required=True)

    args = parser.parse_args()

    os.makedirs(args.output_dir, exist_ok=True)
    log(args.output_dir, "started")

    fname = "presentation_final.json" if args.target == "final" else "presentation_draft.json"
    path = os.path.join(args.output_dir, fname)
    if not os.path.exists(path):
        log(args.output_dir, f"ERROR: missing input {path}")
        raise SystemExit(f"Missing input: {path}")

    with open(path, encoding="utf-8") as f:
        draft = json.load(f)
    log(args.output_dir, f"input file used: {path}")
    log(args.output_dir, f"operation: {args.op} {vars(args)}")

    if args.op == "move-slide":
        result = op_move_slide(draft, args.id, args.before)
    elif args.op == "edit-title":
        result = op_edit_title(draft, args.id, args.title)
    elif args.op == "remove-bullet":
        result = op_remove_bullet(draft, args.id, args.index)
    elif args.op == "duplicate-slide":
        result = op_duplicate_slide(draft, args.id)
    elif args.op == "delete-slide":
        result = op_delete_slide(draft, args.id)
    else:
        raise SystemExit(f"Unknown operation: {args.op}")

    with open(path, "w", encoding="utf-8") as f:
        json.dump(draft, f, indent=2)

    log(args.output_dir, f"output file written: {path}")
    log(args.output_dir, f"result: {result}")
    print(result)
    print(f"Wrote {path}. Re-run render_presentation_svg.py + export_presentation_pptx.py to see it in the .pptx.")


if __name__ == "__main__":
    main()
