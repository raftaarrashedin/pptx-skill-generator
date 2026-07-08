"""
run_presentation_builder.py

Top-level entrypoint. Doesn't contain presentation logic itself -- it calls
the other scripts in sequence, in one place, so there's one spot to start
the system and one spot to see the full flow.

Ensures output/ exists before anything else runs. Every other script reads
from and writes to that folder only.

Subcommands map 1:1 to pipeline stages:

    python run_presentation_builder.py extract-template --template "Nature Journal"
    python run_presentation_builder.py template   [--template path.pptx]
    python run_presentation_builder.py templates
    python run_presentation_builder.py request     --title ... --topic ... --slides N   (or no flags = interactive)
    python run_presentation_builder.py orchestrate [--feedback "..."] [--mock]
    python run_presentation_builder.py review      [--approve]
    python run_presentation_builder.py revise      --feedback "..." [--mock]
    python run_presentation_builder.py render
    python run_presentation_builder.py export
    python run_presentation_builder.py edit         <edit_presentation.py args>

    # Or run the default end-to-end flow in one go (stops before export
    # unless you pass --auto-approve; always uses --mock unless --live is set):
    python run_presentation_builder.py all --title "..." --topic "..." --slides 8 [--live] [--auto-approve]
"""

import argparse
import os
import shutil
import sys

import capture_presentation_request
import export_presentation_pptx
import extract_presentation_template
import prepare_template_context
import render_presentation_svg
import review_presentation_draft
import revise_presentation_draft
import run_presentation_orchestrator


def reset_generated_outputs(output_dir):
    generated_files = (
        "presentation_draft.json",
        "presentation_final.json",
        "final_presentation.pptx",
        "preview.json",
        "review.txt",
    )
    generated_dirs = ("slides_svg", "assets", "template_svgs")

    for name in generated_files:
        path = os.path.join(output_dir, name)
        if os.path.exists(path):
            os.remove(path)

    for name in generated_dirs:
        path = os.path.join(output_dir, name)
        if os.path.isdir(path):
            shutil.rmtree(path)


def run_template(args):
    sys.argv = ["prepare_template_context.py", "--output-dir", args.output_dir]
    if args.template:
        sys.argv += ["--template", args.template]
    if getattr(args, "list_templates", False):
        sys.argv += ["--list-templates"]
    prepare_template_context.main()


def run_extract_template(args):
    sys.argv = ["extract_presentation_template.py", "--template", args.template]
    if getattr(args, "force", False):
        sys.argv += ["--force"]
    extract_presentation_template.main()


def run_request(args, extra):
    sys.argv = [
        "capture_presentation_request.py",
        "--output-dir",
        args.output_dir,
    ] + extra
    capture_presentation_request.main()


def run_orchestrate(args):
    sys.argv = ["run_presentation_orchestrator.py", "--output-dir", args.output_dir]
    if args.feedback:
        sys.argv += ["--feedback", args.feedback]
    if args.mock:
        sys.argv += ["--mock"]
    run_presentation_orchestrator.main()


def run_review(args):
    sys.argv = ["review_presentation_draft.py", "--output-dir", args.output_dir]
    if args.approve:
        sys.argv += ["--approve"]
    review_presentation_draft.main()


def run_revise(args):
    sys.argv = [
        "revise_presentation_draft.py",
        "--output-dir",
        args.output_dir,
        "--feedback",
        args.feedback,
    ]
    if args.mock:
        sys.argv += ["--mock"]
    revise_presentation_draft.main()


def run_render(args):
    sys.argv = ["render_presentation_svg.py", "--output-dir", args.output_dir]
    render_presentation_svg.main()


def run_export(args):
    sys.argv = ["export_presentation_pptx.py", "--output-dir", args.output_dir]
    export_presentation_pptx.main()


def run_edit(args, extra):
    import edit_presentation

    sys.argv = ["edit_presentation.py", "--output-dir", args.output_dir] + extra
    edit_presentation.main()


def run_all(args, extra):
    os.makedirs(args.output_dir, exist_ok=True)
    reset_generated_outputs(args.output_dir)
    mock = not args.live

    step = 1
    total_steps = 7 if args.template else 6

    if args.template:
        print(f"== {step}/{total_steps} extract template ==")
        sys.argv = ["extract_presentation_template.py", "--template", args.template]
        extract_presentation_template.main()
        step += 1

    print(f"== {step}/{total_steps} template ==")
    sys.argv = ["prepare_template_context.py", "--output-dir", args.output_dir]
    if args.template:
        sys.argv += ["--template", args.template]
    prepare_template_context.main()
    step += 1

    print(f"== {step}/{total_steps} request ==")
    request_argv = [
        "capture_presentation_request.py",
        "--output-dir",
        args.output_dir,
    ] + extra
    sys.argv = request_argv
    capture_presentation_request.main()
    step += 1

    print(f"== {step}/{total_steps} orchestrate (generate) ==")
    sys.argv = ["run_presentation_orchestrator.py", "--output-dir", args.output_dir]
    if mock:
        sys.argv += ["--mock"]
    run_presentation_orchestrator.main()
    step += 1

    print(f"== {step}/{total_steps} review ==")
    sys.argv = ["review_presentation_draft.py", "--output-dir", args.output_dir]
    if args.auto_approve:
        sys.argv += ["--approve"]
    review_presentation_draft.main()
    step += 1

    if not args.auto_approve:
        print(
            "\nStopped after review. Re-run with --auto-approve once you've checked output/review.txt,"
        )
        print("or approve manually: python review_presentation_draft.py --approve")
        return

    print(f"== {step}/{total_steps} render ==")
    sys.argv = ["render_presentation_svg.py", "--output-dir", args.output_dir]
    render_presentation_svg.main()
    step += 1

    print(f"== {step}/{total_steps} export ==")
    sys.argv = ["export_presentation_pptx.py", "--output-dir", args.output_dir]
    export_presentation_pptx.main()


def main():
    parser = argparse.ArgumentParser(
        description="Top-level presentation builder runner."
    )
    parser.add_argument("--output-dir", default="output")
    sub = parser.add_subparsers(dest="command", required=True)

    p_extract = sub.add_parser("extract-template")
    p_extract.add_argument("--template", required=True)
    p_extract.add_argument("--force", action="store_true")

    p_template = sub.add_parser("template")
    p_template.add_argument("--template", default=None)
    p_template.add_argument("--list-templates", action="store_true")

    sub.add_parser("templates")

    sub.add_parser("request")  # flags forwarded as-is via extra args

    p_orch = sub.add_parser("orchestrate")
    p_orch.add_argument("--feedback", default=None)
    p_orch.add_argument("--mock", action="store_true")

    p_review = sub.add_parser("review")
    p_review.add_argument("--approve", action="store_true")

    p_revise = sub.add_parser("revise")
    p_revise.add_argument("--feedback", required=True)
    p_revise.add_argument("--mock", action="store_true")

    sub.add_parser("render")
    sub.add_parser("export")
    sub.add_parser("edit")

    p_all = sub.add_parser("all")
    p_all.add_argument("--template", default=None)
    p_all.add_argument(
        "--live", action="store_true", help="Use a real LLM call instead of --mock"
    )
    p_all.add_argument(
        "--auto-approve",
        action="store_true",
        help="Skip the manual review pause and run through export",
    )

    # Split known top-level args from anything meant for the sub-script
    # (e.g. `request --title ... --topic ...` or `edit move-slide ...`).
    args, extra = parser.parse_known_args()

    os.makedirs(args.output_dir, exist_ok=True)

    if args.command == "extract-template":
        run_extract_template(args)
    elif args.command == "template":
        run_template(args)
    elif args.command == "templates":
        run_template(
            argparse.Namespace(
                output_dir=args.output_dir, template=None, list_templates=True
            )
        )
    elif args.command == "request":
        run_request(args, extra)
    elif args.command == "orchestrate":
        run_orchestrate(args)
    elif args.command == "review":
        run_review(args)
    elif args.command == "revise":
        run_revise(args)
    elif args.command == "render":
        run_render(args)
    elif args.command == "export":
        run_export(args)
    elif args.command == "edit":
        run_edit(args, extra)
    elif args.command == "all":
        run_all(args, extra)


if __name__ == "__main__":
    main()
