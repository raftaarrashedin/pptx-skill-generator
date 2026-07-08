"""
extract_presentation_template.py

Extract a reusable template package into `actor/<template-name>/`.

The package stores:
- the source template pptx
- extracted design context
- embedded image assets
- a reserved `svgs/` directory for template-derived SVG artifacts
"""

import argparse
import json
import os
import shutil

from template_actor import (
    package_assets_dir,
    package_context_path,
    package_dir_for_template,
    package_exists,
    package_source_pptx_path,
    package_svgs_dir,
    reset_package_dir,
    write_package_context,
)
from template_catalog import resolve_template_reference, template_name_from_path
from template_context_builder import build_context


def extract_template_package(
    template_reference: str, force: bool = False
) -> tuple[str, dict, bool]:
    template_path = resolve_template_reference(template_reference)
    package_dir = package_dir_for_template(template_path)

    if package_exists(package_dir) and not force:
        with open(package_context_path(package_dir), encoding="utf-8") as f:
            return package_dir, json.load(f), True

    reset_package_dir(package_dir)

    cached_template_path = package_source_pptx_path(package_dir)
    shutil.copyfile(template_path, cached_template_path)

    context = build_context(cached_template_path, package_dir)
    original_basename = os.path.basename(template_path)
    context["source_template"] = original_basename
    context["summary"] = context.get("summary", "").replace(
        "source_template.pptx", original_basename
    )
    context["source_template_name"] = template_name_from_path(template_path)
    context["source_template_path"] = cached_template_path
    context["actor_package_dir"] = package_dir
    context["template_svgs_dir"] = package_svgs_dir(package_dir)
    context["template_assets_dir"] = package_assets_dir(package_dir)

    write_package_context(package_dir, context)
    return package_dir, context, False


def main():
    parser = argparse.ArgumentParser(
        description="Extract and cache a PowerPoint template package."
    )
    parser.add_argument(
        "--template", required=True, help="Template path or saved template name"
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Re-extract even if a cached package already exists",
    )
    args = parser.parse_args()

    package_dir, context, reused = extract_template_package(
        args.template, force=args.force
    )
    status = "Reused" if reused else "Extracted"
    print(f"{status} template package: {package_dir}")
    print(json.dumps(context, indent=2))


if __name__ == "__main__":
    main()
