"""
template_actor.py

Persistent storage helpers for extracted PowerPoint templates.

Each template gets its own package under `actor/<template-slug>/` so we can
reuse extracted assets and context across presentation runs.
"""

from __future__ import annotations

import json
import os
import shutil

from template_catalog import template_name_from_path, template_slug

ACTOR_DIRNAME = "actor"


def project_root() -> str:
    return os.path.dirname(os.path.abspath(__file__))


def actor_root(base_dir: str | None = None) -> str:
    return os.path.join(os.path.abspath(base_dir or project_root()), ACTOR_DIRNAME)


def package_slug_for_template(template_path: str) -> str:
    return template_slug(template_name_from_path(template_path))


def package_dir_for_template(template_path: str, base_dir: str | None = None) -> str:
    return os.path.join(actor_root(base_dir), package_slug_for_template(template_path))


def package_context_path(package_dir: str) -> str:
    return os.path.join(package_dir, "template_context.json")


def package_source_pptx_path(package_dir: str) -> str:
    return os.path.join(package_dir, "source_template.pptx")


def package_assets_dir(package_dir: str) -> str:
    return os.path.join(package_dir, "assets")


def package_svgs_dir(package_dir: str) -> str:
    return os.path.join(package_dir, "svgs")


def package_exists(package_dir: str) -> bool:
    return (
        os.path.exists(package_context_path(package_dir))
        and os.path.exists(package_source_pptx_path(package_dir))
        and os.path.isdir(package_assets_dir(package_dir))
        and os.path.isdir(package_svgs_dir(package_dir))
    )


def package_has_svgs(package_dir: str) -> bool:
    svg_dir = package_svgs_dir(package_dir)
    return os.path.isdir(svg_dir) and any(name.lower().endswith(".svg") for name in os.listdir(svg_dir))


def reset_package_dir(package_dir: str) -> None:
    if os.path.isdir(package_dir):
        shutil.rmtree(package_dir)
    os.makedirs(package_assets_dir(package_dir), exist_ok=True)
    os.makedirs(package_svgs_dir(package_dir), exist_ok=True)


def load_package_context(package_dir: str) -> dict:
    with open(package_context_path(package_dir), encoding="utf-8") as f:
        return json.load(f)


def write_package_context(package_dir: str, context: dict) -> str:
    path = package_context_path(package_dir)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(context, f, indent=2)
    return path


def stage_package_into_output(package_dir: str, output_dir: str) -> dict:
    context = load_package_context(package_dir)

    src_assets = package_assets_dir(package_dir)
    dst_assets = os.path.join(output_dir, "assets")
    if os.path.isdir(dst_assets):
        shutil.rmtree(dst_assets)
    if os.path.isdir(src_assets):
        shutil.copytree(src_assets, dst_assets)

    src_svgs = package_svgs_dir(package_dir)
    dst_svgs = os.path.join(output_dir, "template_svgs")
    if os.path.isdir(dst_svgs):
        shutil.rmtree(dst_svgs)
    if os.path.isdir(src_svgs) and os.listdir(src_svgs):
        shutil.copytree(src_svgs, dst_svgs)

    staged = dict(context)
    staged["source_template_path"] = package_source_pptx_path(package_dir)
    staged["actor_package_dir"] = package_dir
    staged["actor_package_name"] = os.path.basename(package_dir)
    staged["template_svgs_dir"] = src_svgs
    staged["images"] = [
        os.path.join("assets", os.path.basename(path))
        for path in context.get("images", [])
    ]
    return staged
