"""
template_catalog.py

Helpers for working with static PowerPoint templates stored in the project.

Templates can live either:
- at the project root, e.g. `Nature Journal.pptx`
- in a dedicated `templates/` directory

Callers can pass either a filesystem path or a human-friendly template name
such as "Nature Journal".
"""

from __future__ import annotations

import os
import re
from dataclasses import dataclass


@dataclass(frozen=True)
class TemplateInfo:
    name: str
    path: str


def _project_root() -> str:
    return os.path.dirname(os.path.abspath(__file__))


def _slugify(value: str) -> str:
    lowered = value.lower().strip()
    return re.sub(r"[^a-z0-9]+", "-", lowered).strip("-")


def template_slug(value: str) -> str:
    return _slugify(value)


def template_name_from_path(path: str) -> str:
    return os.path.splitext(os.path.basename(path))[0]


def _candidate_dirs(base_dir: str | None = None) -> list[str]:
    root = os.path.abspath(base_dir or _project_root())
    templates_dir = os.path.join(root, "templates")
    dirs = [root]
    if os.path.isdir(templates_dir):
        dirs.insert(0, templates_dir)
    return dirs


def list_templates(base_dir: str | None = None) -> list[TemplateInfo]:
    seen_paths: set[str] = set()
    templates: list[TemplateInfo] = []

    for directory in _candidate_dirs(base_dir):
        for entry in sorted(os.listdir(directory)):
            if not entry.lower().endswith(".pptx"):
                continue
            full_path = os.path.abspath(os.path.join(directory, entry))
            if full_path in seen_paths:
                continue
            seen_paths.add(full_path)
            templates.append(
                TemplateInfo(name=os.path.splitext(entry)[0], path=full_path)
            )

    templates.sort(key=lambda item: item.name.lower())
    return templates


def resolve_template_reference(
    reference: str | None, base_dir: str | None = None
) -> str | None:
    if not reference:
        return None

    if os.path.exists(reference):
        return os.path.abspath(reference)

    root = os.path.abspath(base_dir or _project_root())
    joined = os.path.join(root, reference)
    if os.path.exists(joined):
        return os.path.abspath(joined)

    normalized_ref = reference.strip().lower()
    slug_ref = _slugify(reference)

    for template in list_templates(root):
        filename = os.path.basename(template.path)
        candidates = {
            template.name.lower(),
            filename.lower(),
            _slugify(template.name),
            _slugify(filename),
        }
        if normalized_ref in candidates or slug_ref in candidates:
            return template.path

    available = (
        ", ".join(template.name for template in list_templates(root)) or "<none>"
    )
    raise FileNotFoundError(
        f"Template '{reference}' was not found. Use a path or one of the saved template names: {available}"
    )
