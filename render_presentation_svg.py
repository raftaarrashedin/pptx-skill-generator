"""
render_presentation_svg.py

Deterministic, rule-based renderer (no LLM). Reads presentation_final.json +
template_context.json and converts each slide into one SVG file. If a slide
says "title + 3 bullets", the renderer places those elements in a predictable
spot every time -- no model variability here, on purpose.

Every text/shape/image element gets a `data-role` attribute (and a few
`data-*` fields for raw values). export_presentation_pptx.py reads these
attributes back to build native, editable PowerPoint shapes. This SVG is a
deliberately narrow, renderer-specific dialect -- not general-purpose SVG.

Output: output/slides_svg/slide_NN.svg (one per slide)

Usage:
    python render_presentation_svg.py
"""

import argparse
import html
import json
import os
import textwrap
from datetime import datetime, timezone

PX_PER_INCH = 96
MARGIN_IN = 0.6


def log(output_dir, message):
    path = os.path.join(output_dir, "run.log")
    with open(path, "a", encoding="utf-8") as f:
        f.write(f"{datetime.now(timezone.utc).isoformat()} | render_presentation_svg | {message}\n")


def _in(px_or_in_value):
    return px_or_in_value * PX_PER_INCH


def _esc(text):
    return html.escape(str(text), quote=True)


def _wrap(text, width_chars):
    """Naive char-based wrap -- good enough for a deterministic v1 renderer."""
    return textwrap.wrap(text, width=width_chars) or [""]


class SvgCanvas:
    """Tiny helper to accumulate SVG element strings with data-role tagging."""

    def __init__(self, width_px, height_px, bg_color):
        self.width = width_px
        self.height = height_px
        self.elements = [
            f'<rect data-role="background" x="0" y="0" width="{width_px}" height="{height_px}" '
            f'fill="#{bg_color}"/>'
        ]

    def rect(self, x, y, w, h, fill, role="shape", extra=""):
        self.elements.append(
            f'<rect data-role="{role}" x="{x:.1f}" y="{y:.1f}" width="{w:.1f}" height="{h:.1f}" '
            f'fill="#{fill}" {extra}/>'
        )

    def text(self, x, y, content, size, color, role="text", weight="normal", anchor="start", family="Calibri", data_attrs=None):
        attrs = ""
        if data_attrs:
            attrs = " ".join(f'data-{k}="{_esc(v)}"' for k, v in data_attrs.items())
        self.elements.append(
            f'<text data-role="{role}" {attrs} x="{x:.1f}" y="{y:.1f}" font-size="{size}" '
            f'font-family="{family}" font-weight="{weight}" fill="#{color}" text-anchor="{anchor}">'
            f"{_esc(content)}</text>"
        )

    def multiline(self, x, y, lines, size, color, role, line_height=None, weight="normal", family="Calibri", role_prefix=None):
        line_height = line_height or size * 1.4
        for i, line in enumerate(lines):
            r = f"{role_prefix}-{i}" if role_prefix else role
            self.text(x, y + i * line_height, line, size, color, role=r, weight=weight, family=family)

    def to_svg(self):
        body = "\n  ".join(self.elements)
        return f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {self.width} {self.height}">\n  {body}\n</svg>\n'


def render_slide(slide: dict, colors: dict, fonts: dict, width_in: float, height_in: float) -> str:
    w, h = _in(width_in), _in(height_in)
    margin = _in(MARGIN_IN)
    canvas = SvgCanvas(w, h, colors["background"])
    title_font = fonts.get("title", "Cambria")
    body_font = fonts.get("body", "Calibri")
    text_color = colors.get("text", "212121")
    primary = colors.get("primary", "1E2761")
    secondary = colors.get("secondary", "CADCFC")

    stype = slide.get("type")
    title = slide.get("title") or ""

    if stype == "title":
        canvas.rect(0, 0, w, h, primary, role="background")
        canvas.text(margin, h * 0.42, title, 44, "FFFFFF", role="title", weight="bold", family=title_font)
        if slide.get("subtitle"):
            canvas.text(margin, h * 0.42 + 50, slide["subtitle"], 20, secondary, role="subtitle", family=body_font)

    elif stype == "section":
        canvas.rect(0, 0, w, h, primary, role="background")
        canvas.text(w / 2, h / 2, title, 40, "FFFFFF", role="title", weight="bold", anchor="middle", family=title_font)
        if slide.get("subtitle"):
            canvas.text(w / 2, h / 2 + 44, slide["subtitle"], 18, secondary, role="subtitle", anchor="middle", family=body_font)

    elif stype == "bullets":
        canvas.text(margin, margin + 40, title, 32, text_color, role="title", weight="bold", family=title_font)
        y = margin + 100
        for i, bullet in enumerate(slide.get("bullets", [])):
            lines = _wrap(bullet, 80)
            canvas.text(margin, y, f"\u2022 {lines[0]}", 18, text_color, role=f"bullet-{i}", family=body_font)
            for cont in lines[1:]:
                y += 26
                canvas.text(margin + 24, y, cont, 18, text_color, role=f"bullet-{i}-cont", family=body_font)
            y += 40

    elif stype == "metrics":
        canvas.text(margin, margin + 40, title, 32, text_color, role="title", weight="bold", family=title_font)
        metrics = slide.get("metrics", []) or [{"label": "", "value": ""}]
        col_w = (w - 2 * margin) / max(len(metrics), 1)
        for i, m in enumerate(metrics):
            cx = margin + col_w * i + col_w / 2
            canvas.text(cx, h * 0.55, str(m.get("value", "")), 48, primary, role=f"metric-value-{i}", weight="bold", anchor="middle", family=title_font)
            canvas.text(cx, h * 0.55 + 36, str(m.get("label", "")), 16, text_color, role=f"metric-label-{i}", anchor="middle", family=body_font)

    elif stype == "comparison":
        canvas.text(margin, margin + 40, title, 32, text_color, role="title", weight="bold", family=title_font)
        comp = slide.get("comparison") or {"left": {}, "right": {}}
        col_w = (w - 2 * margin - 40) / 2
        left_x = margin
        right_x = margin + col_w + 40
        for side_x, side_key in ((left_x, "left"), (right_x, "right")):
            side = comp.get(side_key) or {}
            canvas.text(side_x, margin + 110, side.get("heading", ""), 22, primary, role=f"{side_key}-heading", weight="bold", family=title_font)
            y = margin + 150
            for i, item in enumerate(side.get("items", [])):
                canvas.text(side_x, y, f"\u2022 {item}", 16, text_color, role=f"{side_key}-item-{i}", family=body_font)
                y += 30

    elif stype == "chart":
        canvas.text(margin, margin + 40, title, 32, text_color, role="title", weight="bold", family=title_font)
        chart = slide.get("chart") or {"categories": [], "series": []}
        categories = chart.get("categories", [])
        series = chart.get("series", [])
        chart_x, chart_y = margin, margin + 100
        chart_w, chart_h = w - 2 * margin, h - chart_y - margin
        canvas.rect(chart_x, chart_y, chart_w, chart_h, "FFFFFF", role="chart-area")
        all_values = [v for s in series for v in s.get("values", [])] or [1]
        max_val = max(all_values) or 1
        n = max(len(categories), 1)
        bar_group_w = chart_w / n
        n_series = max(len(series), 1)
        for ci, cat in enumerate(categories):
            for si, s in enumerate(series):
                value = s.get("values", [0] * n)[ci] if ci < len(s.get("values", [])) else 0
                bar_w = bar_group_w / (n_series + 1)
                bar_h = (value / max_val) * (chart_h - 30)
                bx = chart_x + ci * bar_group_w + si * bar_w + bar_w / 2
                by = chart_y + chart_h - bar_h - 20
                canvas.rect(bx, by, bar_w * 0.8, bar_h, primary, role=f"chart-bar-{ci}-{si}", extra=f'data-category="{_esc(cat)}" data-series="{_esc(s.get("name",""))}" data-value="{value}"')
            canvas.text(chart_x + ci * bar_group_w + bar_group_w / 2, chart_y + chart_h - 4, cat, 12, text_color, role=f"chart-cat-{ci}", anchor="middle", family=body_font)

    elif stype == "image_text":
        canvas.text(margin, margin + 40, title, 32, text_color, role="title", weight="bold", family=title_font)
        img_w = (w - 2 * margin) * 0.45
        img_h = h - margin - 130
        canvas.rect(margin, margin + 90, img_w, img_h, secondary, role="image-placeholder",
                    extra=f'data-label="{_esc((slide.get("image") or {}).get("placeholder_label",""))}" '
                          f'data-path="{_esc((slide.get("image") or {}).get("path") or "")}"')
        canvas.text(margin + img_w / 2, margin + 90 + img_h / 2, (slide.get("image") or {}).get("placeholder_label", "Image"),
                    14, text_color, role="image-label", anchor="middle", family=body_font)
        text_x = margin + img_w + 40
        y = margin + 110
        for i, bullet in enumerate(slide.get("bullets", [])):
            for j, line in enumerate(_wrap(bullet, 50)):
                prefix = "\u2022 " if j == 0 else "  "
                canvas.text(text_x, y, f"{prefix}{line}", 16, text_color, role=f"bullet-{i}", family=body_font)
                y += 24
            y += 16

    else:
        canvas.text(margin, margin + 40, title or f"Unsupported slide type: {stype}", 28, text_color, role="title", family=body_font)

    if slide.get("notes"):
        canvas.elements.append(f'<metadata data-role="speaker-notes">{_esc(slide["notes"])}</metadata>')

    return canvas.to_svg()


def render_all(final_draft: dict, template_context: dict, output_dir: str) -> list:
    colors = template_context.get("colors", {})
    fonts = template_context.get("fonts", {})
    width_in = template_context.get("slide_width_in", 13.333)
    height_in = template_context.get("slide_height_in", 7.5)

    svg_dir = os.path.join(output_dir, "slides_svg")
    os.makedirs(svg_dir, exist_ok=True)

    written = []
    for slide in final_draft.get("slides", []):
        svg = render_slide(slide, colors, fonts, width_in, height_in)
        fname = f"slide_{slide['id']:02d}.svg"
        fpath = os.path.join(svg_dir, fname)
        with open(fpath, "w", encoding="utf-8") as f:
            f.write(svg)
        written.append(fpath)
    return written


def main():
    parser = argparse.ArgumentParser(description="Render presentation_final.json to SVG slides.")
    parser.add_argument("--output-dir", default="output")
    args = parser.parse_args()

    os.makedirs(args.output_dir, exist_ok=True)
    log(args.output_dir, "started")

    final_path = os.path.join(args.output_dir, "presentation_final.json")
    ctx_path = os.path.join(args.output_dir, "template_context.json")
    for p in (final_path, ctx_path):
        if not os.path.exists(p):
            log(args.output_dir, f"ERROR: missing input {p}")
            raise SystemExit(f"Missing input: {p}. Approve a draft first (review_presentation_draft.py --approve).")

    with open(final_path, encoding="utf-8") as f:
        final_draft = json.load(f)
    with open(ctx_path, encoding="utf-8") as f:
        template_context = json.load(f)
    log(args.output_dir, f"input file used: {final_path}, {ctx_path}")

    try:
        written = render_all(final_draft, template_context, args.output_dir)
    except Exception as e:
        log(args.output_dir, f"ERROR: {e}")
        raise

    for fpath in written:
        log(args.output_dir, f"output file written: {fpath}")

    print(f"Rendered {len(written)} slide(s) to {os.path.join(args.output_dir, 'slides_svg')}")
    for fpath in written:
        print(f"  {fpath}")


if __name__ == "__main__":
    main()
