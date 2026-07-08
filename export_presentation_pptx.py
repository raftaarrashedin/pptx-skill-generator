"""
export_presentation_pptx.py

Mechanical conversion step (no LLM). Reads the SVG slides produced by
render_presentation_svg.py and rebuilds them as native, editable PowerPoint
shapes -- text boxes, slide background fill, image placeholders, and real
embedded chart objects (not flattened images) for chart slides.

This is a narrow parser tailored to the SVG dialect render_presentation_svg.py
produces (data-role / data-* attributes) -- it does not handle arbitrary SVG.
By the time you reach this step all content decisions are already made, so
this is meant to be a dumb, predictable, mechanical pass.

Output: output/final_presentation.pptx

Usage:
    python export_presentation_pptx.py
"""

import argparse
import glob
import os
import xml.etree.ElementTree as ET
from datetime import datetime, timezone

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Emu, Pt

PX_PER_INCH = 96  # must match render_presentation_svg.py


def log(output_dir, message):
    path = os.path.join(output_dir, "run.log")
    with open(path, "a", encoding="utf-8") as f:
        f.write(
            f"{datetime.now(timezone.utc).isoformat()} | export_presentation_pptx | {message}\n"
        )


def px_to_emu(px: float) -> int:
    return int(px / PX_PER_INCH * 914400)


def px_to_pt(px: float) -> float:
    return px * 0.75  # 96 px/in, 72 pt/in


def _tag(elem):
    return elem.tag.split("}")[-1]


def _rgb(hex_color: str) -> RGBColor:
    hex_color = (hex_color or "212121").lstrip("#")
    return RGBColor.from_string(hex_color.upper())


def _remove_all_slides(prs):
    slide_ids = list(prs.slides._sldIdLst)
    for slide_id in slide_ids:
        rel_id = slide_id.rId
        prs.part.drop_rel(rel_id)
        prs.slides._sldIdLst.remove(slide_id)


def _select_layout(prs):
    for layout in prs.slide_layouts:
        if "blank" in (layout.name or "").lower():
            return layout
    return prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]


def add_textbox(slide, elem, width_px, height_px):
    text = elem.text or ""
    if not text.strip():
        return
    x = float(elem.get("x", 0))
    y = float(elem.get("y", 0))
    size_px = float(elem.get("font-size", 16))
    color = elem.get("fill", "#212121")
    weight = elem.get("font-weight", "normal")
    family = elem.get("font-family", "Calibri")
    anchor = elem.get("text-anchor", "start")

    size_pt = px_to_pt(size_px)
    box_h = px_to_emu(size_px * 1.6)
    # Baseline -> box top approximation.
    top = px_to_emu(max(y - size_px * 1.1, 0))

    est_width_px = max(len(text) * size_px * 0.62, size_px * 2)
    if anchor == "middle":
        left = px_to_emu(x - est_width_px / 2)
        box_w = px_to_emu(est_width_px)
        align = PP_ALIGN.CENTER
    else:
        left = px_to_emu(x)
        box_w = px_to_emu(max(width_px - x, est_width_px))
        align = PP_ALIGN.LEFT

    box = slide.shapes.add_textbox(left, top, max(box_w, px_to_emu(20)), box_h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size_pt)
    run.font.bold = weight == "bold"
    run.font.name = family
    run.font.color.rgb = _rgb(color)


def set_background(slide, elem):
    fill_color = elem.get("fill", "#FFFFFF")
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = _rgb(fill_color)


def add_image_placeholder(slide, elem, output_dir):
    x = float(elem.get("x", 0))
    y = float(elem.get("y", 0))
    w = float(elem.get("width", 0))
    h = float(elem.get("height", 0))
    path = elem.get("data-path") or ""
    label = elem.get("data-label") or "Image placeholder"

    full_path = os.path.join(output_dir, path) if path else None
    if full_path and os.path.exists(full_path):
        slide.shapes.add_picture(
            full_path, px_to_emu(x), px_to_emu(y), px_to_emu(w), px_to_emu(h)
        )
        return

    shape = slide.shapes.add_shape(
        1, px_to_emu(x), px_to_emu(y), px_to_emu(w), px_to_emu(h)
    )  # 1 = RECTANGLE
    shape.fill.solid()
    shape.fill.fore_color.rgb = _rgb(elem.get("fill", "#CADCFC"))
    shape.line.fill.background()
    tf = shape.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = label
    run.font.size = Pt(12)
    run.font.color.rgb = _rgb("404040")


def add_chart(slide, slide_elements, width_px, height_px, output_dir):
    """Collect chart-bar-* rects and rebuild as a native, editable pptx chart."""
    bars = [
        e
        for e in slide_elements
        if _tag(e) == "rect" and (e.get("data-role") or "").startswith("chart-bar-")
    ]
    if not bars:
        return

    categories_order = []
    series_map = {}
    for bar in bars:
        cat = bar.get("data-category", "")
        ser = bar.get("data-series", "series")
        val = float(bar.get("data-value", 0))
        if cat not in categories_order:
            categories_order.append(cat)
        series_map.setdefault(ser, {})[cat] = val

    chart_data = CategoryChartData()
    chart_data.categories = categories_order
    for ser_name, values_by_cat in series_map.items():
        chart_data.add_series(
            ser_name, [values_by_cat.get(c, 0) for c in categories_order]
        )

    margin_px = 0.6 * PX_PER_INCH
    chart_x, chart_y = margin_px, margin_px + 100
    chart_w, chart_h = width_px - 2 * margin_px, height_px - chart_y - margin_px

    slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        px_to_emu(chart_x),
        px_to_emu(chart_y),
        px_to_emu(chart_w),
        px_to_emu(chart_h),
        chart_data,
    )


def add_notes(slide, slide_elements):
    for elem in slide_elements:
        if _tag(elem) == "metadata" and elem.get("data-role") == "speaker-notes":
            slide.notes_slide.notes_text_frame.text = elem.text or ""


def build_slide(prs, svg_path, output_dir):
    tree = ET.parse(svg_path)
    root = tree.getroot()
    viewbox = root.get("viewBox", "0 0 1280 720").split()
    width_px, height_px = float(viewbox[2]), float(viewbox[3])

    blank_layout = _select_layout(prs)
    slide = prs.slides.add_slide(blank_layout)

    elements = list(root)

    has_chart_bars = any(
        _tag(e) == "rect" and (e.get("data-role") or "").startswith("chart-bar-")
        for e in elements
    )

    for elem in elements:
        tag = _tag(elem)
        role = elem.get("data-role", "")
        if tag == "rect" and role == "background":
            set_background(slide, elem)
        elif tag == "rect" and role == "image-placeholder":
            add_image_placeholder(slide, elem, output_dir)
        elif tag == "rect" and role.startswith("chart-bar-"):
            continue  # handled once, below
        elif tag == "rect" and role == "chart-area":
            continue  # superseded by the native chart object
        elif tag == "text":
            if role.startswith("chart-cat-"):
                continue  # categories come from the native chart, not redundant labels
            if role == "image-label":
                continue  # already rendered inside the placeholder shape itself
            add_textbox(slide, elem, width_px, height_px)
        elif tag == "metadata":
            continue  # handled via add_notes

    if has_chart_bars:
        add_chart(slide, elements, width_px, height_px, output_dir)

    add_notes(slide, elements)
    return slide


def export(output_dir: str, template_context: dict) -> str:
    width_in = template_context.get("slide_width_in", 13.333)
    height_in = template_context.get("slide_height_in", 7.5)

    source_template_path = template_context.get("source_template_path")
    if source_template_path and os.path.exists(source_template_path):
        prs = Presentation(source_template_path)
        _remove_all_slides(prs)
    else:
        prs = Presentation()
    prs.slide_width = Emu(int(width_in * 914400))
    prs.slide_height = Emu(int(height_in * 914400))

    svg_dir = os.path.join(output_dir, "slides_svg")
    svg_files = sorted(glob.glob(os.path.join(svg_dir, "slide_*.svg")))
    if not svg_files:
        raise SystemExit(
            f"No rendered slides found in {svg_dir}. Run render_presentation_svg.py first."
        )

    for svg_path in svg_files:
        build_slide(prs, svg_path, output_dir)

    out_path = os.path.join(output_dir, "final_presentation.pptx")
    prs.save(out_path)
    return out_path


def main():
    parser = argparse.ArgumentParser(
        description="Export rendered SVG slides to an editable pptx."
    )
    parser.add_argument("--output-dir", default="output")
    args = parser.parse_args()

    os.makedirs(args.output_dir, exist_ok=True)
    log(args.output_dir, "started")

    ctx_path = os.path.join(args.output_dir, "template_context.json")
    if not os.path.exists(ctx_path):
        log(args.output_dir, f"ERROR: missing input {ctx_path}")
        raise SystemExit(f"Missing input: {ctx_path}")

    import json

    with open(ctx_path, encoding="utf-8") as f:
        template_context = json.load(f)

    log(
        args.output_dir,
        f"input file used: {os.path.join(args.output_dir, 'slides_svg')}, {ctx_path}",
    )

    try:
        out_path = export(args.output_dir, template_context)
    except Exception as e:
        log(args.output_dir, f"ERROR: {e}")
        raise

    log(args.output_dir, f"output file written: {out_path}")
    print(f"Wrote {out_path}")


if __name__ == "__main__":
    main()
