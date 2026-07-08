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
import json
import os
import xml.etree.ElementTree as ET
from datetime import datetime, timezone

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import PP_PLACEHOLDER
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


def _find_layout_by_name(prs, *names):
    wanted = {name.lower() for name in names}
    for layout in prs.slide_layouts:
        if (layout.name or "").lower() in wanted:
            return layout
    return None


def _template_can_drive_layouts(prs):
    useful = {"title", "section_header", "title_and_body", "title_and_two_columns"}
    return any((layout.name or "").lower() in useful for layout in prs.slide_layouts)


def _remove_shape(shape):
    shape._element.getparent().remove(shape._element)


def _placeholder_idx(shape):
    if not getattr(shape, "is_placeholder", False):
        return None
    return shape.placeholder_format.idx


def _find_placeholder(slide, placeholder_type=None, idx=None):
    for shape in slide.placeholders:
        if idx is not None and shape.placeholder_format.idx != idx:
            continue
        if placeholder_type is not None and shape.placeholder_format.type != placeholder_type:
            continue
        return shape
    return None


def _find_all_placeholders(slide, placeholder_type=None):
    found = []
    for shape in slide.placeholders:
        if placeholder_type is not None and shape.placeholder_format.type != placeholder_type:
            continue
        found.append(shape)
    return found


def _set_text(shape, text):
    if shape is None or not getattr(shape, "has_text_frame", False):
        return
    tf = shape.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = text or ""


def _set_bullets(shape, bullets):
    if shape is None or not getattr(shape, "has_text_frame", False):
        return
    tf = shape.text_frame
    tf.clear()
    items = bullets or [""]
    for i, bullet in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = bullet
        p.level = 0


def _set_lines(shape, lines):
    if shape is None or not getattr(shape, "has_text_frame", False):
        return
    tf = shape.text_frame
    tf.clear()
    items = lines or [""]
    for i, line in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.level = 0


def _body_geometry(shape, prs):
    if shape is None:
        margin = Emu(int(0.8 * 914400))
        return margin, margin * 2, prs.slide_width - margin * 2, prs.slide_height - margin * 3
    return shape.left, shape.top, shape.width, shape.height


def _cleanup_unused_placeholders(slide, used_idxs):
    removable_types = {
        PP_PLACEHOLDER.BODY,
        PP_PLACEHOLDER.PICTURE,
        PP_PLACEHOLDER.SUBTITLE,
        PP_PLACEHOLDER.OBJECT,
    }
    for shape in list(slide.placeholders):
        idx = shape.placeholder_format.idx
        ptype = shape.placeholder_format.type
        if idx in used_idxs:
            continue
        if ptype in removable_types:
            _remove_shape(shape)


def _slide_layout_for_type(prs, slide_type):
    mapping = {
        "title": ("title", "title_only"),
        "section": ("section_header", "section_title_and_description", "main_point"),
        "bullets": ("title_and_body", "title_and_two_columns", "caption_only"),
        "metrics": ("title_and_two_columns", "title_and_body", "main_point"),
        "comparison": ("title_and_two_columns", "title_and_body"),
        "chart": ("title_and_body", "title_and_two_columns"),
        "image_text": ("title_and_two_columns", "title_and_body", "title_only"),
    }
    names = mapping.get(slide_type, ("title_and_body", "blank"))
    return _find_layout_by_name(prs, *names) or _select_layout(prs)


def _add_chart_in_box(slide, chart, left, top, width, height):
    chart = chart or {"categories": [], "series": []}
    chart_data = CategoryChartData()
    chart_data.categories = chart.get("categories", [])
    for series in chart.get("series", []):
        chart_data.add_series(series.get("name", "Series"), series.get("values", []))
    slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        left,
        top,
        width,
        height,
        chart_data,
    )


def _add_image_in_box(slide, image_spec, output_dir, left, top, width, height):
    image_spec = image_spec or {}
    path = image_spec.get("path") or ""
    full_path = os.path.join(output_dir, path) if path else None
    if full_path and os.path.exists(full_path):
        slide.shapes.add_picture(full_path, left, top, width, height)
        return True
    return False


def _populate_template_slide(slide, prs, draft_slide, output_dir):
    used_idxs = set()

    title_shape = _find_placeholder(slide, idx=0) or _find_placeholder(slide, PP_PLACEHOLDER.TITLE)
    if title_shape is not None:
        _set_text(title_shape, draft_slide.get("title") or "")
        used_idxs.add(_placeholder_idx(title_shape))

    subtitle_text = draft_slide.get("subtitle") or ""
    subtitle_shape = _find_placeholder(slide, PP_PLACEHOLDER.SUBTITLE)
    body_shapes = _find_all_placeholders(slide, PP_PLACEHOLDER.BODY)
    picture_shapes = _find_all_placeholders(slide, PP_PLACEHOLDER.PICTURE)

    slide_type = draft_slide.get("type")
    if slide_type == "title":
        if subtitle_shape is not None:
            _set_text(subtitle_shape, subtitle_text)
            used_idxs.add(_placeholder_idx(subtitle_shape))

    elif slide_type == "section":
        if subtitle_shape is not None and subtitle_text:
            _set_text(subtitle_shape, subtitle_text)
            used_idxs.add(_placeholder_idx(subtitle_shape))
        elif body_shapes and subtitle_text:
            _set_text(body_shapes[0], subtitle_text)
            used_idxs.add(_placeholder_idx(body_shapes[0]))

    elif slide_type == "bullets":
        target = body_shapes[0] if body_shapes else subtitle_shape
        _set_bullets(target, draft_slide.get("bullets", []))
        if target is not None:
            used_idxs.add(_placeholder_idx(target))

    elif slide_type == "metrics":
        metrics = draft_slide.get("metrics", [])
        metric_lines = [f"{item.get('label', '')}: {item.get('value', '')}" for item in metrics]
        if len(body_shapes) >= 2:
            midpoint = max(1, (len(metric_lines) + 1) // 2)
            _set_lines(body_shapes[0], metric_lines[:midpoint])
            _set_lines(body_shapes[1], metric_lines[midpoint:])
            used_idxs.add(_placeholder_idx(body_shapes[0]))
            used_idxs.add(_placeholder_idx(body_shapes[1]))
        else:
            target = body_shapes[0] if body_shapes else subtitle_shape
            _set_lines(target, metric_lines)
            if target is not None:
                used_idxs.add(_placeholder_idx(target))

    elif slide_type == "comparison":
        comp = draft_slide.get("comparison") or {}
        left = comp.get("left") or {}
        right = comp.get("right") or {}
        left_lines = [left.get("heading", "")] + [f"- {item}" for item in left.get("items", [])]
        right_lines = [right.get("heading", "")] + [f"- {item}" for item in right.get("items", [])]
        if len(body_shapes) >= 2:
            _set_lines(body_shapes[0], left_lines)
            _set_lines(body_shapes[1], right_lines)
            used_idxs.add(_placeholder_idx(body_shapes[0]))
            used_idxs.add(_placeholder_idx(body_shapes[1]))
        else:
            target = body_shapes[0] if body_shapes else subtitle_shape
            _set_lines(target, left_lines + [""] + right_lines)
            if target is not None:
                used_idxs.add(_placeholder_idx(target))

    elif slide_type == "chart":
        body_shape = body_shapes[0] if body_shapes else None
        left, top, width, height = _body_geometry(body_shape, prs)
        if body_shape is not None:
            _remove_shape(body_shape)
        _add_chart_in_box(slide, draft_slide.get("chart"), left, top, width, height)

    elif slide_type == "image_text":
        bullets_target = body_shapes[0] if body_shapes else subtitle_shape
        _set_bullets(bullets_target, draft_slide.get("bullets", []))
        if bullets_target is not None:
            used_idxs.add(_placeholder_idx(bullets_target))

        picture_target = picture_shapes[0] if picture_shapes else None
        if picture_target is not None:
            used_idxs.add(_placeholder_idx(picture_target))
            left, top, width, height = _body_geometry(picture_target, prs)
            _remove_shape(picture_target)
            _add_image_in_box(slide, draft_slide.get("image"), output_dir, left, top, width, height)

    if draft_slide.get("notes"):
        slide.notes_slide.notes_text_frame.text = draft_slide["notes"]

    _cleanup_unused_placeholders(slide, {idx for idx in used_idxs if idx is not None})


def build_template_slide(prs, draft_slide, output_dir):
    layout = _slide_layout_for_type(prs, draft_slide.get("type"))
    slide = prs.slides.add_slide(layout)
    _populate_template_slide(slide, prs, draft_slide, output_dir)
    return slide


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


def export_from_svg(output_dir: str, template_context: dict) -> str:
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


def export_from_template(output_dir: str, template_context: dict, final_draft: dict) -> str:
    source_template_path = template_context.get("source_template_path")
    prs = Presentation(source_template_path)
    _remove_all_slides(prs)

    width_in = template_context.get("slide_width_in", 13.333)
    height_in = template_context.get("slide_height_in", 7.5)
    prs.slide_width = Emu(int(width_in * 914400))
    prs.slide_height = Emu(int(height_in * 914400))

    for draft_slide in final_draft.get("slides", []):
        build_template_slide(prs, draft_slide, output_dir)

    out_path = os.path.join(output_dir, "final_presentation.pptx")
    prs.save(out_path)
    return out_path


def export(output_dir: str, template_context: dict, final_draft: dict | None = None) -> str:
    source_template_path = template_context.get("source_template_path")
    if source_template_path and os.path.exists(source_template_path):
        prs = Presentation(source_template_path)
        if final_draft and _template_can_drive_layouts(prs):
            return export_from_template(output_dir, template_context, final_draft)
    return export_from_svg(output_dir, template_context)


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

    with open(ctx_path, encoding="utf-8") as f:
        template_context = json.load(f)
    final_draft = None
    final_path = os.path.join(args.output_dir, "presentation_final.json")
    if os.path.exists(final_path):
        with open(final_path, encoding="utf-8") as f:
            final_draft = json.load(f)

    log(
        args.output_dir,
        f"input file used: {os.path.join(args.output_dir, 'slides_svg')}, {ctx_path}, {final_path}",
    )

    try:
        out_path = export(args.output_dir, template_context, final_draft=final_draft)
    except Exception as e:
        log(args.output_dir, f"ERROR: {e}")
        raise

    log(args.output_dir, f"output file written: {out_path}")
    print(f"Wrote {out_path}")


if __name__ == "__main__":
    main()
