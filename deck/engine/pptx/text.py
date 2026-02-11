from __future__ import annotations

# Auto-generated split from dash_app.py

import io
import base64
import logging
import copy
from difflib import SequenceMatcher
from dataclasses import dataclass, asdict, is_dataclass
from datetime import date, timedelta
from pathlib import Path
import re
import json
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import range_boundaries, get_column_letter
import numbers
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.text import PP_ALIGN
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.oxml.ns import qn
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.parts.embeddedpackage import EmbeddedXlsxPart
from lxml import etree

logger = logging.getLogger(__name__)

def _chart_title_text_frame(chart):
    from .charts import _chart_title_text_frame as _inner

    return _inner(chart)

from .slides import _find_slide_by_marker, _shape_text_snippet, _slide_index, _slide_title, _slides_with_placeholder

def set_text_by_name(slide, shape_name, text):
    for shape in slide.shapes:
        if getattr(shape, "name", None) == shape_name and shape.has_text_frame:
            tf = shape.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            run = p.add_run()
            run.text = str(text)
            p.alignment = PP_ALIGN.LEFT
            return True
    return False


def replace_text_in_slide(slide, old_text, new_text):
    replaced = False
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        text_frame = shape.text_frame
        current_text = text_frame.text
        if old_text in current_text:
            text_frame.text = current_text.replace(old_text, new_text)
            replaced = True
    return replaced


def replace_text_in_slide_preserve_formatting(slide, old_text, new_text):
    replaced = False
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for paragraph in shape.text_frame.paragraphs:
            found_in_runs = False
            for run in paragraph.runs:
                if old_text in run.text:
                    run.text = run.text.replace(old_text, new_text)
                    found_in_runs = True
                    replaced = True
            if found_in_runs or old_text not in paragraph.text:
                continue
            updated_text = paragraph.text.replace(old_text, new_text)
            _rebuild_paragraph_runs(paragraph, updated_text)
            replaced = True
    return replaced


def _replace_placeholder_in_paragraph_runs(paragraph, placeholder: str, replacement: str) -> bool:
    if not replacement:
        return False
    runs = list(paragraph.runs)
    if not runs:
        return False
    replaced = False
    while True:
        full_text = "".join(run.text for run in runs)
        start_idx = full_text.find(placeholder)
        if start_idx == -1:
            break
        end_idx = start_idx + len(placeholder)
        replaced = True
        first_overlap = True
        cumulative = 0
        for run in runs:
            run_text = run.text
            run_start = cumulative
            run_end = cumulative + len(run_text)
            cumulative = run_end
            if run_end <= start_idx or run_start >= end_idx:
                continue
            overlap_start = max(start_idx, run_start)
            overlap_end = min(end_idx, run_end)
            local_start = overlap_start - run_start
            local_end = overlap_end - run_start
            if first_overlap:
                run.text = run_text[:local_start] + replacement + run_text[local_end:]
                first_overlap = False
            else:
                run.text = run_text[:local_start] + run_text[local_end:]
    return replaced


def _replace_placeholders_in_slide_runs(
    slide, replacements: dict[str, str | None]
) -> bool:
    replaced = False
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for paragraph in shape.text_frame.paragraphs:
            for placeholder, value in replacements.items():
                if not value:
                    continue
                if _replace_placeholder_in_paragraph_runs(paragraph, placeholder, value):
                    replaced = True
    return replaced


def _capture_run_formatting(run):
    font = run.font
    color = font.color
    return {
        "name": font.name,
        "size": font.size,
        "bold": font.bold,
        "italic": font.italic,
        "underline": font.underline,
        "color_rgb": color.rgb if color is not None else None,
    }


def _apply_run_formatting(run, formatting):
    font = run.font
    font.name = formatting["name"]
    font.size = formatting["size"]
    font.bold = formatting["bold"]
    font.italic = formatting["italic"]
    font.underline = formatting["underline"]
    if formatting["color_rgb"] is not None:
        font.color.rgb = formatting["color_rgb"]


def _rebuild_paragraph_runs(paragraph, new_text: str) -> None:
    original_runs = list(paragraph.runs)
    if not original_runs:
        paragraph.text = new_text
        return
    formats = [_capture_run_formatting(run) for run in original_runs]
    run_lengths = [len(run.text) for run in original_runs]
    for run in original_runs:
        paragraph._element.remove(run._r)
    cursor = 0
    for idx, fmt in enumerate(formats):
        if idx == len(formats) - 1:
            segment = new_text[cursor:]
        else:
            segment = new_text[cursor : cursor + run_lengths[idx]]
        new_run = paragraph.add_run()
        new_run.text = segment
        _apply_run_formatting(new_run, fmt)
        cursor += len(segment)


def _replace_placeholders_in_text_frame(text_frame, replacements, counts) -> None:
    for paragraph in text_frame.paragraphs:
        for placeholder, replacement in replacements.items():
            paragraph_text = paragraph.text or ""
            occurrences = paragraph_text.count(placeholder)
            if occurrences == 0:
                continue
            counts[placeholder]["found"] += occurrences
            if replacement is None:
                continue
            if paragraph.runs:
                if _replace_placeholder_in_paragraph_runs(paragraph, placeholder, replacement):
                    counts[placeholder]["replaced"] += occurrences
            else:
                paragraph.text = paragraph_text.replace(placeholder, replacement)
                counts[placeholder]["replaced"] += occurrences


def replace_placeholders_strict(prs, slide_selector, replacements: dict[str, str | None]) -> None:
    if slide_selector is None:
        raise ValueError("Slide selector is required to replace placeholders.")
    if isinstance(slide_selector, str):
        slide = _find_slide_by_marker(prs, slide_selector)
    else:
        slide = slide_selector
    if slide is None:
        raise ValueError(f"Could not resolve slide selector: {slide_selector}")

    counts = {
        placeholder: {"found": 0, "replaced": 0} for placeholder in replacements
    }
    for shape in slide.shapes:
        if shape.has_text_frame:
            _replace_placeholders_in_text_frame(shape.text_frame, replacements, counts)
        if shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    _replace_placeholders_in_text_frame(
                        cell.text_frame, replacements, counts
                    )
        if shape.has_chart:
            chart_text_frame = _chart_title_text_frame(shape.chart)
            if chart_text_frame is not None:
                _replace_placeholders_in_text_frame(
                    chart_text_frame, replacements, counts
                )

    slide_idx = _slide_index(prs, slide)
    slide_title = _slide_title(slide)
    slide_name = getattr(slide, "name", None) or ""
    shape_lines = []
    for shape in slide.shapes:
        shape_lines.append(
            " - "
            f"id={getattr(shape, 'shape_id', None)} "
            f"name={getattr(shape, 'name', None)!r} "
            f"type={getattr(shape, 'shape_type', None)} "
            f"has_text_frame={shape.has_text_frame} "
            f"has_table={shape.has_table} "
            f"text={_shape_text_snippet(shape)!r}"
        )
        if shape.has_chart:
            chart_text_frame = _chart_title_text_frame(shape.chart)
            if chart_text_frame is not None:
                chart_text = chart_text_frame.text or ""
                compact = " ".join(chart_text.split())
                shape_lines.append(
                    " - "
                    f"chart_title={compact[:80]!r} "
                    f"shape_name={getattr(shape, 'name', None)!r}"
                )
    counts_lines = [
        f" - {placeholder}: found={stats['found']} replaced={stats['replaced']}"
        for placeholder, stats in counts.items()
    ]

    def build_diagnostics(missing_placeholder: str) -> str:
        locations = _slides_with_placeholder(prs, missing_placeholder)
        location_line = (
            f"Slides containing {missing_placeholder}: {locations}"
            if locations
            else f"Slides containing {missing_placeholder}: []"
        )
        return "\n".join(
            [
                "Placeholder replacement diagnostics:",
                f"Slide index: {slide_idx}",
                f"Slide name: {slide_name}",
                f"Slide title: {slide_title}",
                "Shape inventory:",
                *shape_lines,
                "Replacement counts:",
                *counts_lines,
                location_line,
            ]
        )

    for placeholder in replacements:
        if counts[placeholder]["found"] > 0:
            continue
        locations = _slides_with_placeholder(prs, placeholder)
        if not locations:
            raise ValueError(
                f"Placeholder {placeholder} not found in deck\n"
                f"{build_diagnostics(placeholder)}"
            )
        intended_idx = slide_idx if slide_idx is not None else "unknown"
        raise ValueError(
            f"Placeholder {placeholder} found on slide {locations[0]} "
            f"not on Waterfall slide {intended_idx}\n"
            f"{build_diagnostics(placeholder)}"
        )


def append_text_after_label(slide, label_text, appended_text):
    if not appended_text:
        return False
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for paragraph in shape.text_frame.paragraphs:
            if label_text not in paragraph.text:
                continue
            if appended_text in paragraph.text:
                return True
            spacer = "" if label_text.endswith(" ") else " "
            for run in paragraph.runs:
                if label_text in run.text:
                    new_run = paragraph.add_run()
                    new_run.text = f"{spacer}{appended_text}"
                    return True
            new_run = paragraph.add_run()
            new_run.text = f"{spacer}{appended_text}"
            return True
    return False


def append_paragraph_after_label(slide, label_text, appended_text):
    if not appended_text:
        return False
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        text_frame = shape.text_frame
        if label_text not in text_frame.text:
            continue
        if any(paragraph.text.strip() == appended_text for paragraph in text_frame.paragraphs):
            return True
        for paragraph in text_frame.paragraphs:
            if label_text in paragraph.text:
                new_paragraph = text_frame.add_paragraph()
                new_paragraph.text = appended_text
                paragraph._p.addnext(new_paragraph._p)
                return True
    return False


def append_paragraphs_after_label(slide, label_text, appended_texts):
    if not appended_texts:
        return False
    appended_texts = [text for text in appended_texts if text and text.strip()]
    if not appended_texts:
        return False
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        text_frame = shape.text_frame
        if label_text not in text_frame.text:
            continue
        existing_texts = {paragraph.text.strip() for paragraph in text_frame.paragraphs}
        to_add = [text for text in appended_texts if text not in existing_texts]
        if not to_add:
            return True
        insert_after = None
        for paragraph in text_frame.paragraphs:
            if label_text in paragraph.text:
                insert_after = paragraph
                break
        if insert_after is None:
            continue
        last_paragraph = insert_after
        for text in to_add:
            new_paragraph = text_frame.add_paragraph()
            new_paragraph.text = text
            last_paragraph._p.addnext(new_paragraph._p)
            last_paragraph = new_paragraph
        return True
    return False


def add_table(slide, table_name, df: pd.DataFrame):
    # Identify an existing table to reuse, preferring one with the expected name.
    target_shape = None
    for shape in slide.shapes:
        if getattr(shape, "name", None) == table_name and shape.has_table:
            target_shape = shape
            break

    if target_shape is None:
        for shape in slide.shapes:
            if shape.has_table:
                target_shape = shape
                target_shape.name = table_name
                break

    if target_shape and target_shape.has_table:
        tbl = target_shape.table
        # Resize (simple): write headers to row 0, then rows afterward if room allows
        n_rows = min(len(df) + 1, tbl.rows.__len__())
        n_cols = min(len(df.columns), tbl.columns.__len__())
        # headers
        for j, col in enumerate(df.columns[:n_cols]):
            cell = tbl.cell(0, j)
            cell.text = str(col)
        # cells
        for i in range(1, n_rows):
            for j in range(n_cols):
                tbl.cell(i, j).text = str(df.iloc[i-1, j])
        # Clear any leftover rows beyond the populated range
        for i in range(n_rows, tbl.rows.__len__()):
            for j in range(tbl.columns.__len__()):
                tbl.cell(i, j).text = ""
        return True

    # Remove non-table placeholders with the desired name so we can insert a fresh table.
    for shape in list(slide.shapes):
        if getattr(shape, "name", None) == table_name and not shape.has_table:
            sp = shape._element
            sp.getparent().remove(sp)

    # Otherwise, add a new table
    rows, cols = len(df) + 1, len(df.columns)
    left, top, width, height = Inches(1), Inches(1.5), Inches(8), Inches(1 + 0.3 * rows)
    table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
    table_shape.name = table_name
    table = table_shape.table
    for j, col in enumerate(df.columns):
        table.cell(0, j).text = str(col)
    for i in range(len(df)):
        for j in range(len(df.columns)):
            table.cell(i+1, j).text = str(df.iloc[i, j])

    # Avoid manipulating the low-level XML that may not exist across templates.
    # python-pptx represents the table as a ``CT_GraphicalObjectFrame`` whose
    # schema does not expose a ``graphicFrame`` attribute.  Some versions of the
    # library can therefore raise an AttributeError when we try to clear borders
    # by touching ``graphicFrame`` directly.  Since this styling tweak is only a
    # nice-to-have, we simply rely on the template/theme defaults instead of
    # editing the XML manually.
    return True


def remove_empty_placeholders(slide):
    """Remove placeholder shapes that have no meaningful content."""
    for shape in list(slide.shapes):
        if not getattr(shape, "is_placeholder", False):
            continue

        # Keep placeholders that now contain text, tables, or charts with data.
        if shape.has_text_frame:
            if shape.text_frame.text and shape.text_frame.text.strip():
                continue
        elif shape.has_table:
            # If every cell is blank, treat as empty.
            if any(
                cell.text.strip()
                for row in shape.table.rows
                for cell in row.cells
            ):
                continue
        elif shape.has_chart:
            # Assume populated charts should remain.
            continue

        sp = shape._element
        sp.getparent().remove(sp)


def _normalize_label(value: str) -> str:
    return " ".join(value.strip().lower().replace(":", "").split())
