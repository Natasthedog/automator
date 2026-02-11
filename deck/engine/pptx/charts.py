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


def update_or_add_column_chart(slide, chart_name, categories, series_dict):
    """
    If a chart with name=chart_name exists on the slide, update its data.
    Otherwise insert a new clustered column chart in a sensible spot.
    Charts produced here remain EDITABLE in PowerPoint.
    """
    chart_shape = None
    for shape in slide.shapes:
        if getattr(shape, "name", None) == chart_name:
            if shape.has_chart:
                chart_shape = shape
                break
            else:
                # Remove placeholder artifacts that aren't real charts
                sp = shape._element
                sp.getparent().remove(sp)

    cd = ChartData()
    cd.categories = categories
    for s_name, values in series_dict.items():
        cd.add_series(s_name, list(values))

    if chart_shape:
        # Replace data in existing chart (preserves template styling)
        chart_shape.chart.replace_data(cd)
        updated_wb = _load_chart_workbook(chart_shape.chart)
        from .chart_cache import _update_chart_label_caches

        _update_chart_label_caches(chart_shape.chart, updated_wb)
        return chart_shape
    else:
        # Fallback: repurpose the first chart on the slide if present.
        for shape in slide.shapes:
            if shape.has_chart:
                shape.chart.replace_data(cd)
                shape.name = chart_name
                updated_wb = _load_chart_workbook(shape.chart)
                from .chart_cache import _update_chart_label_caches

                _update_chart_label_caches(shape.chart, updated_wb)
                return shape
        # Remove any stale shapes with the target name before adding a new chart
        for shape in list(slide.shapes):
            if getattr(shape, "name", None) == chart_name:
                sp = shape._element
                sp.getparent().remove(sp)

        # Insert a new chart (fallback)
        left, top, width, height = Inches(1), Inches(2), Inches(8), Inches(4.5)
        chart_shape = slide.shapes.add_chart(
            XL_CHART_TYPE.COLUMN_CLUSTERED, left, top, width, height, cd
        )
        chart_shape.name = chart_name
        chart = chart_shape.chart
        # Light touch formatting; rely on template/theme for styling
        chart.has_legend = True
        updated_wb = _load_chart_workbook(chart)
        from .chart_cache import _update_chart_label_caches

        _update_chart_label_caches(chart, updated_wb)
        return chart


def update_or_add_waterfall_chart(slide, chart_name, categories, series_dict):
    """
    Update an existing waterfall chart or insert one if missing.
    """
    chart_shape = None
    for shape in slide.shapes:
        if getattr(shape, "name", None) == chart_name and shape.has_chart:
            chart_shape = shape
            break

    if chart_shape is None:
        for shape in slide.shapes:
            if shape.has_chart:
                chart_shape = shape
                shape.name = chart_name
                break

    cd = ChartData()
    cd.categories = categories
    for s_name, values in series_dict.items():
        cd.add_series(s_name, list(values))

    if chart_shape:
        chart_shape.chart.replace_data(cd)
        return chart_shape

    waterfall_type = getattr(XL_CHART_TYPE, "WATERFALL", XL_CHART_TYPE.COLUMN_STACKED)
    left, top, width, height = Inches(1), Inches(2), Inches(8), Inches(4.5)
    chart_shape = slide.shapes.add_chart(
        waterfall_type, left, top, width, height, cd
    )
    chart_shape.name = chart_name
    return chart_shape


def _chart_title_text_frame(chart):
    try:
        if chart.has_title:
            return chart.chart_title.text_frame
    except Exception:
        return None
    return None


def _waterfall_chart_from_slide(slide, chart_name: str):
    for shape in slide.shapes:
        if shape.has_chart and chart_name in (getattr(shape, "name", "") or ""):
            return shape.chart
    for shape in slide.shapes:
        if shape.has_chart:
            return shape.chart
    return None


def _waterfall_chart_shape_from_slide(slide, chart_name: str):
    for shape in slide.shapes:
        if shape.has_chart and chart_name in (getattr(shape, "name", "") or ""):
            return shape
    for shape in slide.shapes:
        if shape.has_chart:
            return shape
    return None


def _clone_chart_part(chart_part):
    package = chart_part.package
    new_partname = package.next_partname(chart_part.partname_template)
    new_chart_part = chart_part.__class__.load(
        new_partname,
        chart_part.content_type,
        package,
        chart_part.blob,
    )
    xlsx_part = chart_part.chart_workbook.xlsx_part
    if xlsx_part is not None:
        new_xlsx_part = EmbeddedXlsxPart.new(xlsx_part.blob, package)
        new_chart_part.chart_workbook.xlsx_part = new_xlsx_part
    return new_chart_part


def _ensure_unique_chart_parts_on_slide(slide, seen_partnames: set[str]) -> None:
    for shape in slide.shapes:
        if not shape.has_chart:
            continue
        chart_part = shape.chart.part
        partname = str(chart_part.partname)
        if partname in seen_partnames:
            new_chart_part = _clone_chart_part(chart_part)
            new_rid = shape.part.relate_to(new_chart_part, RT.CHART)
            chart_element = shape._element.graphic.graphicData.chart
            chart_element.set(qn("r:id"), new_rid)
            chart_part = shape.chart.part
            partname = str(chart_part.partname)
        seen_partnames.add(partname)


def _categories_from_chart(chart) -> list[str]:
    categories = []
    try:
        plot_categories = chart.plots[0].categories
    except Exception:
        plot_categories = []
    for category in plot_categories:
        label = getattr(category, "label", None)
        categories.append(str(label) if label is not None else str(category))
    return categories


def _load_chart_workbook(chart):
    xlsx_blob = chart.part.chart_workbook.xlsx_part.blob
    return load_workbook(io.BytesIO(xlsx_blob))


def _save_chart_workbook(chart, workbook) -> None:
    stream = io.BytesIO()
    workbook.save(stream)
    chart.part.chart_workbook.xlsx_part.blob = stream.getvalue()
