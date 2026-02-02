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

from .charts import _chart_title_text_frame

def _find_slide_by_marker(prs, marker_text: str):
    marker_text = marker_text.strip()
    for slide in prs.slides:
        for shape in slide.shapes:
            shape_name = getattr(shape, "name", "") or ""
            if marker_text and marker_text in shape_name:
                return slide
            if shape.has_text_frame and marker_text in shape.text_frame.text:
                return slide
    return None


def _shape_text_snippet(shape) -> str:
    if not shape.has_text_frame:
        return ""
    text = shape.text_frame.text or ""
    compact = " ".join(text.split())
    return compact[:80]


def _slide_title(slide) -> str:
    try:
        title_shape = slide.shapes.title
    except Exception:
        title_shape = None
    if title_shape is not None and title_shape.has_text_frame:
        title_text = title_shape.text_frame.text or ""
        return title_text.strip()
    return ""


def _slide_index(prs, target_slide) -> int | None:
    for idx, slide in enumerate(prs.slides, start=1):
        if slide is target_slide:
            return idx
    return None


def _slides_with_placeholder(prs, placeholder: str) -> list[int]:
    matches: list[int] = []
    for idx, slide in enumerate(prs.slides, start=1):
        found = False
        for shape in slide.shapes:
            if shape.has_text_frame and placeholder in (shape.text_frame.text or ""):
                found = True
                break
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        if placeholder in (cell.text_frame.text or ""):
                            found = True
                            break
                    if found:
                        break
            if found:
                break
            if shape.has_chart:
                chart_text_frame = _chart_title_text_frame(shape.chart)
                if chart_text_frame and placeholder in (chart_text_frame.text or ""):
                    found = True
                    break
        if found:
            matches.append(idx)
    return matches
