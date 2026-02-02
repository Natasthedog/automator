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


def _normalize_project_details_label(text: object) -> str:
    if pd.isna(text):
        return ""
    normalized = str(text).strip().lower()
    normalized = " ".join(normalized.split())
    normalized = normalized.rstrip(" :;,.!?")
    return normalized


def _project_detail_value_from_df(
    project_details_df: pd.DataFrame | None,
    label_key: str,
    synonyms: list[str],
    canonical: str,
):
    if project_details_df is None or project_details_df.empty or project_details_df.shape[1] < 2:
        return None

    synonym_matches = []
    normalized_synonyms = {_normalize_project_details_label(item) for item in synonyms}
    for row_idx, row in project_details_df.iterrows():
        cell_value = row.iloc[0]
        normalized_cell = _normalize_project_details_label(cell_value)
        if not normalized_cell:
            continue
        if normalized_cell in normalized_synonyms:
            synonym_matches.append((row_idx, cell_value, normalized_cell, 1.0))

    if synonym_matches:
        candidates = synonym_matches
    else:
        candidates = []
        canonical_norm = _normalize_project_details_label(canonical)
        for row_idx, row in project_details_df.iterrows():
            cell_value = row.iloc[0]
            normalized_cell = _normalize_project_details_label(cell_value)
            if not normalized_cell:
                continue
            score = SequenceMatcher(None, normalized_cell, canonical_norm).ratio()
            if score >= 0.85:
                candidates.append((row_idx, cell_value, normalized_cell, score))

    if not candidates:
        raise ValueError(f"Could not find Project Details label for {label_key}")

    candidates.sort(key=lambda item: item[3], reverse=True)
    if len(candidates) > 1 and candidates[1][3] >= candidates[0][3] - 0.03:
        details = ", ".join(
            f"{str(item[1]).strip()} ({item[3]:.2f})" for item in candidates
        )
        raise ValueError(
            f"Ambiguous Project Details label for {label_key}. Candidates: {details}"
        )

    row_idx, original_label, _, _ = candidates[0]
    logger.info(
        "Matched Project Details label for %s: %s", label_key, str(original_label)
    )
    raw_value = project_details_df.iloc[row_idx, 1]
    if pd.isna(raw_value):
        return ""
    return str(raw_value).strip()
