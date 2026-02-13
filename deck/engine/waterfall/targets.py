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

from ..pptx.text import _normalize_label
from .resolve import _normalize_target_label_value, _normalize_text_value

def target_brand_from_scope_df(scope_df):
    if scope_df is None or scope_df.empty:
        return None

    column_lookup = {str(col).strip().lower(): col for col in scope_df.columns}
    if "target brand" in column_lookup:
        series = scope_df[column_lookup["target brand"]].dropna()
        if not series.empty:
            return str(series.iloc[0])

    for _, row in scope_df.iterrows():
        if not len(row):
            continue
        label = str(row.iloc[0]).strip().lower()
        normalized_label = label.rstrip(":")
        if normalized_label == "target brand" and len(row) > 1 and pd.notna(row.iloc[1]):
            return str(row.iloc[1])

    return None


def modelled_category_from_scope_df(scope_df):
    if scope_df is None or scope_df.empty:
        return None

    if scope_df.shape[1] >= 2:
        for _, row in scope_df.iterrows():
            if not len(row):
                continue
            label = str(row.iloc[0]).strip()
            normalized_label = _normalize_label(label)
            if normalized_label == "category" and pd.notna(row.iloc[1]):
                return str(row.iloc[1])

    return None


def _normalize_column_name(value: object) -> str:
    if isinstance(value, (list, tuple, set)):
        value = " ".join(str(item) for item in value if item is not None)
    if value is None:
        return ""
    text = str(value).strip().lower()
    return "".join(ch for ch in text if ch.isalnum())


def _find_sheet_by_candidates(sheet_names: list[str], target: str) -> str | None:
    normalized_target = _normalize_column_name(target)
    normalized_sheets = {
        _normalize_column_name(str(sheet_name)): sheet_name
        for sheet_name in sheet_names
    }
    if normalized_target in normalized_sheets:
        return normalized_sheets[normalized_target]
    for normalized_sheet, sheet_name in normalized_sheets.items():
        if normalized_target in normalized_sheet or normalized_sheet in normalized_target:
            return sheet_name
    from difflib import get_close_matches

    matches = get_close_matches(
        normalized_target,
        list(normalized_sheets.keys()),
        n=1,
        cutoff=0.7,
    )
    if matches:
        return normalized_sheets[matches[0]]
    return None


def _find_column_by_candidates(df: pd.DataFrame, candidates: list[str]):
    normalized_columns = {_normalize_column_name(str(col)): col for col in df.columns}
    candidate_normalized = [_normalize_column_name(candidate) for candidate in candidates]
    for candidate in candidate_normalized:
        if candidate in normalized_columns:
            return normalized_columns[candidate]
    for column_key, column_name in normalized_columns.items():
        for candidate in candidate_normalized:
            if candidate in column_key or column_key in candidate:
                return column_name
    from difflib import get_close_matches

    matches = get_close_matches(
        " ".join(candidate_normalized),
        list(normalized_columns.keys()),
        n=1,
        cutoff=0.75,
    )
    if matches:
        return normalized_columns[matches[0]]
    return None


def _find_column_by_row_values(row: pd.Series, candidates: list[str]):
    normalized_values = {}
    for column, value in row.items():
        if pd.isna(value):
            continue
        normalized_values[_normalize_column_name(str(value))] = column
    if not normalized_values:
        return None

    candidate_normalized = [_normalize_column_name(candidate) for candidate in candidates]
    for candidate in candidate_normalized:
        if candidate in normalized_values:
            return normalized_values[candidate]
    for value_key, column_name in normalized_values.items():
        for candidate in candidate_normalized:
            if candidate in value_key or value_key in candidate:
                return column_name
    from difflib import get_close_matches

    matches = get_close_matches(
        " ".join(candidate_normalized),
        list(normalized_values.keys()),
        n=1,
        cutoff=0.75,
    )
    if matches:
        return normalized_values[matches[0]]
    return None


def _is_target_flag(value):
    if pd.isna(value):
        return False
    try:
        return float(value) == 1
    except (TypeError, ValueError):
        text = str(value).strip().lower()
        return text in {"1", "yes", "y", "true", "t"}


def _target_values_from_scope(
    scope_df: pd.DataFrame,
    target_col_candidates: list[str],
    value_col_candidates: list[str],
):
    if scope_df is None or scope_df.empty:
        return None
    target_col = _find_column_by_candidates(scope_df, target_col_candidates)
    value_col = _find_column_by_candidates(scope_df, value_col_candidates)
    if not target_col or not value_col:
        return None

    values = []
    seen = set()
    for _, row in scope_df.iterrows():
        if not _is_target_flag(row[target_col]):
            continue
        value = row[value_col]
        if pd.isna(value):
            continue
        name = str(value).strip()
        if not name or name in seen:
            continue
        seen.add(name)
        values.append(name)
    return values or None


def target_brands_from_scope_df(scope_df: pd.DataFrame):
    return _target_values_from_scope(
        scope_df,
        ["Target Brand", "Target_Brand"],
        ["Brand", "Brand Name"],
    )


def target_manufacturers_from_scope_df(scope_df: pd.DataFrame):
    return _target_values_from_scope(
        scope_df,
        ["Target Manufacturer", "Target_Manufacturer", "Target Mfr", "Target Mfg"],
        ["Manufacturer", "Mfr", "Mfg"],
    )


def target_brands_from_product_description(product_df: pd.DataFrame):
    if product_df is None or product_df.empty:
        return None
    target_col = _find_column_by_candidates(product_df, ["Target Brand", "Target_Brand"])
    brand_col = _find_column_by_candidates(product_df, ["Brand"])
    if not target_col or not brand_col:
        return None

    brands = []
    seen = set()
    for _, row in product_df.iterrows():
        if not _is_target_flag(row[target_col]):
            continue
        brand_value = row[brand_col]
        if pd.isna(brand_value):
            continue
        brand_name = str(brand_value).strip()
        if not brand_name or brand_name in seen:
            continue
        seen.add(brand_name)
        brands.append(brand_name)
    return brands or None


def target_dimensions_from_product_description(product_df: pd.DataFrame) -> list[str]:
    if product_df is None or product_df.empty:
        return []

    normalized_columns = {
        _normalize_column_name(str(col)): col for col in product_df.columns
    }
    lines = []
    seen_dimensions = set()
    for column in product_df.columns:
        column_name = str(column)
        if not column_name.strip().lower().startswith("target"):
            continue
        base_name = column_name.strip()[len("target"):].lstrip(" _-").strip()
        if not base_name:
            continue
        base_key = _normalize_column_name(base_name)
        if base_key in seen_dimensions:
            continue
        base_column = normalized_columns.get(base_key) or _find_column_by_candidates(
            product_df, [base_name]
        )
        if not base_column:
            continue

        values = []
        seen_values = set()
        for _, row in product_df.iterrows():
            if not _is_target_flag(row[column]):
                continue
            value = row[base_column]
            if pd.isna(value):
                continue
            value_name = str(value).strip()
            if not value_name or value_name in seen_values:
                continue
            seen_values.add(value_name)
            values.append(value_name)

        if values:
            base_label = str(base_column).strip()
            lines.append(f"Target {base_label}(s): {', '.join(values)}")
            seen_dimensions.add(base_key)

    return lines


def target_lines_from_product_description(product_df: pd.DataFrame) -> list[str]:
    return target_dimensions_from_product_description(product_df)


def _target_level_labels_from_gathered_df(gathered_df: pd.DataFrame) -> list[str]:
    if gathered_df is None or gathered_df.empty:
        return []
    label_col, data_start_idx = _target_level_label_column_exact(gathered_df)
    if not label_col:
        raise ValueError("The gatheredCN10 file is missing the Target Level Label column.")
    labels = (
        gathered_df.iloc[data_start_idx:][label_col]
        .dropna()
        .astype(str)
        .map(str.strip)
    )
    unique_labels = []
    seen = set()
    for label in labels:
        if not label or label in seen:
            continue
        seen.add(label)
        unique_labels.append(label)
    return unique_labels


def _target_level_label_column_exact(gathered_df: pd.DataFrame) -> tuple[str | None, int]:
    if gathered_df is None or gathered_df.empty:
        return None, 0
    column = _find_column_by_candidates(
        gathered_df,
        ["Target Level Label", "Target Level"],
    )
    if column:
        return column, 0
    header_row = gathered_df.iloc[0]
    column = _find_column_by_row_values(header_row, ["Target Level Label", "Target Level"])
    if column:
        return column, 1
    return None, 0


def _target_level_labels_from_gathered_df_with_filters(
    gathered_df: pd.DataFrame,
    year1: str | None = None,
    year2: str | None = None,
    target_labels: list[str] | None = None,
) -> list[str]:
    if gathered_df is None or gathered_df.empty:
        return []
    label_col, data_start_idx = _target_level_label_column_exact(gathered_df)
    if not label_col:
        raise ValueError("The gatheredCN10 file is missing the Target Level Label column.")
    data_df = gathered_df.iloc[data_start_idx:]
    filtered_df = data_df

    year_col = _find_column_by_candidates(gathered_df, ["Year", "Model Year"])
    if year_col and (year1 is not None or year2 is not None):
        normalized_years = {
            _normalize_text_value(value)
            for value in (year1, year2)
            if value is not None
        }
        if normalized_years:
            year_series = data_df[year_col].map(_normalize_text_value)
            filtered_df = filtered_df[year_series.isin(normalized_years)]

    target_label_col = _find_column_by_candidates(
        gathered_df, ["Target Label", "Target", "Target Type"]
    )
    normalized_targets: set[str] = set()
    for label in target_labels or []:
        normalized = _normalize_text_value(label)
        if not normalized:
            continue
        normalized_targets.add(normalized)
        if normalized == "cross":
            normalized_targets.add("competitor")
        if normalized == "competitor":
            normalized_targets.add("cross")
    if target_label_col and normalized_targets:
        target_series = data_df[target_label_col].map(_normalize_text_value)
        filtered_df = filtered_df[target_series.isin(normalized_targets)]

    labels = (
        filtered_df[label_col]
        .dropna()
        .astype(str)
        .map(str.strip)
    )
    unique_labels = []
    seen = set()
    for label in labels:
        if not label or label in seen:
            continue
        seen.add(label)
        unique_labels.append(label)
    return unique_labels


def _target_label_values_from_gathered_df(gathered_df: pd.DataFrame) -> list[str]:
    if gathered_df is None or gathered_df.empty:
        return []
    header_row = gathered_df.iloc[0] if len(gathered_df) else None
    column = _find_column_by_candidates(
        gathered_df, ["Target Label", "Target", "Target Type"]
    )
    data_start_idx = 0
    if not column and header_row is not None:
        column = _find_column_by_row_values(header_row, ["Target Label", "Target", "Target Type"])
        if column:
            data_start_idx = 1
    if not column:
        return []
    labels = (
        gathered_df.iloc[data_start_idx:][column]
        .dropna()
        .astype(str)
        .map(str.strip)
    )
    unique_labels = []
    seen = set()
    for label in labels:
        normalized = _normalize_target_label_value(label)
        if not normalized or normalized in seen:
            continue
        seen.add(normalized)
        unique_labels.append(label)
    return unique_labels
