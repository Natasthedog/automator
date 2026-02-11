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

from ..waterfall.targets import _find_sheet_by_candidates, _normalize_column_name

logger = logging.getLogger(__name__)


_NUM_CACHE_WARNING_LIMIT = 10


_num_cache_warning_count = 0


def _chart_namespace_map(root) -> dict:
    nsmap = {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
             "a":   "http://schemas.openxmlformats.org/drawingml/2006/main",
             "r":   "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
             "c15": "http://schemas.microsoft.com/office/drawing/2012/chart"}
    for prefix, uri in (root.nsmap or {}).items():
        if prefix and uri:
            nsmap[prefix] = uri
    return nsmap


def _range_values_from_worksheet(ws, ref: str) -> list[list]:
    if not ref:
        return []
    normalized = str(ref)
    if "!" in normalized:
        _, normalized = normalized.split("!", 1)
    normalized = normalized.replace("$", "")
    try:
        min_col, min_row, max_col, max_row = range_boundaries(normalized)
    except ValueError:
        return []
    rows = []
    for row_idx in range(min_row, max_row + 1):
        row = []
        for col_idx in range(min_col, max_col + 1):
            row.append(ws.cell(row=row_idx, column=col_idx).value)
        rows.append(row)
    return rows


def _range_cells_from_worksheet(ws, ref: str) -> list[list]:
    if not ref:
        return []
    normalized = str(ref)
    if "!" in normalized:
        _, normalized = normalized.split("!", 1)
    normalized = normalized.replace("$", "")
    try:
        min_col, min_row, max_col, max_row = range_boundaries(normalized)
    except ValueError:
        return []
    rows = []
    for row_idx in range(min_row, max_row + 1):
        row = []
        for col_idx in range(min_col, max_col + 1):
            row.append(ws.cell(row=row_idx, column=col_idx))
        rows.append(row)
    return rows


def _worksheet_and_range_from_formula(workbook, formula: str) -> tuple:
    if not formula:
        return workbook.active, "", None
    sheet_name = None
    ref = str(formula)
    if "!" in ref:
        match = re.match(
            r"^(?P<sheet>(?:'[^']*(?:''[^']*)*'|[^!]+))!(?P<ref>.+)$",
            ref,
        )
        if match:
            sheet_part = match.group("sheet")
            ref = match.group("ref")
            if sheet_part.startswith("'") and sheet_part.endswith("'"):
                sheet_name = sheet_part[1:-1].replace("''", "'")
            else:
                sheet_name = sheet_part
        else:
            sheet_part, ref = ref.split("!", 1)
            sheet_name = sheet_part.strip("'")
    ref = ref.replace("$", "")
    ws = workbook.active
    if sheet_name:
        if sheet_name not in workbook.sheetnames:
            resolved = _find_sheet_by_candidates(workbook.sheetnames, sheet_name)
            if resolved:
                logger.info(
                    "Chart cache: resolved sheet '%s' -> '%s' from formula '%s'",
                    sheet_name,
                    resolved,
                    formula,
                )
                sheet_name = resolved
            else:
                raise ValueError(
                    f"Chart cache: sheet '{sheet_name}' from formula '{formula}' not found."
                )
        ws = workbook[sheet_name]
    return ws, ref, sheet_name


def _format_sheet_reference(sheet_name: str) -> str:
    if sheet_name is None:
        return ""
    if sheet_name.startswith("'") and sheet_name.endswith("'"):
        return sheet_name
    if re.search(r"[^A-Za-z0-9_]", sheet_name):
        return f"'{sheet_name}'"
    return sheet_name


def _build_cell_range_formula(sheet_name: str | None, col_idx: int, row_start: int, row_end: int) -> str:
    col_letter = get_column_letter(col_idx)
    sheet_prefix = ""
    if sheet_name:
        sheet_prefix = f"{_format_sheet_reference(sheet_name)}!"
    return f"{sheet_prefix}${col_letter}${row_start}:${col_letter}${row_end}"


def _range_boundaries_from_formula(formula: str) -> tuple[int, int, int, int] | None:
    if not formula:
        return None
    ref = str(formula)
    if "!" in ref:
        _, ref = ref.split("!", 1)
    ref = ref.replace("$", "")
    try:
        return range_boundaries(ref)
    except ValueError:
        return None


def _flatten_cell_values(values: list[list]) -> list:
    if not values:
        return []
    if len(values) == 1:
        return list(values[0])
    if all(len(row) == 1 for row in values):
        return [row[0] for row in values]
    flattened = []
    for row in values:
        flattened.extend(row)
    return flattened


def _log_num_cache_warning(
    value,
    fallback: float,
    sheet_name: str | None,
    cell_ref: str | None,
) -> None:
    global _num_cache_warning_count
    if _num_cache_warning_count >= _NUM_CACHE_WARNING_LIMIT:
        return
    _num_cache_warning_count += 1
    location = []
    if sheet_name:
        location.append(f"sheet={sheet_name}")
    if cell_ref:
        location.append(f"cell={cell_ref}")
    location_text = f" ({', '.join(location)})" if location else ""
    logger.warning(
        "Chart cache: coerced non-numeric value%s %r to %s",
        location_text,
        value,
        fallback,
    )


def safe_float(
    value,
    *,
    sheet_name: str | None = None,
    cell_ref: str | None = None,
) -> float:
    if value is None:
        _log_num_cache_warning(value, 0.0, sheet_name, cell_ref)
        return 0.0
    if isinstance(value, str):
        stripped = value.strip()
        if stripped == "":
            _log_num_cache_warning(value, 0.0, sheet_name, cell_ref)
            return 0.0
        if stripped.startswith("="):
            _log_num_cache_warning(value, 0.0, sheet_name, cell_ref)
            return 0.0
        normalized = stripped.replace(",", "")
        if re.search(r"[A-Za-z]", normalized):
            _log_num_cache_warning(value, 0.0, sheet_name, cell_ref)
            return 0.0
        if not re.fullmatch(r"[+-]?(?:\d+(\.\d*)?|\.\d+)", normalized):
            _log_num_cache_warning(value, 0.0, sheet_name, cell_ref)
            return 0.0
        try:
            return float(normalized)
        except (TypeError, ValueError):
            _log_num_cache_warning(value, 0.0, sheet_name, cell_ref)
            return 0.0
    if isinstance(value, numbers.Real):
        if pd.isna(value):
            _log_num_cache_warning(value, 0.0, sheet_name, cell_ref)
            return 0.0
        return float(value)
    try:
        converted = float(value)
    except (TypeError, ValueError):
        _log_num_cache_warning(value, 0.0, sheet_name, cell_ref)
        return 0.0
    if pd.isna(converted):
        _log_num_cache_warning(value, 0.0, sheet_name, cell_ref)
        return 0.0
    return converted


def _all_blank(values: list) -> bool:
    if not values:
        return True
    for value in values:
        if value is None:
            continue
        if isinstance(value, str) and value.strip() == "":
            continue
        return False
    return True


def _ensure_str_cache(str_ref) -> tuple:
    str_cache = str_ref.find("c:strCache", namespaces=_chart_namespace_map(str_ref))
    created = False
    if str_cache is None:
        str_cache = etree.SubElement(
            str_ref,
            "{http://schemas.openxmlformats.org/drawingml/2006/chart}strCache",
        )
        created = True
    return str_cache, created


def _update_num_cache(num_cache, values: list) -> None:
    if num_cache is None:
        return
    pt_count = num_cache.find("c:ptCount", namespaces=_chart_namespace_map(num_cache))
    if pt_count is None:
        pt_count = etree.SubElement(
            num_cache,
            "{http://schemas.openxmlformats.org/drawingml/2006/chart}ptCount",
        )
    pt_count.set("val", str(len(values)))
    for pt in list(num_cache.findall("c:pt", namespaces=_chart_namespace_map(num_cache))):
        num_cache.remove(pt)
    for idx, value in enumerate(values):
        cell = value if hasattr(value, "value") and hasattr(value, "coordinate") else None
        raw_value = value.value if cell is not None else value
        sheet_name = cell.parent.title if cell is not None else None
        cell_ref = cell.coordinate if cell is not None else None
        normalized_value = safe_float(
            raw_value,
            sheet_name=sheet_name,
            cell_ref=cell_ref,
        )
        pt = etree.SubElement(
            num_cache,
            "{http://schemas.openxmlformats.org/drawingml/2006/chart}pt",
            idx=str(idx),
        )
        v = etree.SubElement(
            pt, "{http://schemas.openxmlformats.org/drawingml/2006/chart}v"
        )
        v.text = str(normalized_value)


def _update_str_cache(str_cache, values: list[str]) -> None:
    if str_cache is None:
        return
    pt_count = str_cache.find("c:ptCount", namespaces=_chart_namespace_map(str_cache))
    if pt_count is None:
        pt_count = etree.SubElement(
            str_cache,
            "{http://schemas.openxmlformats.org/drawingml/2006/chart}ptCount",
        )
    pt_count.set("val", str(len(values)))
    for pt in list(str_cache.findall("c:pt", namespaces=_chart_namespace_map(str_cache))):
        str_cache.remove(pt)
    for idx, value in enumerate(values):
        pt = etree.SubElement(
            str_cache,
            "{http://schemas.openxmlformats.org/drawingml/2006/chart}pt",
            idx=str(idx),
        )
        v = etree.SubElement(
            pt, "{http://schemas.openxmlformats.org/drawingml/2006/chart}v"
        )
        v.text = "" if value is None else str(value)


def _update_c15_label_range_cache(
    container,
    formula: str | None,
    labels: list[str],
    nsmap: dict,
    label_context: str,
) -> int:
    c15_blocks = []
    if container is not None and str(getattr(container, "tag", "")).endswith("datalabelsRange"):
        c15_blocks.append(container)
    c15_blocks += container.findall(".//c15:datalabelsRange", namespaces=nsmap)
    if not c15_blocks:
        try:
            c15_blocks = container.xpath(".//*[local-name()='datalabelsRange']")
        except Exception:
            c15_blocks = []
    if c15_blocks:
        seen = set()
        deduped = []
        for block in c15_blocks:
            if id(block) in seen:
                continue
            seen.add(id(block))
            deduped.append(block)
        c15_blocks = deduped
    if not c15_blocks:
        logger.info(
            "Waterfall chart cache update: %s no c15 label-range block found",
            label_context,
        )
        return 0
    logger.info(
        "Waterfall chart cache update: %s c15 label-range blocks found %s",
        label_context,
        len(c15_blocks),
    )
    for c15_block in c15_blocks:
        block_ns = etree.QName(c15_block).namespace or nsmap.get(
            "c15", "http://schemas.microsoft.com/office/drawing/2012/chart"
        )
        f_node = c15_block.find("c15:f", namespaces=nsmap)
        if f_node is None:
            try:
                f_node = c15_block.xpath("./*[local-name()='f']")[0]
            except Exception:
                f_node = None
        if f_node is None:
            f_node = etree.SubElement(c15_block, f"{{{block_ns}}}f")
        if formula:
            f_node.text = formula
            logger.info(
                "Waterfall chart cache update: %s c15 label-range formula set to %s",
                label_context,
                formula,
            )
        cache = c15_block.find("c15:dlblRangeCache", namespaces=nsmap)
        if cache is None:
            try:
                cache = c15_block.xpath("./*[local-name()='dlblRangeCache']")[0]
            except Exception:
                cache = None
        if cache is None:
            cache = etree.SubElement(c15_block, f"{{{block_ns}}}dlblRangeCache")
        pt_count = cache.find("c15:ptCount", namespaces=nsmap)
        if pt_count is None:
            try:
                pt_count = cache.xpath("./*[local-name()='ptCount']")[0]
            except Exception:
                pt_count = None
        if pt_count is None:
            pt_count = etree.SubElement(cache, f"{{{block_ns}}}ptCount")
        pt_count.set("val", str(len(labels)))
        for pt in list(cache.findall("c15:pt", namespaces=nsmap)):
            cache.remove(pt)
        for pt in list(cache.xpath("./*[local-name()='pt']")):
            cache.remove(pt)
        for idx, value in enumerate(labels):
            pt = etree.SubElement(cache, f"{{{block_ns}}}pt", idx=str(idx))
            v = etree.SubElement(pt, f"{{{block_ns}}}v")
            v.text = "" if value is None else str(value)
        logger.info(
            "Waterfall chart cache update: %s c15 label-range cached %s points",
            label_context,
            len(labels),
        )
    return len(c15_blocks)


def _update_waterfall_chart_caches(chart, workbook, categories: list[str]) -> None:
    from ..waterfall.inject import _find_header_column, _resolve_waterfall_labs_column
    chart_part = chart.part
    root = chart_part._element
    nsmap = _chart_namespace_map(root)
    ws = workbook.active
    label_columns = {
        col_idx: ws.cell(row=1, column=col_idx).value
        for col_idx in range(1, ws.max_column + 1)
        if ws.cell(row=1, column=col_idx).value
        and _normalize_column_name(str(ws.cell(row=1, column=col_idx).value)).startswith("labs")
    }
    if label_columns:
        logger.info(
            "Waterfall chart cache update: label columns found %s",
            {idx: str(value) for idx, value in label_columns.items()},
        )
    categories_values = ["" if value is None else str(value) for value in categories]
    categories_count = len(categories_values)
    logger.info("Waterfall chart cache update: %s category points", categories_count)

    series_names = [series.name for series in chart.series]
    series_point_counts: dict[int, int] = {}
    series_category_bounds: dict[int, tuple[int, int, str | None]] = {}
    series_value_bounds: dict[int, tuple[int, int, str | None]] = {}

    for idx, ser in enumerate(root.findall(".//c:ser", namespaces=nsmap), start=1):
        num_ref = ser.find("c:val/c:numRef", namespaces=nsmap)
        if num_ref is None:
            continue
        f_node = num_ref.find("c:f", namespaces=nsmap)
        if f_node is None or not f_node.text:
            continue
        value_ws, value_ref, _ = _worksheet_and_range_from_formula(workbook, f_node.text)
        value_rows = _range_cells_from_worksheet(value_ws, value_ref)
        series_values = _flatten_cell_values(value_rows)
        num_cache = num_ref.find("c:numCache", namespaces=nsmap)
        _update_num_cache(num_cache, series_values)
        series_point_counts[idx] = len(series_values)
        bounds = _range_boundaries_from_formula(f_node.text)
        if bounds:
            _, min_row, _, max_row = bounds
            series_value_bounds[idx] = (min_row, max_row, value_ws.title)
        logger.info(
            "Waterfall chart cache update: series %s cached %s points",
            idx,
            len(series_values),
        )

    category_cache_updates = 0
    category_cache_count = None
    for idx, ser in enumerate(root.findall(".//c:ser", namespaces=nsmap), start=1):
        series_label = series_names[idx - 1] if idx - 1 < len(series_names) else f"Series {idx}"
        cat_node = ser.find("c:cat", namespaces=nsmap)
        if cat_node is None:
            logger.info(
                "Waterfall chart cache update: series %s category ref not found",
                series_label,
            )
            continue
        cat_ref = cat_node.find("c:strRef", namespaces=nsmap)
        cat_ref_type = "strRef"
        num_ref = None
        if cat_ref is None:
            num_ref = cat_node.find("c:numRef", namespaces=nsmap)
            cat_ref_type = "numRef"
            cat_ref = num_ref
        if cat_ref is None:
            logger.info(
                "Waterfall chart cache update: series %s category ref not found",
                series_label,
            )
            continue
        f_node = cat_ref.find("c:f", namespaces=nsmap)
        if f_node is None or not f_node.text:
            logger.info(
                "Waterfall chart cache update: series %s category ref formula missing",
                series_label,
            )
            continue
        logger.info(
            "Waterfall chart cache update: series %s category ref type %s formula %s",
            series_label,
            cat_ref_type,
            f_node.text,
        )
        if cat_ref_type == "numRef" and num_ref is not None:
            f_text = f_node.text
            num_ref_index = list(cat_node).index(num_ref)
            cat_node.remove(num_ref)
            cat_ref = etree.Element("{http://schemas.openxmlformats.org/drawingml/2006/chart}strRef")
            f_node = etree.SubElement(
                cat_ref, "{http://schemas.openxmlformats.org/drawingml/2006/chart}f"
            )
            f_node.text = f_text
            cat_node.insert(num_ref_index, cat_ref)
            cat_ref_type = "strRef"
        logger.info(
            "Waterfall chart cache update: series %s category ref formula %s",
            series_label,
            f_node.text,
        )
        cat_ws, cat_ref_range, cat_sheet = _worksheet_and_range_from_formula(
            workbook, f_node.text
        )
        category_rows = _range_values_from_worksheet(cat_ws, cat_ref_range)
        category_values = _flatten_cell_values(category_rows)
        if _all_blank(category_values):
            raise ValueError(
                f"Chart cache: category range '{f_node.text}' for series '{series_label}' is blank."
            )
        if not category_values and categories_values:
            category_values = categories_values
        category_values = ["" if value is None else str(value) for value in category_values]
        bounds = _range_boundaries_from_formula(f_node.text)
        if bounds:
            _, min_row, _, max_row = bounds
            series_category_bounds[idx] = (min_row, max_row, cat_sheet or cat_ws.title)
        str_cache, created = _ensure_str_cache(cat_ref)
        logger.info(
            "Waterfall chart cache update: series %s category strCache %s",
            series_label,
            "created" if created else "existing",
        )
        _update_str_cache(str_cache, category_values)
        category_cache_updates += 1
        category_cache_count = len(category_values)
        logger.info(
            "Waterfall chart cache update: series %s cached %s category points",
            series_label,
            len(category_values),
        )
    if category_cache_updates:
        logger.info(
            "Waterfall chart cache update: %s category cache points",
            category_cache_count if category_cache_count is not None else categories_count,
        )

    label_cache_updates = 0
    label_cache_missing = 0
    c15_label_updates = 0
    def _collect_label_refs():
        series_refs = []
        for idx, ser in enumerate(root.findall(".//c:ser", namespaces=nsmap), start=1):
            series_label = series_names[idx - 1] if idx - 1 < len(series_names) else f"Series {idx}"
            chart_series = chart.series[idx - 1] if idx - 1 < len(chart.series) else None
            label_refs = ser.findall(".//c:dLbls//c:dLbl//c:tx//c:strRef", namespaces=nsmap)
            label_refs += ser.findall(".//c:dLbls//c:tx//c:strRef", namespaces=nsmap)
            series_refs.append((idx, series_label, chart_series, label_refs, ser))
        plot_level_refs = root.findall(
            "c:plotArea//c:dLbls//c:tx//c:strRef", namespaces=nsmap
        )
        plot_level_dlbls = root.findall("c:plotArea/c:dLbls", namespaces=nsmap)
        plot_refs = []
        for ref_node in plot_level_refs:
            current = ref_node.getparent()
            has_series_ancestor = False
            while current is not None:
                if current.tag.endswith("ser"):
                    has_series_ancestor = True
                    break
                current = current.getparent()
            if not has_series_ancestor:
                plot_refs.append(ref_node)
        series_ref_count_local = sum(len(entry[3]) for entry in series_refs)
        return series_refs, plot_refs, series_ref_count_local, plot_level_dlbls

    series_label_refs, plot_only_refs, series_ref_count, plot_level_dlbls = _collect_label_refs()
    logger.info(
        "Waterfall chart cache update: data label strRef nodes found (series=%s, plot-level=%s)",
        series_ref_count,
        len(plot_only_refs),
    )
    if series_ref_count == 0 and (plot_only_refs or plot_level_dlbls):
        plot_dlbls = None
        if plot_level_dlbls:
            plot_dlbls = plot_level_dlbls[0]
        for ref_node in plot_only_refs:
            current = ref_node
            while current is not None and not current.tag.endswith("dLbls"):
                current = current.getparent()
            if current is not None:
                plot_dlbls = current
                break
        if plot_dlbls is None:
            plot_dlbls = root.find(".//c:plotArea//c:dLbls", namespaces=nsmap)
        if plot_dlbls is None:
            logger.info(
                "Waterfall chart cache update: plot-level data labels not found for promotion",
            )
        else:
            promoted = 0
            for ser in root.findall(".//c:plotArea//c:ser", namespaces=nsmap):
                if ser.find("c:dLbls", namespaces=nsmap) is None:
                    ser.append(copy.deepcopy(plot_dlbls))
                    promoted += 1
            parent = plot_dlbls.getparent()
            if parent is not None:
                parent.remove(plot_dlbls)
            logger.info(
                "Waterfall chart cache update: promoted plot-level data labels into %s series and removed plot-level node",
                promoted,
            )
            series_label_refs, plot_only_refs, series_ref_count, plot_level_dlbls = _collect_label_refs()
            logger.info(
                "Waterfall chart cache update: data label strRef nodes found after promotion (series=%s, plot-level=%s)",
                series_ref_count,
                len(plot_only_refs),
            )
    for idx, series_label, chart_series, label_refs, ser in series_label_refs:
        seen_refs = set()
        deduped_refs = []
        for ref in label_refs:
            if id(ref) in seen_refs:
                continue
            seen_refs.add(id(ref))
            deduped_refs.append(ref)
        mapped_header, mapped_labs_column = _resolve_waterfall_labs_column(ws, series_label)
        if mapped_labs_column is None:
            value_formula_node = ser.find("c:val/c:numRef/c:f", namespaces=nsmap)
            if value_formula_node is not None and value_formula_node.text:
                value_bounds = _range_boundaries_from_formula(value_formula_node.text)
                if value_bounds:
                    value_col = value_bounds[0]
                    base_header = ws.cell(row=1, column=value_col).value
                    if base_header:
                        derived_header = f"labs-{str(base_header).strip()}"
                        derived_col = _find_header_column(ws, [derived_header])
                        if derived_col is not None:
                            resolved_header = ws.cell(row=1, column=derived_col).value
                            mapped_header = derived_header
                            mapped_labs_column = derived_col
                            logger.info(
                                "Waterfall chart cache update: series %s fallback labs header resolved %s -> %s (%s)",
                                series_label,
                                derived_header,
                                resolved_header,
                                get_column_letter(derived_col),
                            )
        if mapped_header and mapped_labs_column:
            logger.info(
                "Waterfall chart cache update: series %s mapped to %s (%s)",
                series_label,
                mapped_header,
                get_column_letter(mapped_labs_column),
            )
        else:
            logger.info(
                "Waterfall chart cache update: series %s labs mapping missing",
                series_label,
            )
        bounds = series_category_bounds.get(idx) or series_value_bounds.get(idx)
        if bounds:
            min_row, max_row, sheet_name = bounds
        else:
            value_formula_node = ser.find("c:val/c:numRef/c:f", namespaces=nsmap)
            min_row = max_row = None
            sheet_name = None
            if value_formula_node is not None and value_formula_node.text:
                value_bounds = _range_boundaries_from_formula(value_formula_node.text)
                if value_bounds:
                    _, min_row, _, max_row = value_bounds
                    _, _, sheet_name = _worksheet_and_range_from_formula(
                        workbook, value_formula_node.text
                    )
        c15_updated = False
        if deduped_refs:
            for ref_node in deduped_refs:
                f_node = ref_node.find("c:f", namespaces=nsmap)
                if f_node is None or not f_node.text:
                    logger.info(
                        "Waterfall chart cache update: series %s data label ref formula missing",
                        series_label,
                    )
                    continue
                original_formula = f_node.text
                logger.info(
                    "Waterfall chart cache update: series %s data label ref formula %s",
                    series_label,
                    f_node.text,
                )
                expected_formula = None
                formula_bounds = _range_boundaries_from_formula(f_node.text)
                formula_col = formula_bounds[0] if formula_bounds else None
                effective_label_col = mapped_labs_column or formula_col
                if mapped_labs_column and min_row is not None and max_row is not None:
                    expected_formula = _build_cell_range_formula(
                        sheet_name,
                        effective_label_col,
                        min_row,
                        max_row,
                    )
                    if f_node.text != expected_formula:
                        logger.info(
                            "Waterfall chart cache update: series %s label formula %s -> %s",
                            series_label,
                            f_node.text,
                            expected_formula,
                        )
                        f_node.text = expected_formula
                elif effective_label_col and min_row is not None and max_row is not None:
                    expected_formula = _build_cell_range_formula(
                        sheet_name,
                        effective_label_col,
                        min_row,
                        max_row,
                    )
                elif effective_label_col:
                    series_points = series_point_counts.get(idx)
                    if series_points:
                        expected_formula = _build_cell_range_formula(
                            sheet_name or ws.title,
                            effective_label_col,
                            2,
                            1 + series_points,
                        )
                        if f_node.text != expected_formula:
                            logger.info(
                                "Waterfall chart cache update: series %s label formula %s -> %s",
                                series_label,
                                f_node.text,
                                expected_formula,
                            )
                            f_node.text = expected_formula
                logger.info(
                    "Waterfall chart cache update: series %s label formula resolved %s -> %s",
                    series_label,
                    original_formula,
                    f_node.text,
                )
                label_ws, label_ref_range, _ = _worksheet_and_range_from_formula(
                    workbook, f_node.text
                )
                label_rows = _range_values_from_worksheet(label_ws, label_ref_range)
                label_values = _flatten_cell_values(label_rows)
                if _all_blank(label_values):
                    logger.info(
                        "Waterfall chart cache update: series %s data label range '%s' is blank; skipping cache update",
                        series_label,
                        f_node.text,
                    )
                    continue
                series_points = series_point_counts.get(idx, len(label_values))
                if len(label_values) < series_points:
                    label_values += ["" for _ in range(series_points - len(label_values))]
                elif len(label_values) > series_points:
                    label_values = label_values[:series_points]
                str_cache, created = _ensure_str_cache(ref_node)
                if created:
                    label_cache_missing += 1
                _update_str_cache(
                    str_cache,
                    ["" if value is None else str(value) for value in label_values],
                )
                label_cache_updates += 1
                if not c15_updated:
                    d_lbls_node = ref_node
                    while d_lbls_node is not None and not d_lbls_node.tag.endswith("dLbls"):
                        d_lbls_node = d_lbls_node.getparent()
                    c15_container = d_lbls_node if d_lbls_node is not None else ser
                    c15_label_updates += _update_c15_label_range_cache(
                        c15_container,
                        expected_formula or f_node.text,
                        ["" if value is None else str(value) for value in label_values],
                        nsmap,
                        f"series {series_label}",
                    )
                    c15_updated = True
                logger.info(
                    "Waterfall chart cache update: series %s cached %s data label points",
                    series_label,
                    len(label_values),
                )
                if expected_formula:
                    logger.info(
                        "Waterfall chart cache update: series %s data label ref updated to %s",
                        series_label,
                        expected_formula,
                    )
                logger.info(
                    "Waterfall chart cache update: series %s data label formula now %s",
                    series_label,
                    f_node.text,
                )
                logger.info(
                    "Waterfall chart cache update: series %s data label formula %s cached %s points",
                    series_label,
                    f_node.text,
                    len(label_values),
                )
        if not deduped_refs:
            c15_ranges = ser.findall(".//c15:datalabelsRange", namespaces=nsmap)
            if not c15_ranges:
                try:
                    c15_ranges = ser.xpath(".//*[local-name()='datalabelsRange']")
                except Exception:
                    c15_ranges = []
            if c15_ranges:
                logger.info(
                    "Waterfall chart cache update: series %s c15 label ranges found without c:strRef",
                    series_label,
                )
            for c15_range in c15_ranges:
                c15_formula_node = c15_range.find("c15:f", namespaces=nsmap)
                if c15_formula_node is None:
                    try:
                        c15_formula_node = c15_range.xpath("./*[local-name()='f']")[0]
                    except Exception:
                        c15_formula_node = None
                if c15_formula_node is None or not c15_formula_node.text:
                    logger.info(
                        "Waterfall chart cache update: series %s c15 label formula missing",
                        series_label,
                    )
                    continue
                original_c15_formula = c15_formula_node.text
                logger.info(
                    "Waterfall chart cache update: series %s c15 label formula %s",
                    series_label,
                    c15_formula_node.text,
                )
                expected_formula = None
                formula_bounds = _range_boundaries_from_formula(c15_formula_node.text)
                formula_col = formula_bounds[0] if formula_bounds else None
                effective_label_col = mapped_labs_column or formula_col
                if effective_label_col and min_row is not None and max_row is not None:
                    expected_formula = _build_cell_range_formula(
                        sheet_name,
                        effective_label_col,
                        min_row,
                        max_row,
                    )
                elif effective_label_col:
                    series_points = series_point_counts.get(idx)
                    if series_points:
                        expected_formula = _build_cell_range_formula(
                            sheet_name or ws.title,
                            effective_label_col,
                            2,
                            1 + series_points,
                        )
                if expected_formula and c15_formula_node.text != expected_formula:
                    logger.info(
                        "Waterfall chart cache update: series %s c15 label formula %s -> %s",
                        series_label,
                        c15_formula_node.text,
                        expected_formula,
                    )
                    c15_formula_node.text = expected_formula
                logger.info(
                    "Waterfall chart cache update: series %s c15 label formula resolved %s -> %s",
                    series_label,
                    original_c15_formula,
                    c15_formula_node.text,
                )
                label_ws, label_ref_range, _ = _worksheet_and_range_from_formula(
                    workbook, c15_formula_node.text
                )
                label_rows = _range_values_from_worksheet(label_ws, label_ref_range)
                label_values = _flatten_cell_values(label_rows)
                if _all_blank(label_values):
                    logger.info(
                        "Waterfall chart cache update: series %s c15 label range '%s' is blank; skipping cache update",
                        series_label,
                        c15_formula_node.text,
                    )
                    continue
                series_points = series_point_counts.get(idx, len(label_values))
                if len(label_values) < series_points:
                    label_values += ["" for _ in range(series_points - len(label_values))]
                elif len(label_values) > series_points:
                    label_values = label_values[:series_points]
                c15_label_updates += _update_c15_label_range_cache(
                    c15_range,
                    expected_formula or c15_formula_node.text,
                    ["" if value is None else str(value) for value in label_values],
                    nsmap,
                    f"series {series_label}",
                )
                logger.info(
                    "Waterfall chart cache update: series %s c15 label formula now %s",
                    series_label,
                    c15_formula_node.text,
                )
    for ref_node in plot_only_refs:
        f_node = ref_node.find("c:f", namespaces=nsmap)
        if f_node is None or not f_node.text:
            logger.info(
                "Waterfall chart cache update: plot-level data label ref formula missing",
            )
            continue
        logger.info(
            "Waterfall chart cache update: plot-level data label ref formula %s",
            f_node.text,
        )
        label_ws, label_ref_range, _ = _worksheet_and_range_from_formula(
            workbook, f_node.text
        )
        label_rows = _range_values_from_worksheet(label_ws, label_ref_range)
        label_values = _flatten_cell_values(label_rows)
        if _all_blank(label_values):
            raise ValueError(
                f"Chart cache: plot-level data label range '{f_node.text}' is blank."
            )
        series_points = categories_count or len(label_values)
        if len(label_values) < series_points:
            label_values += ["" for _ in range(series_points - len(label_values))]
        elif len(label_values) > series_points:
            label_values = label_values[:series_points]
        str_cache, created = _ensure_str_cache(ref_node)
        if created:
            label_cache_missing += 1
        _update_str_cache(
            str_cache,
            ["" if value is None else str(value) for value in label_values],
        )
        label_cache_updates += 1
        d_lbls_node = ref_node
        while d_lbls_node is not None and not d_lbls_node.tag.endswith("dLbls"):
            d_lbls_node = d_lbls_node.getparent()
        if d_lbls_node is not None:
            c15_label_updates += _update_c15_label_range_cache(
                d_lbls_node,
                f_node.text,
                ["" if value is None else str(value) for value in label_values],
                nsmap,
                "plot-level",
            )
        logger.info(
            "Waterfall chart cache update: plot-level cached %s data label points",
            len(label_values),
        )
        logger.info(
            "Waterfall chart cache update: plot-level data label formula %s cached %s points",
            f_node.text,
            len(label_values),
        )
    if label_cache_updates:
        logger.info(
            "Waterfall chart cache update: %s data label caches updated",
            label_cache_updates,
        )
    if c15_label_updates:
        logger.info(
            "Waterfall chart cache update: %s c15 label-range caches updated",
            c15_label_updates,
        )
    elif label_cache_missing:
        logger.info(
            "Waterfall chart cache update: chart is not using value-from-cells labels",
        )
    else:
        logger.info(
            "Waterfall chart cache update: chart is not using value-from-cells labels",
        )


def _update_chart_label_caches(chart, workbook) -> None:
    from ..waterfall.inject import _find_header_column, _resolve_waterfall_labs_column
    root = chart.part._element
    nsmap = _chart_namespace_map(root)
    label_refs = root.findall(".//c:dLbls//c:tx//c:strRef", namespaces=nsmap)
    if not label_refs:
        return
    label_cache_updates = 0
    c15_label_updates = 0
    for ref_node in label_refs:
        f_node = ref_node.find("c:f", namespaces=nsmap)
        if f_node is None or not f_node.text:
            logger.info("Chart label cache update: data label ref formula missing")
            continue
        label_ws, label_ref_range, _ = _worksheet_and_range_from_formula(
            workbook, f_node.text
        )
        label_rows = _range_values_from_worksheet(label_ws, label_ref_range)
        label_values = _flatten_cell_values(label_rows)
        if _all_blank(label_values):
            logger.info(
                "Chart label cache update: data label range '%s' is blank",
                f_node.text,
            )
            continue
        str_cache, _ = _ensure_str_cache(ref_node)
        _update_str_cache(
            str_cache,
            ["" if value is None else str(value) for value in label_values],
        )
        label_cache_updates += 1
        d_lbls_node = ref_node
        while d_lbls_node is not None and not d_lbls_node.tag.endswith("dLbls"):
            d_lbls_node = d_lbls_node.getparent()
        if d_lbls_node is not None:
            c15_label_updates += _update_c15_label_range_cache(
                d_lbls_node,
                f_node.text,
                ["" if value is None else str(value) for value in label_values],
                nsmap,
                "chart",
            )
    if label_cache_updates:
        logger.info(
            "Chart label cache update: %s data label caches updated",
            label_cache_updates,
        )
    if c15_label_updates:
        logger.info(
            "Chart label cache update: %s c15 label-range caches updated",
            c15_label_updates,
        )
