#WATERFALL ORCHESTRATION/UPLOAD FLOW

def _waterfall_template_marker(index: int) -> str:
    if index < 0:
        raise ValueError("Waterfall template index must be non-negative.")
    if index == 0:
        return "<Waterfall Template>"
    return f"<Waterfall Template{index + 1}>"


def _available_waterfall_template_slides(prs) -> list[tuple[str, object]]:
    slides = []
    idx = 0
    while True:
        marker = _waterfall_template_marker(idx)
        slide = _find_slide_by_marker(prs, marker)
        if slide is None:
            break
        slides.append((marker, slide))
        idx += 1
    return slides

def _update_waterfall_chart(
    slide,
    payload: WaterfallPayload,
) -> None:
    chart_shapes = [shape for shape in slide.shapes if shape.has_chart]
    if not chart_shapes:
        raise ValueError("Could not find the waterfall chart on the <Waterfall Template> slide.")
    for chart_shape in chart_shapes:
        chart = chart_shape.chart
        series_names = [series.name for series in chart.series]
        label_columns = _capture_label_columns(_load_chart_workbook(chart).active, series_names)
        cd = _chart_data_from_payload(payload)
        chart.replace_data(cd)
        updated_wb = _load_chart_workbook(chart)
        total_rows = len(payload.categories)
        _update_lab_base_label(
            label_columns,
            payload.base_indices,
            payload.base_values,
            total_rows,
        )
        _apply_label_columns(updated_wb.active, label_columns, total_rows)
        _ensure_negatives_column_positive(updated_wb.active)
        applied_headers = _apply_gathered_waterfall_labels(
            updated_wb.active,
            payload.gathered_label_values,
            total_rows,
        )
        _update_all_waterfall_labs(
            updated_wb.active,
            payload.base_indices,
            payload.base_values,
            skip_headers=applied_headers,
        )
        _save_chart_workbook(chart, updated_wb)
        chart_type = getattr(
            chart,
            "chart_type",
            getattr(XL_CHART_TYPE, "WATERFALL", XL_CHART_TYPE.COLUMN_STACKED),
        )
        if chart_type == getattr(XL_CHART_TYPE, "WATERFALL", XL_CHART_TYPE.COLUMN_STACKED):
            _update_waterfall_chart_caches(chart, updated_wb, payload.categories)
        else:
            _update_chart_label_caches(chart, updated_wb)
    _update_waterfall_yoy_arrows(slide, payload.base_values)

def populate_category_waterfall(
    prs,
    gathered_df: pd.DataFrame,
    scope_df: pd.DataFrame | None = None,
    target_labels: list[str] | None = None,
    bucket_data: dict | None = None,
    modelled_in_value: str | None = None,
    metric_value: str | None = None,
):
    labels = _normalize_target_level_labels(target_labels)
    if not labels:
        labels = _target_level_labels_from_gathered_df_with_filters(
            gathered_df,
            year1=bucket_data.get("year1") if bucket_data else None,
            year2=bucket_data.get("year2") if bucket_data else None,
            target_labels=bucket_data.get("target_labels") if bucket_data else None,
        )
    if not labels:
        return

    available_slides = _available_waterfall_template_slides(prs)
    available_count = len(available_slides)
    if available_count == 0:
        raise ValueError("Could not find the <Waterfall Template> slide in the template.")
    if len(labels) > available_count:
        raise ValueError(
            "Need {needed} waterfall slides but only found {available} "
            "(<Waterfall Template>...<Waterfall Template{available}>) in template. "
            "Please add more duplicated template slides or use a larger template deck.".format(
                needed=len(labels),
                available=available_count,
            )
        )

    template_chart = _waterfall_chart_from_slide(available_slides[0][1], "Waterfall Template")
    if template_chart is None:
        raise ValueError("Could not find the waterfall chart on the <Waterfall Template> slide.")
    payloads_by_label = compute_waterfall_payloads_for_all_labels(
        gathered_df,
        scope_df,
        bucket_data,
        template_chart,
        target_labels=labels,
    )

    seen_partnames: set[str] = set()
    remaining_labels = labels.copy()
    for idx in range(len(labels)):
        marker_text, slide = available_slides[idx]
        resolved_label = _resolve_target_level_label_for_slide(slide, remaining_labels)
        if resolved_label is None:
            if not remaining_labels:
                raise ValueError("No remaining Target Level Label values to assign.")
            resolved_label = remaining_labels[0]
            logger.info(
                "No slide title/name found for %s; using ordered Target Level Label %r.",
                marker_text,
                resolved_label,
            )
        if resolved_label in remaining_labels:
            remaining_labels.remove(resolved_label)
        if resolved_label not in payloads_by_label:
            raise ValueError(
                f"Missing precomputed waterfall payload for Target Level Label {resolved_label!r}."
            )
        _ensure_unique_chart_parts_on_slide(slide, seen_partnames)
        _update_waterfall_axis_placeholders(
            prs,
            slide,
            target_level_label_value=resolved_label,
            modelled_in_value=modelled_in_value,
            metric_value=metric_value,
        )
        _update_waterfall_chart(
            slide,
            payloads_by_label[resolved_label],
        )
        _set_waterfall_slide_header(slide, resolved_label, marker_text=marker_text)

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

#PAYLOAD CONSTRUCTION

def _waterfall_series_from_gathered_df(
    gathered_df: pd.DataFrame,
    scope_df: pd.DataFrame | None,
    target_level_label: str,
) -> tuple[list[str], dict[str, list[float]], dict[str, list]] | None:
    waterfall_df = _build_category_waterfall_df(
        gathered_df,
        target_level_label=target_level_label,
    )
    if waterfall_df.empty:
        return None
    categories = (
        waterfall_df["Vars"]
        .fillna("")
        .astype(str)
        .tolist()
    )
    categories = _replace_modelling_period_placeholders_in_categories(categories, scope_df)
    series_values = {}
    for key in ["Base", "Promo", "Media", "Blanks", "Positives", "Negatives"]:
        if key in waterfall_df.columns:
            series_values[key] = (
                pd.to_numeric(waterfall_df[key], errors="coerce")
                .fillna(0)
                .astype(float)
                .tolist()
            )
    label_values: dict[str, list] = {}
    for key in [
        "labs-Base",
        "labs-Promo",
        "labs-Media",
        "labs-Blanks",
        "labs-Positives",
        "labs-Negatives",
    ]:
        if key in waterfall_df.columns:
            values = []
            for value in waterfall_df[key].tolist():
                if pd.isna(value):
                    values.append(None)
                else:
                    values.append(value)
            label_values[key] = values
    if not series_values and not label_values:
        return None
    return categories, series_values, label_values

def _build_waterfall_chart_data(
    chart,
    scope_df: pd.DataFrame | None,
    gathered_df: pd.DataFrame | None = None,
    target_level_label: str | None = None,
    bucket_labels: list[str] | None = None,
    bucket_values: list[float] | None = None,
    year1: str | None = None,
    year2: str | None = None,
) -> tuple[
    ChartData,
    list[str],
    tuple[int, int] | None,
    tuple[float, float] | None,
    list[tuple[str, list[float]]],
    dict[str, list],
]:
    gathered_override = None
    gathered_label_values: dict[str, list] = {}
    if gathered_df is not None and target_level_label:
        try:
            gathered_override = _waterfall_series_from_gathered_df(
                gathered_df,
                scope_df,
                target_level_label,
            )
        except Exception as exc:
            logger.info(
                "Skipping gatheredCN10 waterfall data for %r: %s",
                target_level_label,
                exc,
            )
            gathered_override = None
    if gathered_override is not None:
        categories, gathered_series, gathered_label_values = gathered_override
    else:
        categories = _categories_from_chart(chart)
        gathered_series = {}
        categories = _replace_modelling_period_placeholders_in_categories(categories, scope_df)
        gathered_label_values = {}
    base_indices = _waterfall_base_indices(categories)
    original_base_indices = base_indices
    bucket_labels = list(bucket_labels or [])
    bucket_values = _sanitize_numeric_list(
        list(bucket_values or []),
        label=target_level_label,
        field_prefix="bucket_values",
        bucket_labels=bucket_labels,
        year1=year1,
        year2=year2,
    )
    if bucket_labels and bucket_values:
        bucket_len = min(len(bucket_labels), len(bucket_values))
        bucket_labels = bucket_labels[:bucket_len]
        bucket_values = bucket_values[:bucket_len]
    if bucket_labels and base_indices:
        categories, base_indices = _apply_bucket_categories(
            categories,
            bucket_labels,
            base_indices,
        )
    bucket_count = len(bucket_labels)
    base_values = None
    if (
        gathered_df is not None
        and target_level_label
        and base_indices is not None
    ):
        base_values = _waterfall_base_values(
            gathered_df,
            target_level_label,
            year1=year1,
            year2=year2,
        )
        base_values = (
            _sanitize_numeric_value(
                base_values[0],
                label=target_level_label,
                field="base_values[0]",
                year1=year1,
                year2=year2,
            ),
            _sanitize_numeric_value(
                base_values[1],
                label=target_level_label,
                field="base_values[1]",
                year1=year1,
                year2=year2,
            ),
        )
    cd = ChartData()
    cd.categories = categories
    base_start_value = None
    if base_values and base_values[0] is not None:
        base_start_value = float(base_values[0])
    elif base_indices is not None:
        for series in chart.series:
            if _should_update_base_series(series):
                series_values = list(series.values)
                if base_indices[0] < len(series_values):
                    base_start_value = float(series_values[base_indices[0]])
                break
    if base_start_value is None:
        base_start_value = 0.0

    positive_bucket_values = []
    negative_bucket_values = []
    blank_bucket_values = []
    if bucket_labels and bucket_values:
        positive_bucket_values, negative_bucket_values = _bucket_value_split(bucket_values)
        blank_bucket_values = _bucket_blank_values(bucket_values, base_start_value)

    series_candidates = list(gathered_series.keys())
    series_values: list[tuple[str, list[float]]] = []
    for series in chart.series:
        values = list(series.values)
        if gathered_series:
            resolved_series = None
            try:
                resolved_series = _resolve_label_from_text(
                    str(series.name),
                    series_candidates,
                )
            except Exception as exc:
                logger.info(
                    "No gatheredCN10 series match for %r: %s",
                    series.name,
                    exc,
                )
            if resolved_series:
                values = list(gathered_series.get(resolved_series, values))
                values = _align_series_values(values, len(categories))
        if original_base_indices and bucket_labels:
            if _is_positive_series(series):
                if positive_bucket_values:
                    values = _apply_bucket_values(
                        values,
                        original_base_indices,
                        positive_bucket_values,
                    )
                else:
                    values = _apply_bucket_placeholders(
                        values,
                        original_base_indices,
                        bucket_count,
                    )
            elif _is_negative_series(series):
                if negative_bucket_values:
                    values = _apply_bucket_values(
                        values,
                        original_base_indices,
                        negative_bucket_values,
                    )
                else:
                    values = _apply_bucket_placeholders(
                        values,
                        original_base_indices,
                        bucket_count,
                    )
            elif _is_blanks_series(series):
                if blank_bucket_values:
                    values = _apply_bucket_values(
                        values,
                        original_base_indices,
                        blank_bucket_values,
                    )
                else:
                    values = _apply_bucket_placeholders(
                        values,
                        original_base_indices,
                        bucket_count,
                    )
            else:
                values = _apply_bucket_placeholders(
                    values,
                    original_base_indices,
                    bucket_count,
                )
        if base_values and base_indices:
            should_update = _should_update_base_series(series)
            if not should_update and len(chart.series) == 1:
                should_update = True
            if should_update:
                if base_indices[0] < len(values):
                    values[base_indices[0]] = base_values[0]
                if base_indices[1] < len(values):
                    values[base_indices[1]] = base_values[1]
        values = _sanitize_numeric_list(
            values,
            label=target_level_label,
            field_prefix=f"series_values[{len(series_values)}]",
            categories=categories,
            year1=year1,
            year2=year2,
        )
        cd.add_series(series.name, values)
        series_values.append((series.name, values))
    return cd, categories, base_indices, base_values, series_values, gathered_label_values

@dataclass(frozen=True)
class WaterfallPayload:
    categories: list[str]
    series_values: list[tuple[str, list[float]]]
    base_indices: tuple[int, int] | None
    base_values: tuple[float, float] | None
    gathered_label_values: dict[str, list]


def _payload_checksum(series_values: list[tuple[str, list[float]]]) -> float:
    if not series_values:
        return 0.0
    checksum = 0.0
    if isinstance(series_values[0], tuple):
        for series_idx, (_, values) in enumerate(series_values):
            for value_idx, value in enumerate(values):
                if value is None or pd.isna(value):
                    logger.warning(
                        '[waterfall][checksum] field="series_values[%d][%d]" was=%r -> 0.0',
                        series_idx,
                        value_idx,
                        value,
                    )
                    continue
                checksum += abs(float(value))
        return checksum
    for value_idx, value in enumerate(series_values):
        if value is None or pd.isna(value):
            logger.warning(
                '[waterfall][checksum] field="values[%d]" was=%r -> 0.0',
                value_idx,
                value,
            )
            continue
        checksum += abs(float(value))
    return checksum


def _chart_data_from_payload(payload: WaterfallPayload) -> ChartData:
    cd = ChartData()
    cd.categories = payload.categories
    for name, values in payload.series_values:
        cd.add_series(name, values)
    return cd


def compute_payload_for_label(
    gathered_df: pd.DataFrame,
    scope_df: pd.DataFrame | None,
    target_level_label: str,
    bucket_data: dict | None,
    template_chart,
) -> WaterfallPayload:
    if template_chart is None:
        raise ValueError("Template chart is required to compute waterfall payloads.")
    (
        _,
        categories,
        base_indices,
        base_values,
        series_values,
        gathered_label_values,
    ) = _build_waterfall_chart_data(
        template_chart,
        scope_df,
        gathered_df,
        target_level_label,
        bucket_data.get("labels") if bucket_data else None,
        bucket_data.get("values") if bucket_data else None,
        year1=bucket_data.get("year1") if bucket_data else None,
        year2=bucket_data.get("year2") if bucket_data else None,
    )
    return WaterfallPayload(
        categories=list(categories),
        series_values=[(name, list(values)) for name, values in series_values],
        base_indices=base_indices,
        base_values=base_values,
        gathered_label_values={
            key: list(values) for key, values in gathered_label_values.items()
        },
    )


def compute_waterfall_payloads_for_all_labels(
    gathered_df: pd.DataFrame,
    scope_df: pd.DataFrame | None,
    bucket_data: dict | None,
    template_chart,
    target_labels: list[str] | None = None,
) -> dict[str, WaterfallPayload]:
    labels = _normalize_target_level_labels(target_labels)
    if not labels:
        labels = _target_level_labels_from_gathered_df_with_filters(
            gathered_df,
            year1=bucket_data.get("year1") if bucket_data else None,
            year2=bucket_data.get("year2") if bucket_data else None,
            target_labels=bucket_data.get("target_labels") if bucket_data else None,
        )

#FUZZY MATCHING

def _resolve_column_from_candidates(
    df: pd.DataFrame,
    header_row: pd.Series | None,
    candidates: list[str],
    *,
    context: str,
    threshold: float = 85.0,
) -> tuple[str | None, int, float]:
    column_options: list[tuple[str, object, int]] = []
    for column in df.columns:
        column_options.append((str(column), column, 0))
    if header_row is not None:
        for column, value in header_row.items():
            if pd.isna(value):
                continue
            column_options.append((str(value), column, 1))

    normalized_candidates = [
        (candidate, _normalize_lookup_value(candidate)) for candidate in candidates
    ]
    normalized_options = [
        (label, _normalize_lookup_value(label), column, data_start_idx)
        for label, column, data_start_idx in column_options
        if _normalize_lookup_value(label)
    ]

    exact_matches = []
    for candidate, normalized_candidate in normalized_candidates:
        for label, normalized_label, column, data_start_idx in normalized_options:
            if normalized_candidate == normalized_label:
                exact_matches.append((candidate, label, column, data_start_idx, 100.0))
    if exact_matches:
        unique_columns = {match[2] for match in exact_matches}
        if len(unique_columns) > 1:
            top_candidates = [
                (match[4], match[1]) for match in exact_matches[:5]
            ]
            raise ValueError(
                f"Ambiguous {context} match. Top candidates: "
                f"{_format_fuzzy_candidates(top_candidates)}"
            )
        candidate, label, column, data_start_idx, score = exact_matches[0]
        logger.info(
            'Resolved header "%s" -> "%s" (score %.1f)',
            candidate,
            label,
            score,
        )
        return column, data_start_idx, score

    from difflib import SequenceMatcher

    scored: list[tuple[float, str, object, int, str]] = []
    for candidate, normalized_candidate in normalized_candidates:
        for label, normalized_label, column, data_start_idx in normalized_options:
            score = SequenceMatcher(None, normalized_candidate, normalized_label).ratio() * 100
            scored.append((score, label, column, data_start_idx, candidate))
    if not scored:
        return None, 0, 0.0
    scored.sort(key=lambda item: item[0], reverse=True)
    top_score = scored[0][0]
    if top_score < threshold:
        return None, 0, top_score
    close_matches = [
        (score, label, column)
        for score, label, column, _, _ in scored
        if score >= top_score - 1.0
    ]
    unique_columns = {column for _, _, column in close_matches}
    if len(unique_columns) > 1:
        top_candidates = [(score, label) for score, label, _ in close_matches[:5]]
        raise ValueError(
            f"Ambiguous {context} match. Top candidates: "
            f"{_format_fuzzy_candidates(top_candidates)}"
        )
    score, label, column, data_start_idx, candidate = scored[0]
    logger.info(
        'Resolved header "%s" -> "%s" (score %.1f)',
        candidate,
        label,
        score,
    )
    return column, data_start_idx, score


def _resolve_target_level_label_for_slide(slide, labels: list[str]) -> str | None:
    candidates = []
    slide_title = _slide_title(slide)
    if slide_title:
        candidates.append(("title", slide_title))
    slide_name = getattr(slide, "name", None) or ""
    if slide_name:
        candidates.append(("name", slide_name))
    if not candidates:
        return None
    errors = []
    for source, text in candidates:
        try:
            return _resolve_label_from_text(text, labels)
        except ValueError as exc:
            message = str(exc)
            error_message = (
                f"Could not resolve Target Level Label from slide {source} {text!r}: {exc}"
            )
            if message.startswith("No slide text match found") or message.startswith(
                "Slide text is empty after normalization"
            ):
                logger.info("%s", error_message)
                continue
            errors.append(error_message)
    if errors:
        raise ValueError(" | ".join(errors))
    return None

#EMBEDDED WORKBOOK CACHES

def _load_chart_workbook(chart):
    xlsx_blob = chart.part.chart_workbook.xlsx_part.blob
    return load_workbook(io.BytesIO(xlsx_blob))


def _save_chart_workbook(chart, workbook) -> None:
    stream = io.BytesIO()
    workbook.save(stream)
    chart.part.chart_workbook.xlsx_part.blob = stream.getvalue()

def _update_waterfall_chart_caches(chart, workbook, categories: list[str]) -> None:
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

def _capture_label_columns(ws, series_names: list[str]) -> dict[int, dict[str, list]]:
    label_columns: dict[int, dict[str, list]] = {}
    series_lookup = {str(name).strip().lower() for name in series_names if name}
    for col_idx in range(2, ws.max_column + 1):
        header = ws.cell(row=1, column=col_idx).value
        if not header:
            continue
        header_text = str(header).strip().lower()
        if header_text in series_lookup:
            continue
        values = [
            ws.cell(row=row_idx, column=col_idx).value
            for row_idx in range(2, ws.max_row + 1)
        ]
        label_columns[col_idx] = {"header": header, "values": values}
    return label_columns


def _apply_label_columns(ws, label_columns: dict[int, dict[str, list]], total_rows: int) -> None:
    for col_idx, column in label_columns.items():
        ws.cell(row=1, column=col_idx, value=column["header"])
        values = column["values"]
        if len(values) < total_rows:
            values = values + [None] * (total_rows - len(values))
        for row_offset in range(total_rows):
            ws.cell(row=row_offset + 2, column=col_idx, value=values[row_offset])


def _apply_gathered_waterfall_labels(
    ws,
    label_values: dict[str, list],
    total_rows: int,
) -> set[str]:
    applied_headers: set[str] = set()
    if not label_values:
        return applied_headers
    for header, values in label_values.items():
        col_idx = _find_header_column(ws, [header])
        if col_idx is None:
            continue
        applied_headers.add(_normalize_column_name(header))
        aligned_values = _align_label_values(values, total_rows)
        for row_offset in range(total_rows):
            ws.cell(row=row_offset + 2, column=col_idx, value=aligned_values[row_offset])
    return applied_headers

def _update_all_waterfall_labs(
    ws,
    base_indices: tuple[int, int] | None,
    base_values: tuple[float, float] | None,
    skip_headers: set[str] | None = None,
) -> None:
    skip_headers = {value for value in (skip_headers or set()) if value}
    labs_base_col = _find_header_column(ws, ["labs-Base"])
    labs_promo_col = _find_header_column(ws, ["labs-Promo"])
    labs_media_col = _find_header_column(ws, ["labs-Media"])
    labs_blanks_col = _find_header_column(ws, ["labs-Blanks"])
    labs_positives_col = _find_header_column(ws, ["labs-Positives"])
    labs_negatives_col = _find_header_column(ws, ["labs-Negatives"])

    promo_col = _find_header_column(ws, ["Promo"])
    media_col = _find_header_column(ws, ["Media"])
    positives_col = _find_header_column(ws, ["Positives"])
    negatives_col = _find_header_column(ws, ["Negatives"])

    total_rows = ws.max_row
    if labs_base_col and _normalize_column_name("labs-Base") not in skip_headers:
        for row_idx in range(2, total_rows + 1):
            ws.cell(row=row_idx, column=labs_base_col).value = None
        if base_indices and base_values:
            formatted = [_format_lab_base_value(value) for value in base_values]
            for idx, base_row in enumerate(base_indices):
                if base_row is None or base_row < 0:
                    continue
                row_idx = base_row + 2
                if row_idx <= total_rows:
                    ws.cell(row=row_idx, column=labs_base_col).value = formatted[idx]

#Regression tests to mirror in Django port:

def test_waterfall_labels_render_without_edit_data(tmp_path) -> None:
    template_path = tmp_path / "template.pptx"
    build_test_template(template_path, waterfall_slide_count=1)

    df = build_sample_dataframe(["Alpha"], include_brand=False)
    excel_path = tmp_path / "input.xlsx"
    write_excel(df, excel_path)
    df_from_excel = pd.read_excel(excel_path)
    pptx_bytes = build_deck_bytes(template_path, df_from_excel, waterfall_targets=["Alpha"])

    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as zf:
        chart_files = [
            name
            for name in zf.namelist()
            if name.startswith("ppt/charts/chart") and name.endswith(".xml")
        ]
        assert chart_files
        chart_xml = zf.read(chart_files[0])

    root = ET.fromstring(chart_xml)
    ns = {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart"}
    d_lbls = root.findall(".//c:dLbls", ns)
    assert d_lbls

    label_flags = [
        lbl.find("c:showVal", ns) is not None or lbl.find("c:showCatName", ns) is not None
        for lbl in d_lbls
    ]
    assert any(label_flags)

    cache_points = _series_cache_points(root)
    assert cache_points
    for cat_pt_count, cat_pts, num_pt_count, num_pts in cache_points:
        assert cat_pt_count > 0
        assert cat_pts > 0
        assert num_pt_count > 0
        assert num_pts > 0

def test_waterfall_c15_labels_render_without_edit_data(tmp_path) -> None:
    template_path = tmp_path / "template.pptx"
    build_test_template(template_path, waterfall_slide_count=1)
    _insert_value_from_cells_labels(template_path)

    df = build_sample_dataframe(["Alpha"], include_brand=False)
    excel_path = tmp_path / "input.xlsx"
    write_excel(df, excel_path)
    df_from_excel = pd.read_excel(excel_path)
    pptx_bytes = build_deck_bytes(template_path, df_from_excel, waterfall_targets=["Alpha"])

    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as zf:
        chart_files = [
            name
            for name in zf.namelist()
            if name.startswith("ppt/charts/chart") and name.endswith(".xml")
        ]
        assert chart_files
        chart_xml = zf.read(chart_files[0])
        embedding_files = [
            name
            for name in zf.namelist()
            if name.startswith("ppt/embeddings/") and name.endswith(".xlsx")
        ]
        assert embedding_files
        workbook_bytes = zf.read(embedding_files[0])

    ns = {
        "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
        "c15": "http://schemas.microsoft.com/office/drawing/2012/chart",
    }
    root = ET.fromstring(chart_xml)
    series_node = root.findall(".//c:ser", ns)[1]
    c15_range = series_node.find(".//c15:datalabelsRange", ns)
    assert c15_range is not None
    c15_formula = c15_range.find("c15:f", ns)
    assert c15_formula is not None

    value_formula = series_node.find("c:val/c:numRef/c:f", ns)
    assert value_formula is not None and value_formula.text
    workbook = load_workbook(io.BytesIO(workbook_bytes))
    ws, _, sheet_name = _worksheet_and_range_from_formula(workbook, value_formula.text)
    bounds = _range_boundaries_from_formula(value_formula.text)
    assert bounds is not None
    _, min_row, _, max_row = bounds
    labs_col = _find_header_column(ws, ["labs-Positives"])
    assert labs_col is not None
    expected_formula = _build_cell_range_formula(sheet_name, labs_col, min_row, max_row)

    assert c15_formula.text == expected_formula

    label_values = [
        ws.cell(row=row_idx, column=labs_col).value
        for row_idx in range(min_row, max_row + 1)
    ]
    expected_labels = ["" if value is None else str(value) for value in label_values]
    cache = c15_range.find("c15:dlblRangeCache", ns)
    assert cache is not None
    pt_count = cache.find("c15:ptCount", ns)
    assert pt_count is not None
    assert int(pt_count.attrib.get("val", "0")) == len(expected_labels)
    points = cache.findall("c15:pt", ns)
    assert len(points) == len(expected_labels)
    values = [
        (pt.find("c15:v", ns).text or "") if pt.find("c15:v", ns) is not None else ""
        for pt in points
    ]
    assert values[0] == expected_labels[0]
    assert values[-1] == expected_labels[-1]

def test_multi_target_level_labels_map_to_templates(tmp_path) -> None:
    labels = ["Alpha", "Beta"]
    template_path = tmp_path / "template.pptx"
    build_test_template(template_path, waterfall_slide_count=len(labels))

    df = build_sample_dataframe(labels, include_brand=False)
    pptx_bytes = build_deck_bytes(template_path, df, waterfall_targets=labels)

    with zipfile.ZipFile(template_path) as template_zip:
        template_slide_files = [
            name
            for name in template_zip.namelist()
            if name.startswith("ppt/slides/slide") and name.endswith(".xml")
        ]

    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as zf:
        slide_files = [
            name
            for name in zf.namelist()
            if name.startswith("ppt/slides/slide") and name.endswith(".xml")
        ]
        assert slide_files
        slide_texts = {name: zf.read(name).decode("utf-8", errors="ignore") for name in slide_files}

    for idx, label in enumerate(labels, start=1):
        style_text = f"Template Style {idx}"
        matching_slides = [
            name
            for name, content in slide_texts.items()
            if label in content and style_text in content
        ]
        assert matching_slides

    for marker in ("<Waterfall Template>", "<Waterfall Template2>"):
        assert not any(marker in content for content in slide_texts.values())

    assert len(slide_texts) == len(template_slide_files)


def test_waterfall_charts_are_independent_after_update(tmp_path) -> None:
    labels = ["Alpha", "Beta"]
    template_path = tmp_path / "template.pptx"
    build_test_template(template_path, waterfall_slide_count=len(labels))

    shared_bytes = _force_shared_chart_part(
        template_path.read_bytes(),
        "<Waterfall Template>",
        "<Waterfall Template2>",
    )
    template_path.write_bytes(shared_bytes)

    df = build_sample_dataframe(labels, include_brand=False)
    pptx_bytes = build_deck_bytes(template_path, df, waterfall_targets=labels)

    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as zf:
        slide_primary = _slide_name_for_marker(zf, "Alpha")
        slide_secondary = _slide_name_for_marker(zf, "Beta")
        chart_primary, _ = _chart_part_target(zf, slide_primary)
        chart_secondary, _ = _chart_part_target(zf, slide_secondary)
        assert chart_primary != chart_secondary
        assert zf.read(chart_primary) != zf.read(chart_secondary)

def test_waterfall_base_values_use_own_target_label() -> None:
    df = pd.DataFrame(
        [
            {"Target Level Label": "Alpha", "Target Label": "Own", "Year": "Year1", "Actuals": 10},
            {"Target Level Label": "Alpha", "Target Label": "Own", "Year": "Year2", "Actuals": 20},
            {
                "Target Level Label": "Alpha",
                "Target Label": "Cross",
                "Year": "Year1",
                "Actuals": 100,
            },
            {
                "Target Level Label": "Alpha",
                "Target Label": "Cross",
                "Year": "Year2",
                "Actuals": 150,
            },
        ]
    )
    year1_total, year2_total = _waterfall_base_values(
        df,
        "Alpha",
        year1="Year1",
        year2="Year2",
    )
    assert year1_total == 10
    assert year2_total == 20