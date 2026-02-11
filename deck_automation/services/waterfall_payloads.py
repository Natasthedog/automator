from __future__ import annotations

import json
import logging
from dataclasses import asdict, dataclass, is_dataclass

import pandas as pd

logger = logging.getLogger(__name__)

REQUIRED_COLUMNS = {
    "Target Level Label",
    "Target Label",
    "Year",
    "Actuals",
    "Vars",
    "Base",
    "Promo",
    "Media",
    "Blanks",
    "Positives",
    "Negatives",
}
WATERFALL_SERIES_COLUMNS = ["Base", "Promo", "Media", "Blanks", "Positives", "Negatives"]


@dataclass
class WaterfallPayload:
    categories: list[str]
    series_values: list[tuple[str, list[float]]]
    base_indices: tuple[int, int] | None
    base_values: tuple[float, float] | None
    gathered_label_values: dict[str, list]


def _payload_checksum(series_values: list[tuple[str, list[float]]]) -> float:
    checksum = 0.0
    for _, values in series_values:
        for value in values:
            if value is None or pd.isna(value):
                continue
            checksum += abs(float(value))
    return checksum


def payload_checksum(payload: WaterfallPayload) -> float:
    return _payload_checksum(payload.series_values)


def _normalize_target_level_labels(labels: list[str] | None) -> list[str]:
    normalized: list[str] = []
    seen: set[str] = set()
    for label in labels or []:
        if label is None:
            continue
        value = str(label).strip()
        if not value or value in seen:
            continue
        seen.add(value)
        normalized.append(value)
    return normalized


def _to_float(value) -> float:
    if value is None or pd.isna(value):
        return 0.0
    return float(value)


def _require_columns(gathered_df: pd.DataFrame) -> None:
    missing = sorted(REQUIRED_COLUMNS.difference(set(gathered_df.columns)))
    if missing:
        raise ValueError(
            "Missing required gatheredCN10 column(s): " + ", ".join(missing)
        )


def _target_level_labels_from_gathered_df_with_filters(gathered_df: pd.DataFrame) -> list[str]:
    labels = []
    seen = set()
    for value in gathered_df["Target Level Label"].tolist():
        normalized = str(value).strip() if value is not None else ""
        if not normalized or normalized in seen:
            continue
        seen.add(normalized)
        labels.append(normalized)
    return labels


def _compute_payload_for_label(gathered_df: pd.DataFrame, label: str) -> WaterfallPayload:
    filtered = gathered_df[
        gathered_df["Target Level Label"].astype(str).str.strip() == str(label).strip()
    ].copy()
    if filtered.empty:
        raise ValueError(f"No gatheredCN10 data found for Target Level Label {label!r}.")

    grouped = (
        filtered.groupby("Year", sort=False)[["Actuals", *WATERFALL_SERIES_COLUMNS]]
        .sum(numeric_only=True)
        .reset_index()
    )
    categories = [str(value) for value in grouped["Year"].tolist()]
    if not categories:
        raise ValueError(f"No Year/category rows found for Target Level Label {label!r}.")

    series_values = [
        (series_name, [_to_float(value) for value in grouped[series_name].tolist()])
        for series_name in WATERFALL_SERIES_COLUMNS
    ]

    base_indices: tuple[int, int] | None = None
    if len(categories) >= 2:
        base_indices = (0, len(categories) - 1)

    actuals_values = [_to_float(value) for value in grouped["Actuals"].tolist()]
    base_values: tuple[float, float] | None = None
    if len(actuals_values) >= 2:
        base_values = (actuals_values[0], actuals_values[-1])

    gathered_label_values = {
        "Year": categories,
        "Actuals": actuals_values,
    }

    return WaterfallPayload(
        categories=categories,
        series_values=series_values,
        base_indices=base_indices,
        base_values=base_values,
        gathered_label_values=gathered_label_values,
    )


def compute_waterfall_payloads_for_all_labels(
    gathered_df: pd.DataFrame,
    scope_df: pd.DataFrame | None,
    bucket_data: dict | None,
    template_chart=None,
    target_labels: list[str] | None = None,
) -> dict[str, WaterfallPayload]:
    del scope_df, bucket_data, template_chart
    _require_columns(gathered_df)

    labels = _normalize_target_level_labels(target_labels)
    if not labels:
        labels = _target_level_labels_from_gathered_df_with_filters(gathered_df)

    payloads_by_label: dict[str, WaterfallPayload] = {}
    logger.info("Precomputing waterfall payloads for %d label(s).", len(labels))
    for label in labels:
        payload = _compute_payload_for_label(gathered_df, label)
        payloads_by_label[label] = payload
        logger.info(
            "Computed waterfall payload for %r: %d categories, checksum %.2f",
            label,
            len(payload.categories),
            payload_checksum(payload),
        )

    return payloads_by_label


def _json_safe(value):
    try:
        import numpy as np

        if isinstance(value, (np.integer, np.floating)):
            return float(value)
    except Exception:
        pass
    return value


def _to_jsonable(value):
    if is_dataclass(value):
        value = asdict(value)
    elif hasattr(value, "__dict__"):
        value = value.__dict__
    if isinstance(value, dict):
        return {str(key): _to_jsonable(val) for key, val in value.items()}
    if isinstance(value, (list, tuple, set)):
        return [_to_jsonable(item) for item in value]
    return _json_safe(value)


def waterfall_payloads_to_json(payloads_by_label: dict[str, WaterfallPayload]) -> str:
    payloads_json = {}
    for label, payload in payloads_by_label.items():
        payload_dict = _to_jsonable(payload)
        if not isinstance(payload_dict, dict):
            payload_dict = {"value": payload_dict}
        payload_dict["checksum"] = payload_checksum(payload)
        payloads_json[str(label)] = payload_dict
    return json.dumps(payloads_json, indent=2, ensure_ascii=False)
