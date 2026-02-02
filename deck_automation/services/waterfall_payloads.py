from __future__ import annotations

import json
import logging
from dataclasses import dataclass, asdict, is_dataclass

import pandas as pd

from deck.engine.waterfall.compute import (
    _payload_checksum,
    _waterfall_base_indices,
    _waterfall_base_values,
    _waterfall_series_from_gathered_df,
    compute_payload_for_label,
)
from deck.engine.waterfall.inject import _normalize_target_level_labels
from deck.engine.waterfall.targets import _target_level_labels_from_gathered_df_with_filters

logger = logging.getLogger(__name__)


@dataclass
class WaterfallPayload:
    categories: list[str]
    series_values: list[tuple[str, list[float]]]
    base_indices: tuple[int, int] | None
    base_values: tuple[float, float] | None
    gathered_label_values: dict[str, list]


def _wrap_payload(payload) -> WaterfallPayload:
    return WaterfallPayload(
        categories=list(getattr(payload, "categories", []) or []),
        series_values=[
            (name, list(values)) for name, values in getattr(payload, "series_values", [])
        ],
        base_indices=getattr(payload, "base_indices", None),
        base_values=getattr(payload, "base_values", None),
        gathered_label_values={
            key: list(values)
            for key, values in getattr(payload, "gathered_label_values", {}).items()
        },
    )


def _payload_from_gathered_only(
    gathered_df: pd.DataFrame,
    scope_df: pd.DataFrame | None,
    target_level_label: str,
    bucket_data: dict | None,
) -> WaterfallPayload:
    gathered_override = _waterfall_series_from_gathered_df(
        gathered_df,
        scope_df,
        target_level_label,
    )
    if gathered_override is None:
        raise ValueError(
            f"No gatheredCN10 waterfall data found for Target Level Label {target_level_label!r}."
        )
    categories, series_dict, gathered_label_values = gathered_override
    series_values: list[tuple[str, list[float]]] = []
    for key in ["Base", "Promo", "Media", "Blanks", "Positives", "Negatives"]:
        if key in series_dict:
            series_values.append((key, list(series_dict[key])))
    base_indices = _waterfall_base_indices(categories)
    base_values = None
    try:
        base_values = _waterfall_base_values(
            gathered_df,
            target_level_label,
            year1=bucket_data.get("year1") if bucket_data else None,
            year2=bucket_data.get("year2") if bucket_data else None,
        )
    except Exception as exc:
        logger.info(
            "Skipping waterfall base values for %r: %s",
            target_level_label,
            exc,
        )
    return WaterfallPayload(
        categories=list(categories),
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
    labels = _normalize_target_level_labels(target_labels)
    if not labels:
        labels = _target_level_labels_from_gathered_df_with_filters(
            gathered_df,
            year1=bucket_data.get("year1") if bucket_data else None,
            year2=bucket_data.get("year2") if bucket_data else None,
            target_labels=bucket_data.get("target_labels") if bucket_data else None,
        )
    payloads_by_label: dict[str, WaterfallPayload] = {}
    logger.info("Precomputing waterfall payloads for %d label(s).", len(labels))
    for label in labels:
        if template_chart is None:
            payload = _payload_from_gathered_only(
                gathered_df,
                scope_df,
                label,
                bucket_data,
            )
        else:
            computed = compute_payload_for_label(
                gathered_df,
                scope_df,
                label,
                bucket_data,
                template_chart,
            )
            payload = _wrap_payload(computed)
        payloads_by_label[label] = payload
        checksum = _payload_checksum(payload.series_values)
        logger.info(
            "Computed waterfall payload for %r: %d categories, checksum %.2f",
            label,
            len(payload.categories),
            checksum,
        )
    return payloads_by_label


def payload_checksum(payload: WaterfallPayload) -> float:
    return _payload_checksum(payload.series_values)


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
