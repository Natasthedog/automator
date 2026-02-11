from __future__ import annotations

import json
from dataclasses import asdict, is_dataclass

import pandas as pd

from deck.engine.waterfall.compute import (
    WaterfallPayload,
    _payload_checksum,
    compute_waterfall_payloads_for_all_labels as compute_with_engine,
)


def payload_checksum(payload: WaterfallPayload) -> float:
    return _payload_checksum(payload.series_values)


def compute_waterfall_payloads_for_all_labels(
    gathered_df: pd.DataFrame,
    scope_df: pd.DataFrame | None,
    bucket_data: dict | None,
    template_chart=None,
    target_labels: list[str] | None = None,
) -> dict[str, WaterfallPayload]:
    return compute_with_engine(
        gathered_df,
        scope_df,
        bucket_data,
        template_chart,
        target_labels=target_labels,
    )


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
