from __future__ import annotations

import logging

from .compute import WaterfallPayload, _payload_checksum

logger = logging.getLogger(__name__)


def validate_payloads_or_raise(
    payloads_by_label: dict[str, WaterfallPayload],
    expected_labels: list[str],
) -> None:
    missing = [label for label in expected_labels if label not in payloads_by_label]
    if missing:
        raise ValueError(
            "Missing waterfall payloads for Target Level Label(s): "
            + ", ".join(repr(label) for label in missing)
        )

    checksums: list[float] = []
    for label in expected_labels:
        payload = payloads_by_label[label]
        categories = getattr(payload, "categories", None) or []
        series_values = getattr(payload, "series_values", None) or []
        if not categories:
            raise ValueError(
                f"Waterfall payload for {label!r} has no categories."
            )
        if not series_values:
            raise ValueError(
                f"Waterfall payload for {label!r} has no series values."
            )
        if any(not values for _, values in series_values):
            raise ValueError(
                f"Waterfall payload for {label!r} has empty series values."
            )
        checksums.append(_payload_checksum(series_values))

    if len(checksums) >= 2 and len(set(checksums)) == 1:
        raise ValueError(
            "Waterfall payloads for multiple Target Level Labels are identical; "
            "check gathered data filtering."
        )
    logger.info(
        "Validated %d waterfall payload(s); checksum range %.2f-%.2f.",
        len(checksums),
        min(checksums) if checksums else 0.0,
        max(checksums) if checksums else 0.0,
    )
