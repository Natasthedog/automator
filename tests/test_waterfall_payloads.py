from __future__ import annotations

from pathlib import Path

import pandas as pd
from pptx import Presentation
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches

from deck.engine.waterfall.compute import (
    _payload_checksum,
    compute_waterfall_payloads_for_all_labels,
)


def _build_template_chart():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    chart_data = ChartData()
    chart_data.categories = ["Placeholder"]
    for name in ["Base", "Promo", "Media", "Blanks", "Positives", "Negatives"]:
        chart_data.add_series(name, (0,))
    chart_shape = slide.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_STACKED,
        Inches(1),
        Inches(1),
        Inches(6),
        Inches(3),
        chart_data,
    )
    return chart_shape.chart


def test_payloads_differ_across_labels():
    gathered_df = pd.read_excel(Path("tests/fixtures/gathered_min.xlsx"))
    template_chart = _build_template_chart()

    payloads = compute_waterfall_payloads_for_all_labels(
        gathered_df,
        scope_df=None,
        bucket_data=None,
        template_chart=template_chart,
    )

    assert set(payloads.keys()) == {"Alpha", "Beta"}

    checksums = [_payload_checksum(payload.series_values) for payload in payloads.values()]
    assert len(set(checksums)) > 1

    for payload in payloads.values():
        assert payload.categories
        assert payload.series_values
        assert all(values for _, values in payload.series_values)
