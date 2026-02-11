from __future__ import annotations

import logging
from pathlib import Path
from uuid import uuid4

from django.http import Http404, HttpResponse
from django.shortcuts import render
from pptx import Presentation

from deck_automation.services.readers import read_df, read_scope_df
from deck_automation.services.waterfall_payloads import (
    compute_waterfall_payloads_for_all_labels,
    payload_checksum,
    waterfall_payloads_to_json,
)

logger = logging.getLogger(__name__)

SESSION_PAYLOADS_KEY = "deck_automation_payloads"
WATERFALL_TEMPLATE_MARKER = "<Waterfall Template>"
TEMPLATE_OPTIONS = {
    "MMx": "MMx.pptx",
    "MMM": "MMM.pptx",
    "PnP": "PnP.pptx",
}


def _shape_matches_marker(shape, marker_text: str) -> bool:
    marker_text = marker_text.strip()
    shape_name = getattr(shape, "name", "") or ""
    if marker_text and marker_text in shape_name:
        return True
    if shape.has_text_frame:
        return marker_text in (shape.text_frame.text or "")
    return False


def _find_template_chart(template_pptx):
    if template_pptx is None:
        return None
    template_pptx.seek(0)
    prs = Presentation(template_pptx)
    for slide in prs.slides:
        if not any(_shape_matches_marker(shape, WATERFALL_TEMPLATE_MARKER) for shape in slide.shapes):
            continue
        for shape in slide.shapes:
            if shape.has_chart:
                return shape.chart
        logger.info("Found waterfall template slide marker but no chart shape.")
        return None
    logger.info("No <Waterfall Template> marker found in uploaded template.")
    return None


def deck_automation(request):
    context: dict[str, object] = {
        "template_options": sorted(TEMPLATE_OPTIONS.keys()),
        "selected_template": "MMx",
    }
    if request.method == "POST":
        gathered_file = request.FILES.get("gathered_cn10")
        scope_file = request.FILES.get("scope_workbook")
        selected_template = request.POST.get("template_choice", "").strip()
        context["selected_template"] = selected_template

        template_filename = TEMPLATE_OPTIONS.get(selected_template)
        if not template_filename:
            context["error"] = "Please select a deck template to continue."
            return render(request, "deck_automation/deck_automation.html", context)

        if not gathered_file:
            context["error"] = "Please upload the gatheredCN10 file to continue."
            return render(request, "deck_automation/deck_automation.html", context)

        try:
            template_path = Path(__file__).resolve().parent / "templates" / "deck_automation" / template_filename
            if not template_path.exists():
                raise FileNotFoundError(f"Template not found: {template_filename}")

            gathered_df = read_df(gathered_file)
            scope_df = read_scope_df(scope_file)
            with template_path.open("rb") as template_file:
                template_chart = _find_template_chart(template_file)
            payloads_by_label = compute_waterfall_payloads_for_all_labels(
                gathered_df,
                scope_df,
                bucket_data=None,
                template_chart=template_chart,
            )
            payload_json = waterfall_payloads_to_json(payloads_by_label)

            download_id = str(uuid4())
            payloads = request.session.get(SESSION_PAYLOADS_KEY, {})
            payloads[download_id] = payload_json
            request.session[SESSION_PAYLOADS_KEY] = payloads
            request.session.modified = True

            summaries = []
            for label, payload in payloads_by_label.items():
                summaries.append(
                    {
                        "label": label,
                        "category_count": len(payload.categories),
                        "checksum": payload_checksum(payload),
                    }
                )

            context.update(
                {
                    "computed_count": len(payloads_by_label),
                    "payload_summaries": summaries,
                    "download_id": download_id,
                    "message": "Payload computation complete.",
                    "selected_template": selected_template,
                }
            )
        except Exception as exc:  # noqa: BLE001
            logger.exception("Deck automation payload computation failed")
            context["error"] = str(exc) or "We could not compute payloads from the uploaded file."

    return render(request, "deck_automation/deck_automation.html", context)


def download_payloads_json(request, download_id: str):
    payloads = request.session.get(SESSION_PAYLOADS_KEY, {})
    payload_json = payloads.get(str(download_id))
    if not payload_json:
        raise Http404("Payloads not found.")
    response = HttpResponse(payload_json, content_type="application/json")
    response["Content-Disposition"] = "attachment; filename=payloads.json"
    return response
