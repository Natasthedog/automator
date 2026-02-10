from __future__ import annotations

import io
import logging
import uuid

from django.http import Http404, HttpResponse
from django.shortcuts import render

from pptx import Presentation


logger = logging.getLogger(__name__)
SESSION_PAYLOADS_KEY = "deck_automation_payloads"


def _template_chart_from_pptx(uploaded_file):
    from deck.engine.pptx.charts import _waterfall_chart_from_slide
    from deck.engine.pptx.slides import _find_slide_by_marker

    data = uploaded_file.read()
    uploaded_file.seek(0)
    prs = Presentation(io.BytesIO(data))
    template_slide = _find_slide_by_marker(prs, "<Waterfall Template>")
    if template_slide is None:
        return None
    return _waterfall_chart_from_slide(template_slide, "Waterfall Template")


def deck_automation(request):
    context: dict[str, object] = {}
    if request.method == "POST":
        gathered_file = request.FILES.get("gathered_cn10")
        scope_file = request.FILES.get("scope_workbook")
        template_file = request.FILES.get("template_pptx")
        if not gathered_file:
            context["error"] = "Please upload the gatheredCN10 file to continue."
        else:
            try:
                from .services.readers import read_df, read_scope_df
                from .services.waterfall_payloads import (
                    compute_waterfall_payloads_for_all_labels,
                    payload_checksum,
                    waterfall_payloads_to_json,
                )

                gathered_df = read_df(gathered_file)
                scope_df = read_scope_df(scope_file) if scope_file else None
                template_chart = None
                if template_file:
                    try:
                        template_chart = _template_chart_from_pptx(template_file)
                        if template_chart is None:
                            logger.info(
                                "Template PPTX missing <Waterfall Template> marker or chart; using gathered data."
                            )
                    except Exception:
                        logger.exception("Failed to read template PPTX; using gathered data only.")
                        template_chart = None

                payloads_by_label = compute_waterfall_payloads_for_all_labels(
                    gathered_df,
                    scope_df,
                    bucket_data=None,
                    template_chart=template_chart,
                )
                payload_json = waterfall_payloads_to_json(payloads_by_label)

                download_id = str(uuid.uuid4())
                stored_payloads = request.session.get(SESSION_PAYLOADS_KEY, {})
                stored_payloads[download_id] = payload_json
                request.session[SESSION_PAYLOADS_KEY] = stored_payloads

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
                    }
                )
            except Exception as exc:
                logger.exception("Failed to compute waterfall payloads.")
                context["error"] = str(exc)

    return render(request, "deck_automation/deck_automation.html", context)


def download_payloads_json(request, download_id: str):
    payloads = request.session.get(SESSION_PAYLOADS_KEY, {})
    payload_json = payloads.get(download_id)
    if not payload_json:
        raise Http404("Payloads not found.")
    response = HttpResponse(payload_json, content_type="application/json")
    response["Content-Disposition"] = "attachment; filename=payloads.json"
    return response
