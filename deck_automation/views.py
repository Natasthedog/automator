from __future__ import annotations

from django.http import Http404, HttpResponse
from django.shortcuts import render


SESSION_PAYLOADS_KEY = "deck_automation_payloads"


def deck_automation(request):
    context: dict[str, object] = {}
    if request.method == "POST":
        gathered_file = request.FILES.get("gathered_cn10")
        if not gathered_file:
            context["error"] = "Please upload the gatheredCN10 file to continue."
        else:
            context["message"] = (
                "Upload received. Processing will be wired up in a follow-up change."
            )
    return render(request, "deck_automation/deck_automation.html", context)


def download_payloads_json(request, download_id: str):
    payloads = request.session.get(SESSION_PAYLOADS_KEY, {})
    payload_json = payloads.get(download_id)
    if not payload_json:
        raise Http404("Payloads not found.")
    response = HttpResponse(payload_json, content_type="application/json")
    response["Content-Disposition"] = "attachment; filename=payloads.json"
    return response
