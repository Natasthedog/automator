from __future__ import annotations

from pathlib import Path

import pandas as pd

try:
    from celery import shared_task
except Exception:  # pragma: no cover
    def shared_task(func):
        return func

from deck_automation.models import DeckGenerationJob
from deck_automation.services.waterfall_service_layer import (
    WaterfallGenerationRequest,
    WaterfallOrchestrator,
)


@shared_task
def generate_waterfall_deck(job_pk: int) -> dict[str, str]:
    job = DeckGenerationJob.objects.get(pk=job_pk)
    payload = job.request_payload or {}
    job.status = DeckGenerationJob.Status.RUNNING
    job.error_message = ""
    job.save(update_fields=["status", "error_message", "updated_at"])
    try:
        gathered_df = pd.DataFrame(payload.get("gathered_rows", []))
        scope_rows = payload.get("scope_rows") or []
        scope_df = pd.DataFrame(scope_rows) if scope_rows else None
        request = WaterfallGenerationRequest(
            template_path=Path(payload["template_path"]),
            gathered_df=gathered_df,
            scope_df=scope_df,
            target_labels=payload.get("target_labels"),
            bucket_data=payload.get("bucket_data"),
            modelled_in_value=payload.get("modelled_in_value"),
            metric_value=payload.get("metric_value"),
            job_id=str(job.job_id),
        )
        result = WaterfallOrchestrator().generate(request)
    except Exception as exc:  # noqa: BLE001
        job.status = DeckGenerationJob.Status.FAILED
        job.error_message = str(exc)
        job.save(update_fields=["status", "error_message", "updated_at"])
        raise
    job.status = DeckGenerationJob.Status.SUCCEEDED
    job.output_storage_key = result.storage_key
    job.save(update_fields=["status", "output_storage_key", "updated_at"])
    return {"job_id": str(job.job_id), "storage_key": result.storage_key}
