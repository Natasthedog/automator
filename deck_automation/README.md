# Waterfall Django Service Layer

## Call flow
1. `WaterfallOrchestrator.generate()` loads the template deck and asks `WaterfallPayloadBuilder` for label-specific payloads from `gatheredCN10`.
2. `WaterfallSlideMapper` discovers `<Waterfall Template>`, `<Waterfall Template2>`, ... and resolves each label using slide title/name fuzzy matching.
3. For each mapped slide the orchestrator:
   - ensures chart parts are unique,
   - updates placeholders/header text,
   - calls `WaterfallChartUpdater.update_slide_charts()`.
4. `WaterfallChartUpdater` performs `replace_data()` plus embedded workbook writeback and cache refresh so labels render on first open.
5. `ArtifactStore` writes `/tmp/decks/{job_id}/waterfall_output.pptx`, then promotes to durable storage.

## Async jobs
- `DeckGenerationJob` stores status + request/output metadata.
- `generate_waterfall_deck` task runs orchestration and updates status (`pending -> running -> succeeded/failed`).

## Run tests
```bash
pytest -q
```
