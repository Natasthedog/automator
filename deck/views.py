import logging

from django.shortcuts import render

logger = logging.getLogger(__name__)
PROJECT_TEMPLATES = {}

def home(request):
    return render(request, "deck/home.html")

def home(request):
    return render(request, "deck/home.html")

def generate_deck(
    n_clicks,
    data_contents,
    data_name,
    scope_contents,
    scope_name,
    project_name,
    waterfall_targets,
    bucket_data,
):
    from dash import dcc, no_update

    from .engine.build import build_pptx_from_template
    from .engine.io_readers import (
        df_from_contents,
        product_description_df_from_contents,
        project_details_df_from_contents,
        scope_df_from_contents,
    )
    from .engine.project_details import _project_detail_value_from_df
    from .engine.waterfall.targets import target_brand_from_scope_df

    if not data_contents or not project_name:
        return no_update, "Please upload the data file and select a project."

    template_path = PROJECT_TEMPLATES.get(project_name)
    if not template_path or not template_path.exists():
        return no_update, "The selected project template could not be found."
    try:
        df = df_from_contents(data_contents, data_name)
        scope_df = None
        product_description_df = None
        project_details_df = None
        modelled_in_value = None
        metric_value = None
        if scope_contents:
            try:
                scope_df = scope_df_from_contents(scope_contents, scope_name)
            except Exception:
                scope_df = None
            try:
                product_description_df = product_description_df_from_contents(
                    scope_contents, scope_name
                )
            except Exception:
                product_description_df = None
            try:
                project_details_df = project_details_df_from_contents(
                    scope_contents, scope_name
                )
            except Exception:
                project_details_df = None
        if project_details_df is not None:
            modelled_in_value = _project_detail_value_from_df(
                project_details_df,
                "modelled in",
                [
                    "Sales will be modelled in",
                    "Sales will be modeled in",
                    "Sales modelled in",
                    "Sales modeled in",
                ],
                "Sales will be modelled in",
            )
            metric_value = _project_detail_value_from_df(
                project_details_df,
                "metric",
                [
                    "Volume metric (unique per dataset)",
                    "Volume metric unique per dataset",
                    "Volume metric",
                ],
                "Volume metric (unique per dataset)",
            )
        target_brand = target_brand_from_scope_df(scope_df)
        template_bytes = template_path.read_bytes()

        pptx_bytes = build_pptx_from_template(
            template_bytes,
            df,
            target_brand,
            project_name,
            scope_df,
            product_description_df,
            waterfall_targets,
            bucket_data,
            modelled_in_value,
            metric_value,
        )
        return dcc.send_bytes(lambda buff: buff.write(pptx_bytes), "deck.pptx"), "Building deck..."

    except Exception as exc:
        logger.exception("Deck generation failed.")
        message = str(exc).strip()
        if not message:
            message = "Unknown error. Check server logs for details."
        return no_update, f"Error ({type(exc).__name__}): {message}"

def download_waterfall_payloads(
    n_clicks,
    data_contents,
    data_name,
    scope_contents,
    scope_name,
    project_name,
    waterfall_targets,
    bucket_data,
):
    import io

    from dash import dcc, no_update
    from pptx import Presentation

    from .engine.io_readers import df_from_contents, scope_df_from_contents
    from .engine.pptx.charts import _waterfall_chart_from_slide
    from .engine.waterfall.compute import compute_waterfall_payloads_for_all_labels
    from .engine.waterfall.inject import _available_waterfall_template_slides
    from .engine.waterfall.payloads import _waterfall_payloads_to_json

    if not data_contents or not project_name:
        return no_update, "Please upload the data file and select a project."

    template_path = PROJECT_TEMPLATES.get(project_name)
    if not template_path or not template_path.exists():
        return no_update, "The selected project template could not be found."

    try:
        gathered_df = df_from_contents(data_contents, data_name)
        scope_df = None
        if scope_contents:
            try:
                scope_df = scope_df_from_contents(scope_contents, scope_name)
            except Exception:
                scope_df = None

        prs = Presentation(io.BytesIO(template_path.read_bytes()))
        available_slides = _available_waterfall_template_slides(prs)
        if not available_slides:
            return no_update, "Could not find the <Waterfall Template> slide in the template."
        template_chart = _waterfall_chart_from_slide(
            available_slides[0][1],
            "Waterfall Template",
        )
        if template_chart is None:
            return no_update, "Could not find the waterfall chart on the <Waterfall Template> slide."

        payloads_by_label = compute_waterfall_payloads_for_all_labels(
            gathered_df,
            scope_df,
            bucket_data,
            template_chart,
            target_labels=waterfall_targets,
        )
        payload_json = _waterfall_payloads_to_json(payloads_by_label)
        return dcc.send_string(payload_json, "waterfall_payloads.json"), "Prepared waterfall payload JSON."
    except Exception as exc:
        logger.exception("Waterfall payload JSON generation failed.")
        message = str(exc).strip()
        if not message:
            message = "Unknown error. Check server logs for details."
        return no_update, f"Error ({type(exc).__name__}): {message}"
