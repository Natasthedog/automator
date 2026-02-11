import base64
import io
import logging

from django.shortcuts import redirect, render
from openpyxl import load_workbook

from .engine.waterfall.targets import _find_sheet_by_candidates

logger = logging.getLogger(__name__)
PROJECT_TEMPLATES = {}
PRODUCT_DESCRIPTION_SCOPE_KEY = "product_description_scope_workbook"
PRODUCT_DESCRIPTION_ROLLUPS_KEY = "product_description_rollups"
PRODUCT_DESCRIPTION_ROLLUP_ALIASES_KEY = "product_description_rollup_aliases"


def _product_list_columns_from_sheet(workbook_bytes: bytes, sheet_name: str) -> list[str]:
    workbook = load_workbook(io.BytesIO(workbook_bytes), data_only=True, read_only=True)
    worksheet = workbook[sheet_name]
    for row in worksheet.iter_rows(values_only=True):
        values = [str(cell).strip() for cell in row if cell is not None and str(cell).strip()]
        if values:
            return values
    return []


def _normalized_rollups(rollups: list[str]) -> list[str]:
    normalized: list[str] = []
    seen: set[str] = set()
    for rollup in rollups:
        value = (rollup or "").strip().strip("_")
        if not value or value in seen:
            continue
        seen.add(value)
        normalized.append(value)
    return normalized


def _rollup_alias_map(rollups: list[str], aliases: list[str]) -> dict[str, str]:
    result: dict[str, str] = {}
    for rollup, alias in zip(rollups, aliases):
        key = (rollup or "").strip()
        if not key:
            continue
        value = (alias or "").strip()
        result[key] = value or key
    return result


def home(request):
    return redirect("file-uploads")


def product_description(request):
    context: dict[str, object] = {
        "sheet_names": [],
        "product_list_columns": [],
        "selected_rollups": request.session.get(PRODUCT_DESCRIPTION_ROLLUPS_KEY, []),
        "selected_rollup_aliases": request.session.get(PRODUCT_DESCRIPTION_ROLLUP_ALIASES_KEY, {}),
        "selected_sheet": "",
        "selected_rollup_pairs": [],
    }

    stored_scope = request.session.get(PRODUCT_DESCRIPTION_SCOPE_KEY)
    scope_bytes: bytes | None = None
    if stored_scope:
        scope_bytes = base64.b64decode(stored_scope)

    selected_rollups = context["selected_rollups"]
    selected_rollup_aliases = context["selected_rollup_aliases"]
    context["selected_rollup_pairs"] = [
        {"value": rollup, "alias": selected_rollup_aliases.get(rollup, rollup)}
        for rollup in selected_rollups
    ]

    if request.method == "POST":
        uploaded_scope = request.FILES.get("scope_workbook")
        selected_sheet = (request.POST.get("product_list_sheet") or "").strip()
        submitted_rollups = _normalized_rollups(request.POST.getlist("rollups"))
        submitted_aliases = request.POST.getlist("rollup_aliases")

        if uploaded_scope:
            scope_bytes = uploaded_scope.read()
            request.session[PRODUCT_DESCRIPTION_SCOPE_KEY] = base64.b64encode(scope_bytes).decode("ascii")
            request.session.modified = True

        if not scope_bytes:
            context["error"] = "Please upload a scope workbook to configure Product Description roll ups."
            return render(request, "deck/PRODUCT_DESCRIPTION.html", context)

        workbook = load_workbook(io.BytesIO(scope_bytes), data_only=True, read_only=True)
        sheet_names = workbook.sheetnames
        context["sheet_names"] = sheet_names

        if not selected_sheet:
            selected_sheet = _find_sheet_by_candidates(sheet_names, "Product List") or ""
            if not selected_sheet:
                selected_sheet = _find_sheet_by_candidates(sheet_names, "ProductList") or ""
            if not selected_sheet:
                selected_sheet = _find_sheet_by_candidates(sheet_names, "Product_List") or ""

        context["selected_sheet"] = selected_sheet

        if not selected_sheet:
            context["warning"] = (
                "We could not identify the Product List sheet automatically. "
                "Please choose the correct sheet from the list."
            )
            return render(request, "deck/PRODUCT_DESCRIPTION.html", context)

        if selected_sheet not in sheet_names:
            context["error"] = "The selected Product List sheet does not exist in this workbook."
            return render(request, "deck/PRODUCT_DESCRIPTION.html", context)

        columns = _product_list_columns_from_sheet(scope_bytes, selected_sheet)
        context["product_list_columns"] = columns

        if submitted_rollups:
            alias_map = _rollup_alias_map(submitted_rollups, submitted_aliases)
            request.session[PRODUCT_DESCRIPTION_ROLLUPS_KEY] = submitted_rollups
            request.session[PRODUCT_DESCRIPTION_ROLLUP_ALIASES_KEY] = alias_map
            request.session.modified = True
            context["selected_rollups"] = submitted_rollups
            context["selected_rollup_aliases"] = alias_map
            context["message"] = "Roll up selection saved for this session."
            context["selected_rollup_pairs"] = [
                {"value": rollup, "alias": alias_map.get(rollup, rollup)}
                for rollup in submitted_rollups
            ]

    return render(request, "deck/PRODUCT_DESCRIPTION.html", context)


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
