from __future__ import annotations

# Auto-generated split from dash_app.py

import io
import base64
import logging
import copy
from difflib import SequenceMatcher
from dataclasses import dataclass, asdict, is_dataclass
from datetime import date, timedelta
from pathlib import Path
import re
import json
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import range_boundaries, get_column_letter
import numbers
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.text import PP_ALIGN
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.oxml.ns import qn
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.parts.embeddedpackage import EmbeddedXlsxPart
from lxml import etree

logger = logging.getLogger(__name__)

from .waterfall.targets import _find_sheet_by_candidates

def bytes_from_contents(contents):
    _, content_string = contents.split(',')
    return base64.b64decode(content_string)


def df_from_contents(contents, filename):
    decoded = bytes_from_contents(contents)
    if filename.lower().endswith((".xlsx", ".xls", ".xlsb")):
        read_options = {}
        if filename.lower().endswith(".xlsb"):
            read_options["engine"] = "pyxlsb"
        return pd.read_excel(io.BytesIO(decoded), **read_options)
    elif filename.lower().endswith(".csv"):
        return pd.read_csv(io.StringIO(decoded.decode('utf-8')))
    else:
        raise ValueError("Unsupported file format. Please upload CSV or Excel.")


def raw_df_from_contents(contents, filename):
    decoded = bytes_from_contents(contents)
    if filename.lower().endswith((".xlsx", ".xls", ".xlsb")):
        read_options = {}
        if filename.lower().endswith(".xlsb"):
            read_options["engine"] = "pyxlsb"
        return pd.read_excel(io.BytesIO(decoded), header=None, **read_options)
    elif filename.lower().endswith(".csv"):
        return pd.read_csv(io.StringIO(decoded.decode("utf-8")), header=None)
    else:
        raise ValueError("Unsupported file format. Please upload CSV or Excel.")


def scope_df_from_contents(contents, filename):
    if not filename or not filename.lower().endswith((".xlsx", ".xlsb")):
        raise ValueError("Scope file must be an Excel workbook (.xlsx or .xlsb).")

    decoded = bytes_from_contents(contents)
    read_options = {}
    if filename.lower().endswith(".xlsb"):
        read_options["engine"] = "pyxlsb"
    scope_df = pd.read_excel(
        io.BytesIO(decoded),
        sheet_name="Overall Information",
        **read_options,
    )
    if scope_df.empty:
        return None
    return scope_df


def project_details_df_from_contents(contents, filename):
    if not filename or not filename.lower().endswith((".xlsx", ".xlsb")):
        raise ValueError("Scope file must be an Excel workbook (.xlsx or .xlsb).")

    decoded = bytes_from_contents(contents)
    read_options = {}
    if filename.lower().endswith(".xlsb"):
        read_options["engine"] = "pyxlsb"
    try:
        return pd.read_excel(
            io.BytesIO(decoded),
            sheet_name="Project Details",
            header=None,
            **read_options,
        )
    except ValueError:
        return None


def product_description_df_from_contents(contents, filename):
    if not filename or not filename.lower().endswith((".xlsx", ".xlsb")):
        raise ValueError("Scope file must be an Excel workbook (.xlsx or .xlsb).")

    decoded = bytes_from_contents(contents)
    read_options = {}
    if filename.lower().endswith(".xlsb"):
        read_options["engine"] = "pyxlsb"
    with pd.ExcelFile(io.BytesIO(decoded), **read_options) as excel_file:
        target_sheet = _find_sheet_by_candidates(
            excel_file.sheet_names, "PRODUCT DESCRIPTION"
        )
        if not target_sheet:
            return None

        product_list_sheet = _find_sheet_by_candidates(
            excel_file.sheet_names, "Product List"
        )
        if not product_list_sheet:
            product_list_sheet = _find_sheet_by_candidates(
                excel_file.sheet_names, "ProductList"
            )
        if not product_list_sheet:
            product_list_sheet = _find_sheet_by_candidates(
                excel_file.sheet_names, "Product_List"
            )
        if not product_list_sheet:
            available = ", ".join(excel_file.sheet_names)
            raise ValueError(
                "Could not find a Product List sheet. Please identify which sheet "
                "corresponds to the Product List (looked for Product List, "
                "ProductList, Product_List). Available sheets: "
                f"{available}"
            )

        product_df = excel_file.parse(target_sheet)
    if product_df.empty:
        return None
    return product_df
