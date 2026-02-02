from __future__ import annotations

import io
import logging
from typing import IO

import pandas as pd

logger = logging.getLogger(__name__)


SUPPORTED_EXTENSIONS = (".xls", ".xlsx", ".xlsb", ".csv")


def _read_bytes(uploaded_file: IO[bytes]) -> bytes:
    uploaded_file.seek(0)
    data = uploaded_file.read()
    uploaded_file.seek(0)
    return data


def read_df(uploaded_file: IO[bytes]) -> pd.DataFrame:
    if uploaded_file is None:
        raise ValueError("No file provided.")
    filename = getattr(uploaded_file, "name", "") or ""
    lower_name = filename.lower()
    if not lower_name.endswith(SUPPORTED_EXTENSIONS):
        raise ValueError("Unsupported file format. Please upload CSV or Excel.")

    data = _read_bytes(uploaded_file)
    if lower_name.endswith(".csv"):
        try:
            text = data.decode("utf-8")
        except UnicodeDecodeError:
            text = data.decode("latin-1")
        return pd.read_csv(io.StringIO(text))

    read_options: dict[str, object] = {}
    if lower_name.endswith(".xlsb"):
        read_options["engine"] = "pyxlsb"
    return pd.read_excel(io.BytesIO(data), **read_options)


def read_scope_df(uploaded_file: IO[bytes]) -> pd.DataFrame | None:
    if uploaded_file is None:
        return None
    filename = getattr(uploaded_file, "name", "") or ""
    lower_name = filename.lower()
    if not lower_name.endswith((".xlsx", ".xlsb")):
        raise ValueError("Scope file must be an Excel workbook (.xlsx or .xlsb).")

    data = _read_bytes(uploaded_file)
    read_options: dict[str, object] = {}
    if lower_name.endswith(".xlsb"):
        read_options["engine"] = "pyxlsb"
    try:
        scope_df = pd.read_excel(
            io.BytesIO(data),
            sheet_name="Overall Information",
            **read_options,
        )
    except ValueError:
        logger.info("Scope workbook missing 'Overall Information' sheet.")
        return None
    if scope_df.empty:
        return None
    return scope_df
