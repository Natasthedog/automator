from __future__ import annotations

import io
import logging
from csv import Sniffer
from typing import IO

import pandas as pd

logger = logging.getLogger(__name__)


SUPPORTED_EXTENSIONS = (".xls", ".xlsx", ".xlsb", ".csv", ".tsv")
TARGET_LEVEL_LABEL_COLUMN = "Target Level Label"


def _find_header_row_index(df: pd.DataFrame, marker: str = TARGET_LEVEL_LABEL_COLUMN) -> int:
    marker_normalized = marker.strip().casefold()
    for row_idx in range(len(df.index)):
        row_values = df.iloc[row_idx].tolist()
        for value in row_values:
            if str(value or "").strip().casefold() == marker_normalized:
                return row_idx
    return 0


def _apply_detected_header_row(df: pd.DataFrame) -> pd.DataFrame:
    header_row_index = _find_header_row_index(df)
    if len(df.index) == 0:
        out = df.copy()
        out.attrs["detected_header_row_index"] = 0
        return out
    header_values = df.iloc[header_row_index].tolist()
    data_df = df.iloc[header_row_index + 1 :].copy()
    data_df.columns = header_values
    out = data_df.reset_index(drop=True)
    out.attrs["detected_header_row_index"] = int(header_row_index)
    return out


def _sniff_delimiter(text: str) -> str:
    sample = text[:2048]
    try:
        return Sniffer().sniff(sample, delimiters=",\t;|").delimiter
    except Exception:
        return "\t" if "\t" in sample else ","


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
    if lower_name.endswith((".csv", ".tsv")):
        try:
            text = data.decode("utf-8")
        except UnicodeDecodeError:
            text = data.decode("latin-1")
        delimiter = _sniff_delimiter(text)
        raw_df = pd.read_csv(io.StringIO(text), delimiter=delimiter, header=None)
        return _apply_detected_header_row(raw_df)

    read_options: dict[str, object] = {}
    if lower_name.endswith(".xlsb"):
        read_options["engine"] = "pyxlsb"
    raw_df = pd.read_excel(io.BytesIO(data), header=None, **read_options)
    return _apply_detected_header_row(raw_df)


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
