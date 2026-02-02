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

from .compute import _payload_checksum

def _json_safe(x):
    try:
        import numpy as np
        if isinstance(x, (np.integer, np.floating)):
            return float(x)
    except Exception:
        pass
    return x


def _to_jsonable(value):
    if is_dataclass(value):
        value = asdict(value)
    elif hasattr(value, "__dict__"):
        value = value.__dict__
    if isinstance(value, dict):
        return {str(key): _to_jsonable(val) for key, val in value.items()}
    if isinstance(value, (list, tuple, set)):
        return [_to_jsonable(item) for item in value]
    return _json_safe(value)


def _waterfall_payloads_to_json(payloads_by_label: dict[str, "WaterfallPayload"]) -> str:
    payloads_json = {}
    for label, payload in payloads_by_label.items():
        payload_dict = _to_jsonable(payload)
        if not isinstance(payload_dict, dict):
            payload_dict = {"value": payload_dict}
        checksum = _payload_checksum(getattr(payload, "series_values", []))
        payload_dict["checksum"] = checksum
        payloads_json[str(label)] = payload_dict
    return json.dumps(payloads_json, indent=2, ensure_ascii=False)
