from __future__ import annotations

import base64
import io

import pandas as pd
import pytest

from deck.engine.io_readers import product_description_df_from_contents


def _excel_contents(sheet_to_df: dict[str, pd.DataFrame]) -> str:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet_name, df in sheet_to_df.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    encoded = base64.b64encode(buffer.getvalue()).decode("ascii")
    return f"data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{encoded}"


def test_product_description_reader_accepts_productlist_fuzzy_sheet_name() -> None:
    contents = _excel_contents(
        {
            "PRODUCT DESCRIPTION": pd.DataFrame({"A": [1]}),
            "ProductList": pd.DataFrame({"Product": ["P1"]}),
        }
    )

    product_df = product_description_df_from_contents(contents, "scope.xlsx")

    assert product_df is not None
    assert list(product_df.columns) == ["A"]


def test_product_description_reader_prompts_for_product_list_when_missing() -> None:
    contents = _excel_contents(
        {
            "PRODUCT DESCRIPTION": pd.DataFrame({"A": [1]}),
            "Some Other Sheet": pd.DataFrame({"B": [2]}),
        }
    )

    with pytest.raises(ValueError) as exc:
        product_description_df_from_contents(contents, "scope.xlsx")

    message = str(exc.value)
    assert "Please identify which sheet corresponds to the Product List" in message
    assert "ProductList" in message
    assert "Product_List" in message
