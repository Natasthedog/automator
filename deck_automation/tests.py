from __future__ import annotations

import json
from io import BytesIO

import pandas as pd
from django.core.files.uploadedfile import SimpleUploadedFile
from django.test import SimpleTestCase, TestCase
from django.urls import reverse

from .services.readers import read_df


class DeckAutomationViewsTests(TestCase):
    def test_deck_automation_page_loads(self):
        response = self.client.get(reverse("deck-automation"))

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Deck Automation (MVP)")

    def test_root_redirects_to_deck_automation(self):
        response = self.client.get("/")

        self.assertEqual(response.status_code, 302)
        self.assertEqual(response.headers.get("Location"), "/deck-automation/")

    def test_post_computes_payloads(self):
        gathered_df = _build_minimal_gathered_df()
        uploaded = _make_xlsx_upload(gathered_df, "gathered.xlsx")

        response = self.client.post(
            reverse("deck-automation"),
            data={"gathered_cn10": uploaded},
        )

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Computed")
        self.assertIn("download_id", response.context)

    def test_download_payloads_json(self):
        gathered_df = _build_minimal_gathered_df()
        uploaded = _make_xlsx_upload(gathered_df, "gathered.xlsx")

        response = self.client.post(
            reverse("deck-automation"),
            data={"gathered_cn10": uploaded},
        )

        download_id = response.context["download_id"]
        download_url = reverse("deck-automation-download", args=[download_id])
        download_response = self.client.get(download_url)

        self.assertEqual(download_response.status_code, 200)
        self.assertEqual(download_response.headers.get("Content-Type"), "application/json")
        payload_data = json.loads(download_response.content.decode("utf-8"))
        self.assertIn("Alpha", payload_data)
        self.assertIn("Beta", payload_data)


class DeckAutomationReadersTests(SimpleTestCase):
    def test_read_df_handles_csv_and_xlsx(self):
        df = pd.DataFrame({"A": [1, 2], "B": ["x", "y"]})

        csv_bytes = df.to_csv(index=False).encode("utf-8")
        csv_upload = SimpleUploadedFile("data.csv", csv_bytes, content_type="text/csv")
        csv_df = read_df(csv_upload)
        self.assertListEqual(list(csv_df.columns), ["A", "B"])

        xlsx_upload = _make_xlsx_upload(df, "data.xlsx")
        xlsx_df = read_df(xlsx_upload)
        self.assertListEqual(list(xlsx_df.columns), ["A", "B"])


def _make_xlsx_upload(df: pd.DataFrame, filename: str) -> SimpleUploadedFile:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return SimpleUploadedFile(
        filename,
        buffer.getvalue(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def _build_minimal_gathered_df() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Target Level Label": "Alpha",
                "Target Label": "Own",
                "Year": "Year1",
                "Actuals": 100,
                "Vars": "Var1",
                "Base": 10,
                "Promo": 1,
                "Media": 0,
                "Blanks": 0,
                "Positives": 2,
                "Negatives": -1,
            },
            {
                "Target Level Label": "Alpha",
                "Target Label": "Own",
                "Year": "Year2",
                "Actuals": 120,
                "Vars": "Var2",
                "Base": 12,
                "Promo": 2,
                "Media": 0,
                "Blanks": 0,
                "Positives": 3,
                "Negatives": -2,
            },
            {
                "Target Level Label": "Beta",
                "Target Label": "Own",
                "Year": "Year1",
                "Actuals": 80,
                "Vars": "Var1",
                "Base": 8,
                "Promo": 1,
                "Media": 0,
                "Blanks": 0,
                "Positives": 1,
                "Negatives": -1,
            },
            {
                "Target Level Label": "Beta",
                "Target Label": "Own",
                "Year": "Year2",
                "Actuals": 90,
                "Vars": "Var2",
                "Base": 9,
                "Promo": 1,
                "Media": 0,
                "Blanks": 0,
                "Positives": 2,
                "Negatives": -1,
            },
        ]
    )
