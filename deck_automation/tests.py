from __future__ import annotations

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

