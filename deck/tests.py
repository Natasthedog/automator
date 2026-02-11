from __future__ import annotations

import io

from django.core.files.uploadedfile import SimpleUploadedFile
from django.test import TestCase
from django.urls import reverse
from openpyxl import Workbook


class ProductDescriptionViewTests(TestCase):
    def _scope_upload(self, sheets: dict[str, list[str]]) -> SimpleUploadedFile:
        workbook = Workbook()
        default = workbook.active
        workbook.remove(default)
        for sheet_name, headers in sheets.items():
            worksheet = workbook.create_sheet(title=sheet_name)
            worksheet.append(headers)
            worksheet.append(["value" for _ in headers])

        buffer = io.BytesIO()
        workbook.save(buffer)
        return SimpleUploadedFile(
            "scope.xlsx",
            buffer.getvalue(),
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    def test_product_description_prompts_for_sheet_when_product_list_not_found(self):
        response = self.client.post(
            reverse("product-description"),
            data={
                "scope_workbook": self._scope_upload(
                    {
                        "PRODUCT DESCRIPTION": ["Name", "Value"],
                        "LookupSheet": ["Manufacturer", "Brand", "Subbrand"],
                    }
                )
            },
        )

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Please choose the correct sheet from the list")
        self.assertContains(response, "LookupSheet")

    def test_product_description_renders_three_rollup_dropdowns_and_saves_rollup(self):
        self.client.post(
            reverse("product-description"),
            data={
                "scope_workbook": self._scope_upload(
                    {
                        "PRODUCT DESCRIPTION": ["Name", "Value"],
                        "Product List": ["Manufacturer", "Brand", "Subbrand"],
                    }
                )
            },
        )

        response = self.client.post(
            reverse("product-description"),
            data={
                "product_list_sheet": "Product List",
                "rollups": ["Manufacturer_Brand", "Brand_Subbrand"],
            },
        )

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Build roll ups")
        self.assertContains(response, "Add another roll up")
        self.assertContains(response, "repeat(3, minmax(0, 1fr))")
        self.assertContains(response, "Manufacturer_Brand")
        self.assertContains(response, "Brand_Subbrand")
