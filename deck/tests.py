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

    def _scope_upload_with_rows(self, sheets: dict[str, list[list[object]]]) -> SimpleUploadedFile:
        workbook = Workbook()
        default = workbook.active
        workbook.remove(default)
        for sheet_name, rows in sheets.items():
            worksheet = workbook.create_sheet(title=sheet_name)
            for row in rows:
                worksheet.append(row)

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

    def test_product_description_renders_single_rollup_dropdown_and_saves_rollup(self):
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
                "ppg_correspondence_sheet": "PRODUCT DESCRIPTION",
                "rollups": ["Manufacturer", "Brand"],
                "rollup_aliases": ["MFR", "BRAND"],
            },
        )

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Build roll ups")
        self.assertContains(response, "Add another roll up")
        self.assertContains(response, "minmax(0, 1fr) minmax(0, 1fr) auto")
        self.assertContains(response, "Manufacturer")
        self.assertContains(response, "Brand")
        self.assertContains(response, "Rename roll up")
        self.assertContains(response, "MFR")
        self.assertContains(response, "BRAND")
        self.assertContains(response, "PPG_ID column")
        self.assertContains(response, "PPG_NAME column")
        self.assertContains(response, "PPG_EAN_CORRESPONDENCE sheet")
        self.assertContains(response, "Generate PRODUCT_DESCRIPTION & Download Scope")

    def test_generate_scope_downloads_product_description_sheet(self):
        self.client.post(
            reverse("product-description"),
            data={
                "scope_workbook": self._scope_upload_with_rows(
                    {
                        "Product List": [
                            ["EAN", "Manufacturer", "Brand", "PackType"],
                            ["111", "Mfr A", "Brand A", "Bottle"],
                        ],
                        "PPG_EAN_CORRESPONDENCE": [
                            ["PPG_ID", "PPG_NAME", "EAN"],
                            ["P1", "PPG One", "111"],
                        ],
                    }
                )
            },
        )

        response = self.client.post(
            reverse("product-description"),
            data={
                "product_list_sheet": "Product List",
                "ppg_correspondence_sheet": "PPG_EAN_CORRESPONDENCE",
                "rollups": ["Manufacturer", "PackType"],
                "rollup_aliases": ["MANUFACTURER", "PACK_TYPE"],
                "action": "generate_scope",
            },
        )

        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            response["Content-Type"],
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        self.assertIn(
            'attachment; filename="scope_with_product_description.xlsx"',
            response["Content-Disposition"],
        )

    def test_generate_scope_warns_when_ppg_id_or_name_missing_and_not_selected(self):
        self.client.post(
            reverse("product-description"),
            data={
                "scope_workbook": self._scope_upload_with_rows(
                    {
                        "Product List": [
                            ["EAN", "Manufacturer"],
                            ["111", "Mfr A"],
                        ],
                        "PPG_EAN_CORRESPONDENCE": [
                            ["GROUP_ID", "GROUP_NAME", "EAN"],
                            ["P1", "PPG One", "111"],
                        ],
                    }
                )
            },
        )

        response = self.client.post(
            reverse("product-description"),
            data={
                "product_list_sheet": "Product List",
                "ppg_correspondence_sheet": "PPG_EAN_CORRESPONDENCE",
                "rollups": ["Manufacturer"],
                "rollup_aliases": ["MANUFACTURER"],
                "action": "generate_scope",
            },
        )

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "Please identify those columns before generating PRODUCT_DESCRIPTION")
