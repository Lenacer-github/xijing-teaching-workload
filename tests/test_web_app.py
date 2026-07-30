"""Bootstrap/Flask 版本的关键回归测试。"""

from __future__ import annotations

import io
import json
import re
import unittest

from openpyxl import load_workbook

from web_app import app
from workload_domain import WORKLOAD_DATA


class WebAppTestCase(unittest.TestCase):
    def setUp(self):
        app.config.update(TESTING=True)
        self.client = app.test_client()
        self.category, standards = next(iter(WORKLOAD_DATA.items()))
        self.standard = standards[0]
        self.payload = {
            "dept": "数字商务系",
            "name": "测试教师",
            "title": "讲师",
            "term": "2025-2026学年第二学期",
            "rows": [
                {
                    "category": self.category,
                    "standard": self.standard,
                    "workload": "12.5",
                    "time": "2026-07-30",
                    "remark": "测试成果说明",
                }
            ],
        }

    def test_page_and_health(self):
        page = self.client.get("/")
        self.assertEqual(page.status_code, 200)
        page_text = page.get_data(as_text=True)
        self.assertIn("教师教学工作量填报系统", page_text)
        self.assertIn("record-template", page_text)
        self.assertIn("2.0.0", page_text)
        config_match = re.search(
            r"window\.APP_CONFIG = (.*?);", page_text, flags=re.DOTALL
        )
        self.assertIsNotNone(config_match)
        rendered_config = json.loads(config_match.group(1))
        self.assertEqual(rendered_config["categories"], list(WORKLOAD_DATA))

        health = self.client.get("/health")
        self.assertEqual(health.status_code, 200)
        self.assertEqual(health.get_json()["version"], "2.0.0")

    def test_restore_normalizes_progress(self):
        response = self.client.post("/api/restore", json=self.payload)
        self.assertEqual(response.status_code, 200)
        restored = response.get_json()["data"]
        self.assertEqual(restored["rows"][0]["workload"], 12.5)
        self.assertEqual(restored["rows"][0]["time"], "2026-07-30")

    def test_restore_allows_empty_department_for_draft(self):
        draft = {**self.payload, "dept": ""}
        response = self.client.post("/api/restore", json=draft)
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.get_json()["data"]["dept"], "")

    def test_export_returns_valid_excel(self):
        response = self.client.post("/api/export", json=self.payload)
        self.assertEqual(response.status_code, 200)
        self.assertTrue(response.data.startswith(b"PK"))
        workbook = load_workbook(io.BytesIO(response.data), data_only=False)
        sheet = workbook.active
        self.assertEqual(sheet["C6"].value, 12.5)
        self.assertEqual(sheet["C6"].number_format, "General")
        self.assertTrue(sheet.protection.sheet)

    def test_export_rejects_missing_required_fields(self):
        invalid = {**self.payload, "name": ""}
        response = self.client.post("/api/export", json=invalid)
        self.assertEqual(response.status_code, 400)
        self.assertIn("姓名", response.get_json()["message"])

    def test_export_rejects_duplicate_standard(self):
        duplicate = {
            **self.payload,
            "rows": self.payload["rows"] * 2,
        }
        response = self.client.post("/api/export", json=duplicate)
        self.assertEqual(response.status_code, 400)
        self.assertIn("重复", response.get_json()["message"])


if __name__ == "__main__":
    unittest.main()
