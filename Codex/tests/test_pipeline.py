from __future__ import annotations

import unittest
from pathlib import Path

from excel_pipeline.runner import run_pipeline_for_workbook


class PipelineSmokeTest(unittest.TestCase):
    def test_single_workbook_pipeline(self) -> None:
        excel_files = sorted(Path("ExcelFiles").glob("*.xlsx"))
        self.assertTrue(excel_files, "No .xlsx files found in ExcelFiles/")

        result = run_pipeline_for_workbook(
            source_workbook=excel_files[0],
            output_root=Path("artifacts_test"),
            cache_dir=Path(".cache/normalized"),
        )

        self.assertIn("artifacts", result)
        self.assertIn("mismatches", result)
        self.assertTrue(Path(result["artifacts"]["mapping_report"]).exists())
        self.assertTrue(Path(result["artifacts"]["structured_input"]).exists())


if __name__ == "__main__":
    unittest.main()
