"""Tests for the excel_to_mapping module."""

import os
import shutil
import sys
import tempfile
import unittest

from openpyxl import Workbook, load_workbook

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from tests.create_sample_workbook import create_sample_workbook
from excel_to_mapping.mapper import (
    generate_mapping_report,
    _build_all_referenced_cells,
    _classify_formula_cells,
    _build_sheet_mapping,
    _grouped_descriptions,
)
from excel_to_python import parse_workbook, classify_cells


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _create_simple_workbook(path):
    """Workbook with inputs, intermediate calcs, and final outputs."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["A1"] = "Price"
    ws["A2"] = 10
    ws["B1"] = "Qty"
    ws["B2"] = 5
    ws["C1"] = "Subtotal"
    ws["C2"] = "=A2*B2"       # intermediate calc (referenced by D2)
    ws["D1"] = "Total"
    ws["D2"] = "=C2+100"      # output (not referenced by anything)
    wb.save(path)
    wb.close()


def _create_vertical_drag_workbook(path):
    """Workbook with a dragged formula column and a SUM output."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales"
    ws["A1"] = "Revenue"
    ws["B1"] = "Cost"
    ws["C1"] = "Profit"
    for r in range(2, 7):
        ws[f"A{r}"] = 1000 * r
        ws[f"B{r}"] = 600 * r
        ws[f"C{r}"] = f"=A{r}-B{r}"   # dragged formula (intermediate)
    ws["C7"] = "=SUM(C2:C6)"          # output
    wb.save(path)
    wb.close()


def _create_multi_sheet_workbook(path):
    """Workbook with two sheets and cross-sheet references."""
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Input"
    ws1["A1"] = "Val"
    ws1["A2"] = 42

    ws2 = wb.create_sheet("Calc")
    ws2["A1"] = "Doubled"
    ws2["A2"] = "=Input!A2*2"      # calc (referenced by B2)
    ws2["B1"] = "Final"
    ws2["B2"] = "=A2+10"           # output
    wb.save(path)
    wb.close()


def _create_horizontal_drag_workbook(path):
    """Workbook with horizontally-dragged formulas."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Forecast"
    for ci in range(5):
        col = chr(66 + ci)  # B..F
        ws[f"{col}1"] = 100 + ci * 10
        ws[f"{col}2"] = f"={col}1*1.1"  # dragged horizontally
    ws["A1"] = "Base"
    ws["A2"] = "Grown"
    wb.save(path)
    wb.close()


def _create_all_formulas_workbook(path):
    """Workbook where every cell is a formula (no hardcoded inputs)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Formulas"
    ws["A1"] = "=1+1"
    ws["A2"] = "=A1*2"
    wb.save(path)
    wb.close()


def _create_no_formulas_workbook(path):
    """Workbook with no formulas at all."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Static"
    ws["A1"] = "Hello"
    ws["B1"] = 42
    wb.save(path)
    wb.close()


# ---------------------------------------------------------------------------
# Unit tests: classification helpers
# ---------------------------------------------------------------------------

class TestBuildAllReferencedCells(unittest.TestCase):
    def test_simple_references(self):
        formulas = [
            ("Sheet1", "C", 2, "=A2*B2", {}),
            ("Sheet1", "D", 2, "=C2+100", {}),
        ]
        refs = _build_all_referenced_cells(formulas)
        self.assertIn(("Sheet1", "A", 2), refs)
        self.assertIn(("Sheet1", "B", 2), refs)
        self.assertIn(("Sheet1", "C", 2), refs)
        self.assertNotIn(("Sheet1", "D", 2), refs)

    def test_range_references_expanded(self):
        formulas = [
            ("Sheet1", "C", 7, "=SUM(C2:C6)", {}),
        ]
        refs = _build_all_referenced_cells(formulas)
        for r in range(2, 7):
            self.assertIn(("Sheet1", "C", r), refs)

    def test_cross_sheet_reference(self):
        formulas = [
            ("Calc", "A", 2, "=Input!A2*2", {}),
        ]
        refs = _build_all_referenced_cells(formulas)
        self.assertIn(("Input", "A", 2), refs)


class TestClassifyFormulaCells(unittest.TestCase):
    def test_intermediate_vs_output(self):
        formulas = [
            ("S", "C", 2, "=A2*B2", {}),   # referenced by D2 → calculation
            ("S", "D", 2, "=C2+100", {}),   # not referenced → output
        ]
        calcs, outputs = _classify_formula_cells(formulas)
        calc_keys = {(s, c, r) for s, c, r, *_ in calcs}
        out_keys = {(s, c, r) for s, c, r, *_ in outputs}
        self.assertIn(("S", "C", 2), calc_keys)
        self.assertIn(("S", "D", 2), out_keys)

    def test_all_outputs_when_no_inter_deps(self):
        formulas = [
            ("S", "C", 2, "=A2*B2", {}),
            ("S", "D", 2, "=A2+B2", {}),
        ]
        calcs, outputs = _classify_formula_cells(formulas)
        # Neither C2 nor D2 is referenced by the other
        self.assertEqual(len(calcs), 0)
        self.assertEqual(len(outputs), 2)


# ---------------------------------------------------------------------------
# Unit tests: grouped descriptions
# ---------------------------------------------------------------------------

class TestGroupedDescriptions(unittest.TestCase):
    def test_vertical_group_collapsed(self):
        cells = [
            ("Sales", "C", r, f"=A{r}-B{r}", {})
            for r in range(2, 7)
        ]
        descs = _grouped_descriptions(cells)
        # 5 dragged cells should collapse into one group description
        self.assertEqual(len(descs), 1)
        self.assertIn("5 cells", descs[0])
        self.assertIn("vertical", descs[0])

    def test_single_formula_not_collapsed(self):
        cells = [("S", "D", 2, "=SUM(C2:C6)", {})]
        descs = _grouped_descriptions(cells)
        self.assertEqual(len(descs), 1)
        self.assertIn("SUM", descs[0])

    def test_empty_input(self):
        self.assertEqual(_grouped_descriptions([]), [])


# ---------------------------------------------------------------------------
# Integration: _build_sheet_mapping
# ---------------------------------------------------------------------------

class TestBuildSheetMapping(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmpdir = tempfile.mkdtemp()
        cls.path = os.path.join(cls.tmpdir, "simple.xlsx")
        _create_simple_workbook(cls.path)

        wb = load_workbook(cls.path)
        cls.sheets, cls.tables = parse_workbook(wb)
        cls.formula_cells, cls.hardcoded_cells = classify_cells(
            cls.sheets, cls.tables)
        wb.close()

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmpdir)

    def test_inputs_present(self):
        inputs, _, _ = _build_sheet_mapping(
            "Data", self.hardcoded_cells, self.formula_cells)
        # Should include Price, 10, Qty, 5 header values etc.
        values = [v for _, _, v in inputs]
        self.assertIn(10, values)
        self.assertIn(5, values)

    def test_calculations_and_outputs_split(self):
        _, calcs, outputs = _build_sheet_mapping(
            "Data", self.hardcoded_cells, self.formula_cells)
        # C2 is intermediate, D2 is output
        calc_text = " ".join(calcs)
        out_text = " ".join(outputs)
        self.assertIn("C2", calc_text)
        self.assertIn("D2", out_text)


# ---------------------------------------------------------------------------
# End-to-end: report generation
# ---------------------------------------------------------------------------

class TestGenerateMappingReportSimple(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmpdir = tempfile.mkdtemp()
        cls.wb_path = os.path.join(cls.tmpdir, "simple.xlsx")
        _create_simple_workbook(cls.wb_path)

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmpdir)

    def test_report_created(self):
        rpt = os.path.join(self.tmpdir, "report.xlsx")
        result = generate_mapping_report(self.wb_path, output_path=rpt)
        self.assertTrue(os.path.exists(result))

    def test_report_has_sheet(self):
        rpt = os.path.join(self.tmpdir, "report2.xlsx")
        generate_mapping_report(self.wb_path, output_path=rpt)
        wb = load_workbook(rpt)
        self.assertIn("Data", wb.sheetnames)
        wb.close()

    def test_report_contains_sections(self):
        rpt = os.path.join(self.tmpdir, "report3.xlsx")
        generate_mapping_report(self.wb_path, output_path=rpt)
        wb = load_workbook(rpt)
        ws = wb["Data"]
        values = [ws.cell(row=r, column=1).value for r in range(1, 30)]
        self.assertIn("Inputs", values)
        self.assertIn("Calculations", values)
        self.assertIn("Outputs", values)
        wb.close()


class TestGenerateMappingReportVerticalDrag(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmpdir = tempfile.mkdtemp()
        cls.wb_path = os.path.join(cls.tmpdir, "vertical.xlsx")
        _create_vertical_drag_workbook(cls.wb_path)

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmpdir)

    def test_dragged_formulas_collapsed(self):
        rpt = os.path.join(self.tmpdir, "report.xlsx")
        generate_mapping_report(self.wb_path, output_path=rpt)
        wb = load_workbook(rpt)
        ws = wb["Sales"]
        # Collect all cell values in columns A-B
        all_text = []
        for r in range(1, 40):
            for c in range(1, 4):
                v = ws.cell(row=r, column=c).value
                if v is not None:
                    all_text.append(str(v))
        joined = " ".join(all_text)
        # The 5 dragged C2:C6 formulas should be collapsed (mention "5 cells")
        self.assertIn("5 cells", joined)
        self.assertIn("vertical", joined)
        wb.close()

    def test_sum_is_output(self):
        rpt = os.path.join(self.tmpdir, "report_out.xlsx")
        generate_mapping_report(self.wb_path, output_path=rpt)
        wb = load_workbook(rpt)
        ws = wb["Sales"]
        # Find the Outputs section and check SUM is there
        in_outputs = False
        found_sum = False
        for r in range(1, 40):
            v = ws.cell(row=r, column=1).value
            if v == "Outputs":
                in_outputs = True
                continue
            if in_outputs:
                v2 = ws.cell(row=r, column=2).value
                if v2 and "SUM" in str(v2):
                    found_sum = True
                    break
        self.assertTrue(found_sum, "SUM formula should appear in Outputs section")
        wb.close()


class TestGenerateMappingReportMultiSheet(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmpdir = tempfile.mkdtemp()
        cls.wb_path = os.path.join(cls.tmpdir, "multi.xlsx")
        _create_multi_sheet_workbook(cls.wb_path)

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmpdir)

    def test_all_sheets_in_report(self):
        rpt = os.path.join(self.tmpdir, "report.xlsx")
        generate_mapping_report(self.wb_path, output_path=rpt)
        wb = load_workbook(rpt)
        self.assertIn("Input", wb.sheetnames)
        self.assertIn("Calc", wb.sheetnames)
        wb.close()

    def test_subset_sheets(self):
        rpt = os.path.join(self.tmpdir, "report_sub.xlsx")
        generate_mapping_report(self.wb_path, sheet_names=["Calc"],
                                output_path=rpt)
        wb = load_workbook(rpt)
        self.assertIn("Calc", wb.sheetnames)
        self.assertNotIn("Input", wb.sheetnames)
        wb.close()

    def test_cross_sheet_calc_vs_output(self):
        rpt = os.path.join(self.tmpdir, "report_cs.xlsx")
        generate_mapping_report(self.wb_path, output_path=rpt)
        wb = load_workbook(rpt)
        ws = wb["Calc"]
        all_text = []
        for r in range(1, 30):
            for c in range(1, 4):
                v = ws.cell(row=r, column=c).value
                if v is not None:
                    all_text.append(str(v))
        joined = " ".join(all_text)
        # A2 on Calc sheet references Input!A2 and is used by B2 → calculation
        self.assertIn("A2", joined)
        # B2 is not referenced → output
        self.assertIn("B2", joined)
        wb.close()


class TestGenerateMappingReportHorizontalDrag(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmpdir = tempfile.mkdtemp()
        cls.wb_path = os.path.join(cls.tmpdir, "horiz.xlsx")
        _create_horizontal_drag_workbook(cls.wb_path)

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmpdir)

    def test_horizontal_group_detected(self):
        rpt = os.path.join(self.tmpdir, "report.xlsx")
        generate_mapping_report(self.wb_path, output_path=rpt)
        wb = load_workbook(rpt)
        ws = wb["Forecast"]
        all_text = []
        for r in range(1, 30):
            for c in range(1, 4):
                v = ws.cell(row=r, column=c).value
                if v is not None:
                    all_text.append(str(v))
        joined = " ".join(all_text)
        self.assertIn("5 cells", joined)
        self.assertIn("horizontal", joined)
        wb.close()


class TestGenerateMappingReportEdgeCases(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmpdir = tempfile.mkdtemp()

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmpdir)

    def test_no_formulas(self):
        path = os.path.join(self.tmpdir, "nof.xlsx")
        _create_no_formulas_workbook(path)
        rpt = os.path.join(self.tmpdir, "rpt_nof.xlsx")
        generate_mapping_report(path, output_path=rpt)
        wb = load_workbook(rpt)
        ws = wb["Static"]
        vals = [ws.cell(row=r, column=1).value for r in range(1, 20)]
        self.assertIn("Inputs", vals)
        self.assertIn("Calculations", vals)
        self.assertIn("Outputs", vals)
        # Should show (none) for Calculations and Outputs
        self.assertIn("(none)", vals)
        wb.close()

    def test_all_formulas(self):
        path = os.path.join(self.tmpdir, "allf.xlsx")
        _create_all_formulas_workbook(path)
        rpt = os.path.join(self.tmpdir, "rpt_allf.xlsx")
        generate_mapping_report(path, output_path=rpt)
        wb = load_workbook(rpt)
        ws = wb["Formulas"]
        vals = [ws.cell(row=r, column=1).value for r in range(1, 20)]
        self.assertIn("Inputs", vals)
        # Inputs should show (none) since there are no hardcoded values
        self.assertIn("(none)", vals)
        wb.close()

    def test_nonexistent_sheet_ignored(self):
        path = os.path.join(self.tmpdir, "simple2.xlsx")
        _create_simple_workbook(path)
        rpt = os.path.join(self.tmpdir, "rpt_ne.xlsx")
        generate_mapping_report(path, sheet_names=["NoSuchSheet"],
                                output_path=rpt)
        wb = load_workbook(rpt)
        # No sheet named NoSuchSheet should be created
        self.assertNotIn("NoSuchSheet", wb.sheetnames)
        # The file should still be valid (openpyxl keeps default Sheet)
        self.assertTrue(len(wb.sheetnames) >= 1)
        wb.close()


class TestGenerateMappingReportSampleWorkbook(unittest.TestCase):
    """End-to-end with the shared sample workbook used by vectorized tests."""

    @classmethod
    def setUpClass(cls):
        cls.tmpdir = tempfile.mkdtemp()
        cls.sample_path = os.path.join(cls.tmpdir, "sample.xlsx")
        create_sample_workbook(cls.sample_path)

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmpdir)

    def test_all_sheets_present(self):
        rpt = os.path.join(self.tmpdir, "rpt_sample.xlsx")
        generate_mapping_report(self.sample_path, output_path=rpt)
        wb = load_workbook(rpt)
        self.assertIn("Inputs", wb.sheetnames)
        self.assertIn("Summary", wb.sheetnames)
        self.assertIn("Rates", wb.sheetnames)
        wb.close()

    def test_inputs_sheet_has_inputs(self):
        rpt = os.path.join(self.tmpdir, "rpt_sample2.xlsx")
        generate_mapping_report(self.sample_path, output_path=rpt)
        wb = load_workbook(rpt)
        ws = wb["Inputs"]
        vals = []
        for r in range(1, 50):
            v = ws.cell(row=r, column=2).value
            if v is not None:
                vals.append(v)
        # Should have input values like 10.5, 25, 7.25 etc.
        self.assertTrue(any(isinstance(v, (int, float)) for v in vals))
        wb.close()

    def test_default_output_path(self):
        """When no output_path given, report goes to ./output/."""
        rpt = generate_mapping_report(self.sample_path)
        self.assertTrue(os.path.exists(rpt))
        self.assertIn("mapping_report.xlsx", rpt)
        # Clean up
        os.remove(rpt)


if __name__ == "__main__":
    unittest.main()
