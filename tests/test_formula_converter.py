"""Tests for the formula_converter module."""

import unittest
from formula_converter import (
    FormulaConverter,
    col_letter_to_index,
    index_to_col_letter,
    cell_to_var_name,
    range_to_var_name,
)


class TestColConversion(unittest.TestCase):
    def test_single_letter(self):
        self.assertEqual(col_letter_to_index("A"), 1)
        self.assertEqual(col_letter_to_index("Z"), 26)

    def test_double_letter(self):
        self.assertEqual(col_letter_to_index("AA"), 27)
        self.assertEqual(col_letter_to_index("AZ"), 52)

    def test_index_to_col(self):
        self.assertEqual(index_to_col_letter(1), "A")
        self.assertEqual(index_to_col_letter(26), "Z")
        self.assertEqual(index_to_col_letter(27), "AA")

    def test_roundtrip(self):
        for i in range(1, 100):
            self.assertEqual(col_letter_to_index(index_to_col_letter(i)), i)


class TestCellVarName(unittest.TestCase):
    def test_simple(self):
        self.assertEqual(cell_to_var_name("Sheet1", "A", 1), "s_Sheet1_A1")

    def test_spaces(self):
        self.assertEqual(cell_to_var_name("My Sheet", "B", 2), "s_My_Sheet_B2")


class TestFormulaConverter(unittest.TestCase):
    def setUp(self):
        self.converter = FormulaConverter("Sheet1")

    def test_simple_addition(self):
        result = self.converter.convert("=A1+B1")
        self.assertIn("s_Sheet1_A1", result)
        self.assertIn("s_Sheet1_B1", result)
        self.assertIn("+", result)

    def test_multiplication(self):
        result = self.converter.convert("=B2*C2")
        self.assertIn("s_Sheet1_B2", result)
        self.assertIn("s_Sheet1_C2", result)
        self.assertIn("*", result)

    def test_sum_function(self):
        result = self.converter.convert("=SUM(A1:A10)")
        self.assertIn("xl_sum", result)

    def test_average_function(self):
        result = self.converter.convert("=AVERAGE(B1:B5)")
        self.assertIn("xl_average", result)

    def test_if_function(self):
        result = self.converter.convert('=IF(A1>10,"Yes","No")')
        self.assertIn("xl_if", result)
        self.assertIn('"Yes"', result)
        self.assertIn('"No"', result)

    def test_cross_sheet_reference(self):
        result = self.converter.convert("=Sheet2!A1+B1")
        self.assertIn("s_Sheet2_A1", result)
        self.assertIn("s_Sheet1_B1", result)
        self.assertEqual(len(self.converter.referenced_cells), 2)

    def test_quoted_sheet_reference(self):
        result = self.converter.convert("='My Sheet'!A1")
        self.assertIn("s_My_Sheet_A1", result)
        refs = self.converter.referenced_cells
        self.assertIn(("My Sheet", "A", 1), refs)

    def test_dollar_sign_references(self):
        result = self.converter.convert("=$A$1+$B2")
        self.assertIn("s_Sheet1_A1", result)
        self.assertIn("s_Sheet1_B2", result)

    def test_range_reference(self):
        result = self.converter.convert("=SUM(A1:A5)")
        self.assertTrue(len(self.converter.referenced_ranges) > 0)

    def test_power_operator(self):
        result = self.converter.convert("=A1^2")
        self.assertIn("**", result)

    def test_comparison_operators(self):
        result = self.converter.convert("=IF(A1>=10,1,0)")
        self.assertIn(">=", result)

    def test_not_equal(self):
        result = self.converter.convert("=IF(A1<>0,1,0)")
        self.assertIn("!=", result)

    def test_number_literal(self):
        result = self.converter.convert("=A1+100")
        self.assertIn("100", result)

    def test_string_literal(self):
        result = self.converter.convert('=IF(A1>0,"Positive","Negative")')
        self.assertIn('"Positive"', result)
        self.assertIn('"Negative"', result)

    def test_nested_functions(self):
        result = self.converter.convert("=SUM(IF(A1>0,B1,0))")
        self.assertIn("xl_sum", result)
        self.assertIn("xl_if", result)

    def test_cross_sheet_range(self):
        result = self.converter.convert("=SUM(Sheet2!A1:A5)")
        self.assertTrue(
            any(r[0] == "Sheet2" for r in self.converter.referenced_ranges)
        )

    def test_max_function(self):
        result = self.converter.convert("=MAX(B2:B4)")
        self.assertIn("xl_max", result)

    def test_referenced_cells_tracked(self):
        self.converter.convert("=A1+B2+Sheet2!C3")
        refs = self.converter.referenced_cells
        self.assertIn(("Sheet1", "A", 1), refs)
        self.assertIn(("Sheet1", "B", 2), refs)
        self.assertIn(("Sheet2", "C", 3), refs)


class TestFormulaConverterWithTables(unittest.TestCase):
    def setUp(self):
        self.tables = {
            "SalesTable": {
                "sheet": "Data",
                "ref": "A1:C10",
                "columns": ["Product", "Price", "Qty"],
                "header_row": 1,
                "data_start_row": 2,
                "data_end_row": 10,
                "col_start": "A",
                "col_end": "C",
            }
        }
        self.converter = FormulaConverter("Data", self.tables)

    def test_table_reference(self):
        result = self.converter.convert("=SUM(SalesTable[Price])")
        self.assertIn("xl_sum", result)
        self.assertIn("tbl_SalesTable_Price", result)
        self.assertIn(("SalesTable", "Price"), self.converter.referenced_tables)


if __name__ == "__main__":
    unittest.main()
