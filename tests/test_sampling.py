"""
Tests for the three sampling strategies (smart_random, full, column_n)
and for the MCP-server dispatch logic.
"""

import os
import sys
import json
import pytest

# Make the mcp_server package and tests directory importable
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "mcp_server"))
sys.path.insert(0, os.path.dirname(__file__))

from create_sample_workbook import create_sample_workbook

# Import samplers
from excel_reader_smart_sampler import (
    extract_sheet_data as smart_extract,
    extract_formulas,
    workbook_summary,
    sheet_names,
    detect_regions,
    open_workbook,
    DEFAULT_SAMPLE_ROWS,
)
from fetcher_full import extract_sheet_data as full_extract
from fetcher_column_n import extract_sheet_data as column_n_extract


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture(scope="module")
def sample_wb(tmp_path_factory):
    """Create the sample workbook once per test module."""
    path = str(tmp_path_factory.mktemp("data") / "sample.xlsx")
    create_sample_workbook(path)
    return path


# ---------------------------------------------------------------------------
# Helpers for assertions
# ---------------------------------------------------------------------------

def _result_shape_ok(data: dict):
    """Assert that *data* has the standard result dict shape."""
    assert "sheet_name" in data
    assert "regions" in data
    assert isinstance(data["regions"], list)
    assert "sampled" in data
    assert "total_rows" in data
    assert "sampled_rows" in data
    for reg in data["regions"]:
        assert "headers" in reg
        assert "rows" in reg
        assert "formulas" in reg
        assert "min_row" in reg
        assert "max_row" in reg
        assert "min_col" in reg
        assert "max_col" in reg


# ---------------------------------------------------------------------------
# smart_random strategy
# ---------------------------------------------------------------------------

class TestSmartRandom:

    def test_basic_extraction(self, sample_wb):
        data = smart_extract(sample_wb, "Inputs")
        _result_shape_ok(data)
        assert data["sheet_name"] == "Inputs"

    def test_sampling_flag_small_sheet(self, sample_wb):
        """The sample workbook is small — sampling should NOT be applied."""
        data = smart_extract(sample_wb, "Inputs")
        # With default max rows (100) and a tiny test sheet, sampled should be False
        assert data["sampled"] is False

    def test_formulas_present(self, sample_wb):
        data = smart_extract(sample_wb, "Inputs")
        all_formulas = []
        for reg in data["regions"]:
            all_formulas.extend(reg["formulas"])
        assert len(all_formulas) > 0

    def test_extract_formulas_function(self, sample_wb):
        formulas = extract_formulas(sample_wb, "Inputs")
        assert isinstance(formulas, list)
        assert len(formulas) > 0
        for f in formulas:
            assert "address" in f
            assert "formula" in f

    def test_workbook_summary(self, sample_wb):
        summary = workbook_summary(sample_wb)
        assert "sheets" in summary
        names = [s["name"] for s in summary["sheets"]]
        assert "Inputs" in names
        assert "Summary" in names
        assert "Rates" in names

    def test_sheet_names(self, sample_wb):
        names = sheet_names(sample_wb)
        assert "Inputs" in names
        assert "Summary" in names


# ---------------------------------------------------------------------------
# full strategy
# ---------------------------------------------------------------------------

class TestFull:

    def test_basic_extraction(self, sample_wb):
        data = full_extract(sample_wb, "Inputs")
        _result_shape_ok(data)
        assert data["sheet_name"] == "Inputs"

    def test_all_rows_loaded(self, sample_wb):
        """Full mode with nrows=None should load every row."""
        data = full_extract(sample_wb, "Inputs")
        assert data["sampled"] is False
        assert data["sampled_rows"] == data["total_rows"]

    def test_nrows_cap(self, sample_wb):
        data = full_extract(sample_wb, "Inputs", nrows=3)
        for reg in data["regions"]:
            row_span = reg["max_row"] - reg["min_row"] + 1
            assert row_span <= 3

    def test_ncols_cap(self, sample_wb):
        data_full = full_extract(sample_wb, "Inputs")
        data_capped = full_extract(sample_wb, "Inputs", ncols=2)
        for reg in data_capped["regions"]:
            col_span = reg["max_col"] - reg["min_col"] + 1
            assert col_span <= 2

    def test_formulas_present(self, sample_wb):
        data = full_extract(sample_wb, "Inputs")
        all_formulas = []
        for reg in data["regions"]:
            all_formulas.extend(reg["formulas"])
        assert len(all_formulas) > 0

    def test_all_sheets(self, sample_wb):
        for name in sheet_names(sample_wb):
            data = full_extract(sample_wb, name)
            _result_shape_ok(data)


# ---------------------------------------------------------------------------
# column_n strategy
# ---------------------------------------------------------------------------

class TestColumnN:

    def test_basic_extraction(self, sample_wb):
        data = column_n_extract(sample_wb, "Inputs")
        _result_shape_ok(data)
        assert data["sheet_name"] == "Inputs"

    def test_strip_width(self, sample_wb):
        """With num_columns=2 the strip should have at most 3 columns
        (1 label + 2 data columns)."""
        data = column_n_extract(sample_wb, "Inputs", num_columns=2)
        for reg in data["regions"]:
            col_span = reg["max_col"] - reg["min_col"] + 1
            assert col_span <= 3  # label + 2

    def test_default_strip_width(self, sample_wb):
        """Default num_columns is 10 — strip should not exceed 11 cols."""
        data = column_n_extract(sample_wb, "Rates")
        for reg in data["regions"]:
            col_span = reg["max_col"] - reg["min_col"] + 1
            assert col_span <= 11

    def test_all_rows_included(self, sample_wb):
        """column_n loads ALL rows in the strip."""
        data = column_n_extract(sample_wb, "Inputs")
        assert data["sampled"] is False

    def test_formulas_present(self, sample_wb):
        data = column_n_extract(sample_wb, "Inputs", num_columns=10)
        all_formulas = []
        for reg in data["regions"]:
            all_formulas.extend(reg["formulas"])
        assert len(all_formulas) > 0

    def test_all_sheets(self, sample_wb):
        for name in sheet_names(sample_wb):
            data = column_n_extract(sample_wb, name)
            _result_shape_ok(data)


# ---------------------------------------------------------------------------
# Dispatch logic (mirrors server._dispatch_extract without MCP dependency)
# ---------------------------------------------------------------------------

class TestDispatch:

    def _dispatch(self, mode, path, sheet_name, **kwargs):
        """Replicates the dispatch logic from server.py."""
        mode = mode.lower().strip()
        valid = ("smart_random", "full", "column_n")
        if mode not in valid:
            raise ValueError(f"Invalid mode '{mode}'. Must be one of {valid}.")
        if mode == "smart_random":
            return smart_extract(path, sheet_name,
                                 max_sample_rows=kwargs.get("max_sample_rows",
                                                            DEFAULT_SAMPLE_ROWS))
        elif mode == "full":
            return full_extract(path, sheet_name,
                                nrows=kwargs.get("nrows"),
                                ncols=kwargs.get("ncols"))
        elif mode == "column_n":
            return column_n_extract(path, sheet_name,
                                    num_columns=kwargs.get("num_columns", 10))

    def test_smart_random_dispatch(self, sample_wb):
        data = self._dispatch("smart_random", sample_wb, "Inputs")
        _result_shape_ok(data)

    def test_full_dispatch(self, sample_wb):
        data = self._dispatch("full", sample_wb, "Inputs")
        _result_shape_ok(data)

    def test_column_n_dispatch(self, sample_wb):
        data = self._dispatch("column_n", sample_wb, "Inputs")
        _result_shape_ok(data)

    def test_invalid_mode_raises(self, sample_wb):
        with pytest.raises(ValueError, match="Invalid mode"):
            self._dispatch("nonexistent", sample_wb, "Inputs")

    def test_modes_produce_different_output(self, sample_wb):
        """The three modes should produce structurally valid but potentially
        different data (e.g. different column counts for column_n)."""
        d1 = self._dispatch("smart_random", sample_wb, "Inputs")
        d2 = self._dispatch("full", sample_wb, "Inputs")
        d3 = self._dispatch("column_n", sample_wb, "Inputs", num_columns=2)
        for d in (d1, d2, d3):
            _result_shape_ok(d)
