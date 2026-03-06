"""
Single-sheet pipeline with cross-sheet reference support.

How cross-sheet references are preserved
-----------------------------------------
Problem: if we strip all other sheets (old approach), formulas like
    ='Balance Sheet'!B5
silently resolve to 0, giving wrong answers.

Solution (this module):
1.  freeze_other_sheets() creates intermediate/frozen_workbook.xlsx where
    every sheet EXCEPT the target has its formula cells replaced with the
    cached (last-computed) values stored inside the xlsx file.
2.  Layer 1 is run on the frozen workbook.  Because the other sheets now
    contain only values (no formulas), their cells are all classified as
    "Input" in the mapping report.
3.  Layer 2a therefore includes those cells in unstructured_inputs.xlsx,
    alongside the target sheet's genuine Input cells.
4.  Layer 4a generates Python code only for the target sheet's
    Calculation/Output cells.  Cross-sheet references are emitted as
        c.get(('OtherSheet', 'A', 1), 0)
    which resolves correctly because step 3 loaded those values into c.
5.  The generated script is executed; output.xlsx is the fully populated
    target sheet.

Artifact layout  (output_dir/)
-------------------------------
  intermediate/
      frozen_workbook.xlsx   workbook with other sheets value-only (on disk for inspection)
  mapping_report.xlsx        full cell metadata from frozen workbook
  unstructured_inputs.xlsx   target-sheet inputs + other-sheet frozen values
  unstructured_calculate.py  generated Python calculation script
  output.xlsx                fully populated output for the target sheet
  pipeline.log               run log
"""

import sys
import subprocess
import time
from pathlib import Path

from excel_pipeline.utils.config import config
from excel_pipeline.utils.logging_setup import setup_logging, get_logger
from excel_pipeline.layer1.parser import generate_mapping_report
from excel_pipeline.layer2.unstructured_generator import generate_unstructured_inputs
from excel_pipeline.layer4a.code_generator import generate_unstructured_code
from excel_pipeline.onesheet.freezer import freeze_other_sheets

logger = get_logger(__name__)


def run(
    input_path: str,
    sheet_name: str,
    output_dir: Path,
    log_level: str = "INFO",
) -> None:
    """
    Run the single-sheet pipeline with cross-sheet reference support.

    Args:
        input_path:  Path to the original Excel file (.xlsx).
        sheet_name:  Name of the target sheet to generate code for.
        output_dir:  Directory where all artifacts are written.
        log_level:   Logging verbosity ('DEBUG', 'INFO', 'WARNING', 'ERROR').

    Raises:
        RuntimeError: If the generated calculation script fails to execute.
        ValueError:   If sheet_name is not found in the workbook.
        FileNotFoundError: If input_path does not exist.
    """
    output_dir = Path(output_dir)
    intermediate_dir = output_dir / "intermediate"
    intermediate_dir.mkdir(parents=True, exist_ok=True)

    frozen_path  = str(intermediate_dir / "frozen_workbook.xlsx")
    mapping_path = str(output_dir / "mapping_report.xlsx")
    inputs_path  = str(output_dir / "unstructured_inputs.xlsx")
    script_path  = str(output_dir / "unstructured_calculate.py")

    setup_logging(level=log_level, log_file=str(output_dir / "pipeline.log"))

    try:
        config.load("config.yaml")
    except FileNotFoundError:
        pass  # Use built-in defaults

    wall_start = time.time()

    # ── Step 1: freeze other sheets ──────────────────────────────────────────
    logger.info("=" * 70)
    logger.info(f"Single-sheet pipeline  |  sheet: '{sheet_name}'")
    logger.info("=" * 70)
    logger.info("\n[1/5] Freezing other sheets → intermediate/frozen_workbook.xlsx")
    t0 = time.time()
    stats = freeze_other_sheets(input_path, sheet_name, frozen_path)
    logger.info(
        f"      {time.time()-t0:.1f}s  "
        f"({stats['other_sheets']} sheets frozen, "
        f"{stats['frozen_cells']} cells hardcoded)"
    )

    # ── Step 2: Layer 1 — mapping report ─────────────────────────────────────
    logger.info("\n[2/5] Layer 1: Mapping report from frozen workbook...")
    t0 = time.time()
    generate_mapping_report(frozen_path, mapping_path)
    logger.info(f"      {time.time()-t0:.1f}s  →  {mapping_path}")

    # ── Step 3: Layer 2a — unstructured inputs ───────────────────────────────
    # Includes both target-sheet Input cells AND the frozen values from every
    # other sheet (which are classified as Input because they have no formulas).
    logger.info("\n[3/5] Layer 2a: Unstructured inputs (target + frozen other sheets)...")
    t0 = time.time()
    generate_unstructured_inputs(mapping_path, inputs_path)
    logger.info(f"      {time.time()-t0:.1f}s  →  {inputs_path}")

    # ── Step 4: Layer 4a — code generation ───────────────────────────────────
    # Only target-sheet cells have formulas → only their Calculation/Output
    # cells get code generated.  Cross-sheet refs emit c.get(('Sheet', col, row), 0).
    logger.info("\n[4/5] Layer 4a: Generating calculation script...")
    t0 = time.time()
    generate_unstructured_code(mapping_path, inputs_path, script_path)
    logger.info(f"      {time.time()-t0:.1f}s  →  {script_path}")

    # ── Step 5: execute generated script ─────────────────────────────────────
    # Run from output_dir so that the relative paths hardcoded in the script
    # ("unstructured_inputs.xlsx" and "output.xlsx") resolve correctly.
    logger.info("\n[5/5] Running generated script → output.xlsx...")
    t0 = time.time()
    result = subprocess.run(
        [sys.executable, "unstructured_calculate.py"],
        cwd=str(output_dir),
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        logger.error("Generated script failed:")
        logger.error(result.stderr)
        raise RuntimeError(
            f"unstructured_calculate.py exited with code {result.returncode}.\n"
            f"{result.stderr}"
        )
    logger.info(f"      {time.time()-t0:.1f}s  →  {output_dir}/output.xlsx")

    total = time.time() - wall_start
    logger.info("\n" + "=" * 70)
    logger.info(f"Done in {total:.1f}s  |  artifacts in: {output_dir}/")
    logger.info(
        "  intermediate/frozen_workbook.xlsx  other sheets value-only (inspect if needed)"
    )
    logger.info("  mapping_report.xlsx               cell metadata & classifications")
    logger.info("  unstructured_inputs.xlsx          inputs + frozen cross-sheet values")
    logger.info("  unstructured_calculate.py         generated Python calculation script")
    logger.info("  output.xlsx                       fully populated output")
    logger.info("=" * 70)
