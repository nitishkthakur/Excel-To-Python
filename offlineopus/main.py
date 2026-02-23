"""
Excel-To-Python: Main Entry Point
==================================
Converts an Excel workbook with formulas into a standalone Python calculator.

Usage:
    python main.py [config_file]
    python main.py --excel <path_to_excel> [--output <output_dir>]

The tool generates:
1. calculator.py     - Standalone Python script with all calculations
2. input_template.xlsx - Excel template for user to provide input values
3. excel_functions.py  - Helper library of Excel function implementations
"""

import os
import sys
import shutil
import logging
import argparse
import yaml

# Add parent directory to path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from excel_to_python.excel_parser import parse_workbook
from excel_to_python.code_generator import generate_calculator
from excel_to_python.input_template import generate_input_template


def setup_logging(level_str: str = "INFO"):
    """Configure logging."""
    level = getattr(logging, level_str.upper(), logging.INFO)
    logging.basicConfig(
        level=level,
        format='%(asctime)s [%(levelname)s] %(name)s: %(message)s',
        datefmt='%H:%M:%S'
    )


def load_config(config_path: str) -> dict:
    """Load configuration from YAML file."""
    with open(config_path, 'r') as f:
        return yaml.safe_load(f)


def main():
    parser = argparse.ArgumentParser(
        description='Convert Excel workbook to Python calculator'
    )
    parser.add_argument(
        'config', nargs='?', default='config.yaml',
        help='Path to config YAML file (default: config.yaml)'
    )
    parser.add_argument(
        '--excel', '-e', type=str, default=None,
        help='Path to source Excel file (overrides config)'
    )
    parser.add_argument(
        '--output', '-o', type=str, default=None,
        help='Output directory (overrides config)'
    )
    parser.add_argument(
        '--keep-unreferenced', action='store_true',
        help='Keep unreferenced hardcoded values'
    )
    parser.add_argument(
        '--log-level', type=str, default=None,
        help='Logging level: DEBUG, INFO, WARNING, ERROR'
    )

    args = parser.parse_args()

    # Load config
    config = {}
    if os.path.exists(args.config):
        config = load_config(args.config)

    # Apply command-line overrides
    excel_path = args.excel or config.get('source_excel', '')
    output_dir = args.output or config.get('output_dir', 'generated_output')
    delete_unreferenced = not args.keep_unreferenced and config.get(
        'delete_unreferenced_hardcoded_values', True)
    skip_sheets = config.get('skip_sheets', [])
    log_level = args.log_level or config.get('log_level', 'INFO')

    setup_logging(log_level)
    logger = logging.getLogger(__name__)

    if not excel_path:
        logger.error("No Excel file specified. Use --excel or set source_excel in config.yaml")
        sys.exit(1)

    if not os.path.exists(excel_path):
        logger.error(f"Excel file not found: {excel_path}")
        sys.exit(1)

    logger.info(f"Excel-To-Python Converter")
    logger.info(f"========================")
    logger.info(f"Source Excel: {excel_path}")
    logger.info(f"Output directory: {output_dir}")
    logger.info(f"Delete unreferenced hardcoded values: {delete_unreferenced}")
    logger.info(f"Skip sheets: {skip_sheets}")
    logger.info(f"")

    # Step 1: Parse the Excel workbook
    logger.info("Step 1: Parsing Excel workbook...")
    workbook_info = parse_workbook(excel_path, skip_sheets)
    logger.info(f"  Parsed {len(workbook_info.sheet_names)} sheets, "
                f"{len(workbook_info.all_cells)} total cells")
    logger.info("")

    # Step 2: Generate the calculator Python script
    logger.info("Step 2: Generating calculator script...")
    calc_path = generate_calculator(workbook_info, output_dir, delete_unreferenced)
    logger.info("")

    # Step 3: Generate the input template
    logger.info("Step 3: Generating input template...")
    template_path = generate_input_template(workbook_info, output_dir, delete_unreferenced)
    logger.info("")

    # Step 4: Copy the Excel functions helper to output directory
    logger.info("Step 4: Copying Excel functions helper...")
    src_functions = os.path.join(os.path.dirname(__file__), 'excel_to_python', 'excel_functions.py')
    dst_functions = os.path.join(output_dir, 'excel_functions.py')
    shutil.copy2(src_functions, dst_functions)
    logger.info(f"  Copied to: {dst_functions}")
    logger.info("")

    # Summary
    logger.info("=" * 50)
    logger.info("Generation Complete!")
    logger.info("=" * 50)
    logger.info(f"")
    logger.info(f"Generated files in '{output_dir}/':")
    logger.info(f"  1. calculator.py       - The calculation engine")
    logger.info(f"  2. input_template.xlsx  - Fill in your input values here")
    logger.info(f"  3. excel_functions.py   - Excel function implementations")
    logger.info(f"")
    logger.info(f"Usage:")
    logger.info(f"  1. Edit input_template.xlsx with your desired input values")
    logger.info(f"     (Yellow cells are inputs - modify as needed)")
    logger.info(f"  2. Run: python {os.path.join(output_dir, 'calculator.py')} "
                f"{os.path.join(output_dir, 'input_template.xlsx')} output.xlsx")
    logger.info(f"  3. Open output.xlsx to see the results")


if __name__ == '__main__':
    main()
