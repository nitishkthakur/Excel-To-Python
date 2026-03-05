#!/usr/bin/env python3
"""
Excel-to-Python Conversion Pipeline
Main entry point for running the pipeline.
"""

import argparse
import sys
from pathlib import Path

# Setup path
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from excel_pipeline.utils.config import config
from excel_pipeline.utils.logging_setup import setup_logging
from excel_pipeline.layer1.parser import generate_mapping_report


def main():
    """Main CLI entry point."""
    parser = argparse.ArgumentParser(
        description="Excel-to-Python Conversion Pipeline",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Generate mapping report (Layer 1)
  python run_pipeline.py --layer 1 --input model.xlsx --output mapping_report.xlsx

  # Run complete pipeline
  python run_pipeline.py --input model.xlsx --output-dir results/
        """
    )

    parser.add_argument('--input', required=True, help='Input Excel file')
    parser.add_argument('--output', help='Output file path (for single layer)')
    parser.add_argument('--output-dir', help='Output directory (for complete pipeline)')
    parser.add_argument('--layer', type=int, choices=[1, 2, 3],
                       help='Run specific layer only (1=mapping, 2=inputs, 3=codegen)')
    parser.add_argument('--config', default='config.yaml', help='Config file path')
    parser.add_argument('--log-level', default='INFO',
                       choices=['DEBUG', 'INFO', 'WARNING', 'ERROR'],
                       help='Logging level')

    args = parser.parse_args()

    # Load configuration
    try:
        config.load(args.config)
    except FileNotFoundError:
        print(f"Warning: Config file not found: {args.config}, using defaults")

    # Setup logging
    logger = setup_logging(
        level=args.log_level,
        log_file=config.log_file
    )

    logger.info("=" * 80)
    logger.info("Excel-to-Python Conversion Pipeline")
    logger.info("=" * 80)

    try:
        # Validate input
        input_path = Path(args.input)
        if not input_path.exists():
            logger.error(f"Input file not found: {args.input}")
            return 1

        # Determine output paths
        if args.layer == 1:
            # Layer 1: Generate mapping report
            output_path = args.output or "mapping_report.xlsx"
            logger.info(f"Running Layer 1: Mapping Report Generation")
            generate_mapping_report(str(input_path), output_path)

        elif args.layer == 2:
            logger.error("Layer 2 not yet implemented")
            return 1

        elif args.layer == 3:
            logger.error("Layer 3 not yet implemented")
            return 1

        else:
            # Full pipeline (not yet implemented)
            logger.error("Full pipeline not yet implemented. Use --layer 1 for now.")
            return 1

        logger.info("\n" + "=" * 80)
        logger.info("Pipeline completed successfully!")
        logger.info("=" * 80)
        return 0

    except Exception as e:
        logger.error(f"Pipeline failed: {e}", exc_info=True)
        return 1


if __name__ == "__main__":
    sys.exit(main())
