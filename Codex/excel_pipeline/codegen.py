from __future__ import annotations

from pathlib import Path


UNSTRUCTURED_TEMPLATE = """#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path

from excel_pipeline.runtime import run_unstructured_calculation

DEFAULT_MAPPING_REPORT = Path(r\"{mapping_report}\")
DEFAULT_INPUT = Path(r\"{default_input}\")
DEFAULT_OUTPUT = Path(r\"{default_output}\")


def main() -> None:
    parser = argparse.ArgumentParser(description=\"Rebuild output workbook from unstructured inputs\")
    parser.add_argument(\"--inputs\", type=Path, default=DEFAULT_INPUT)
    parser.add_argument(\"--mapping\", type=Path, default=DEFAULT_MAPPING_REPORT)
    parser.add_argument(\"--output\", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()

    run_unstructured_calculation(
        mapping_report_path=args.mapping,
        unstructured_input_path=args.inputs,
        output_path=args.output,
    )


if __name__ == \"__main__\":
    main()
"""


UNSTRUCTURED_VECTORIZED_TEMPLATE = """#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path

from excel_pipeline.vectorized_runtime import run_unstructured_calculation_vectorized

DEFAULT_MAPPING_REPORT = Path(r\"{mapping_report}\")
DEFAULT_INPUT = Path(r\"{default_input}\")
DEFAULT_OUTPUT = Path(r\"{default_output}\")


def main() -> None:
    parser = argparse.ArgumentParser(
        description=\"Rebuild output workbook from unstructured inputs using vectorized grouped calculations\"
    )
    parser.add_argument(\"--inputs\", type=Path, default=DEFAULT_INPUT)
    parser.add_argument(\"--mapping\", type=Path, default=DEFAULT_MAPPING_REPORT)
    parser.add_argument(\"--output\", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()

    run_unstructured_calculation_vectorized(
        mapping_report_path=args.mapping,
        unstructured_input_path=args.inputs,
        output_path=args.output,
    )


if __name__ == \"__main__\":
    main()
"""


STRUCTURED_TEMPLATE = """#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path

from excel_pipeline.runtime import run_structured_calculation

DEFAULT_MAPPING_REPORT = Path(r\"{mapping_report}\")
DEFAULT_INPUT = Path(r\"{default_input}\")
DEFAULT_OUTPUT = Path(r\"{default_output}\")


def main() -> None:
    parser = argparse.ArgumentParser(description=\"Rebuild output workbook from structured inputs\")
    parser.add_argument(\"--inputs\", type=Path, default=DEFAULT_INPUT)
    parser.add_argument(\"--mapping\", type=Path, default=DEFAULT_MAPPING_REPORT)
    parser.add_argument(\"--output\", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()

    run_structured_calculation(
        mapping_report_path=args.mapping,
        structured_input_path=args.inputs,
        output_path=args.output,
    )


if __name__ == \"__main__\":
    main()
"""


def generate_unstructured_calculate(
    mapping_report_path: Path,
    output_script_path: Path,
    default_input_path: Path,
    default_output_path: Path,
) -> Path:
    script = UNSTRUCTURED_TEMPLATE.format(
        mapping_report=str(mapping_report_path.resolve()),
        default_input=str(default_input_path.resolve()),
        default_output=str(default_output_path.resolve()),
    )
    output_script_path.parent.mkdir(parents=True, exist_ok=True)
    output_script_path.write_text(script, encoding="utf-8")
    return output_script_path


def generate_unstructured_vectorized_calculate(
    mapping_report_path: Path,
    output_script_path: Path,
    default_input_path: Path,
    default_output_path: Path,
) -> Path:
    script = UNSTRUCTURED_VECTORIZED_TEMPLATE.format(
        mapping_report=str(mapping_report_path.resolve()),
        default_input=str(default_input_path.resolve()),
        default_output=str(default_output_path.resolve()),
    )
    output_script_path.parent.mkdir(parents=True, exist_ok=True)
    output_script_path.write_text(script, encoding="utf-8")
    return output_script_path


def generate_structured_calculate(
    mapping_report_path: Path,
    output_script_path: Path,
    default_input_path: Path,
    default_output_path: Path,
) -> Path:
    script = STRUCTURED_TEMPLATE.format(
        mapping_report=str(mapping_report_path.resolve()),
        default_input=str(default_input_path.resolve()),
        default_output=str(default_output_path.resolve()),
    )
    output_script_path.parent.mkdir(parents=True, exist_ok=True)
    output_script_path.write_text(script, encoding="utf-8")
    return output_script_path
