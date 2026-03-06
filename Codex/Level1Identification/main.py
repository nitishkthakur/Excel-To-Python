#!/usr/bin/env python3
from __future__ import annotations

from pathlib import Path
import sys

from level1_hardcode import transform_workbook


def build_output_path(input_path: Path, target_sheet: str) -> Path:
    safe_sheet = "".join(ch if ch.isalnum() or ch in ("-", "_") else "_" for ch in target_sheet).strip("_")
    if not safe_sheet:
        safe_sheet = "target"
    return input_path.with_name(f"{input_path.stem}.{safe_sheet}.level1_hardcoded{input_path.suffix}")


def main() -> int:
    filename = input("Enter Excel filename (with path or in this folder): ").strip()
    if not filename:
        print("No filename provided.")
        return 1

    input_path = Path(filename).expanduser()
    if not input_path.is_absolute():
        input_path = Path.cwd() / input_path
    input_path = input_path.resolve()

    if not input_path.exists():
        print(f"File not found: {input_path}")
        return 1

    if input_path.suffix.lower() != ".xlsx":
        print("Only .xlsx files are supported.")
        return 1

    sheet_name = input("Enter target sheet name: ").strip()
    if not sheet_name:
        print("No sheet name provided.")
        return 1

    output_path = build_output_path(input_path, sheet_name)

    try:
        result = transform_workbook(
            input_xlsx=input_path,
            output_xlsx=output_path,
            target_sheet=sheet_name,
            fail_on_target_errors=True,
        )
    except Exception as exc:
        print(f"Failed: {exc}")
        return 1

    print("Success")
    print(f"Input:  {input_path}")
    print(f"Output: {output_path}")
    print(f"Target: {result['target_sheet']}")
    print(f"Level1 predecessors: {', '.join(result['predecessors']) or '(none)'}")
    print(f"Sheets kept: {', '.join(result['kept_sheets'])}")
    print(f"Sheets removed: {', '.join(result['removed_sheets']) or '(none)'}")
    print(f"Formulas hardcoded: {result['formulas_hardcoded']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
