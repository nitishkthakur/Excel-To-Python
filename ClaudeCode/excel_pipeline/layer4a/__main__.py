"""CLI entry point for Layer 4a."""

import sys
from excel_pipeline.layer4a.code_generator import generate_unstructured_code


def main():
    """Main CLI entry point."""
    if len(sys.argv) < 3:
        print("Layer 4a: Python Code Generator")
        print("")
        print("Usage:")
        print("  python -m excel_pipeline.layer4a <mapping_report.xlsx> <unstructured_inputs.xlsx> [output_script.py]")
        print("")
        print("Arguments:")
        print("  mapping_report.xlsx       Path to mapping report from Layer 1")
        print("  unstructured_inputs.xlsx  Path to unstructured inputs from Layer 2a")
        print("  output_script.py          Path to generated Python script (default: unstructured_calculate.py)")
        print("")
        print("Example:")
        print("  python -m excel_pipeline.layer4a \\")
        print("      output/indigo_mapping_v3.xlsx \\")
        print("      output/indigo_unstructured_inputs.xlsx \\")
        print("      generated_unstructured_calculate.py")
        print("")
        sys.exit(1)

    mapping_path = sys.argv[1]
    inputs_path = sys.argv[2]
    output_path = sys.argv[3] if len(sys.argv) > 3 else "unstructured_calculate.py"

    try:
        generate_unstructured_code(mapping_path, inputs_path, output_path)
    except Exception as e:
        print(f"\nError: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
