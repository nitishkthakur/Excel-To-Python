# Excel Level 1 Calculation Hardcoding

## Objective

Transform an Excel workbook by identifying and hardcoding cached values in direct predecessor sheets while preserving formulas in the target sheet.

## Requirements

Given a user-specified sheet:

1. **Identify Level 1 Predecessors**: Find all sheets that directly feed into the target sheet through formulas.
2. **Hardcode Values in Predecessors**: Replace all formulas in Level 1 sheets with their cached values, removing formula logic.
3. **Prune Sheets**: Delete all sheets except the target sheet and Level 1 predecessor sheets.
4. **Preserve Target Formulas**: Keep the target sheet's formulas intact; they continue to reference Level 1 sheets.
5. **Maintain Appearance**: The output file must display identically to the original (values, formatting). Formula inspection reveals hardcoded predecessors and formula-driven target.
6. **Retain variable names defined in excel**: Anticipate the case when variables have been defined in excel and used in formulas. When hardcoding the values, make sure to retain the variable names in the hardcoded sheets. This is to avoid #Name? Errors. 

## Test

1. Run the script on provided Excel files.
2. Open the output file and verify:
   - Only the target sheet and its direct predecessors remain.
   - Predecessor sheets contain hardcoded values, no formulas.
   - The target sheet's formulas still reference the predecessor sheets.
3. Compare the output file to the original to confirm identical appearance and correct formula references.
4. The target sheet must not contain #Value or #Name? errors, indicating successful hardcoding of predecessors while preserving target formulas and retention of any defined variables.

## Constraints

- **Performance Critical**: Files can be very large; use efficient libraries and batch operations.
- **Avoid Cell-by-Cell Iteration**: Leverage vectorized methods where possible.
- **Deep Planning Required**: Design the solution architecture before implementation.
