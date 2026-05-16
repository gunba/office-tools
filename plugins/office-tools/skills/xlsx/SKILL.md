---
description: Read, search, or edit Microsoft Excel (.xlsx) workbooks. Read-only operations use openpyxl and never save. Closed-workbook edits default to the direct OOXML engine, which rewrites only selected package parts and preserves workbook metadata such as conditional formatting and data validation. Use Excel COM for already-open workbooks, validation/recalculation, sheet copying, and formula-aware structural insertions. NEVER use openpyxl or another pure-Python workbook library to save/write a workbook directly.
---

# Excel Workbook Tool (xlsx)

The tool is the Python script at `${CLAUDE_PLUGIN_ROOT}/tools/xlsx/xlsx_tool.py`, invoked through the plugin's bundled Python at `${CLAUDE_PLUGIN_ROOT}/python-base/python.exe`. The tool's third-party dependencies (openpyxl, xlwings) install lazily into the plugin's persistent data dir on the first call — there is nothing to install manually.

Read-only commands use **openpyxl** because they do not save or rewrite the workbook. Write commands default to the tool's **direct OOXML** engine for closed workbooks: it edits only the required XML package parts and leaves unrelated workbook XML untouched. Excel COM remains available for live workbooks, calculation validation, sheet copying, and structural operations where Excel should adjust formulas/references.

## Usage

In every example below, `<PY>` stands for `${CLAUDE_PLUGIN_ROOT}/python-base/python.exe` and `<TOOL>` stands for `${CLAUDE_PLUGIN_ROOT}/tools/xlsx/xlsx_tool.py`. On Codex the same paths resolve via this skill's directory: from the skill root, the tool is at `../../tools/xlsx/xlsx_tool.py` and the interpreter at `../../python-base/python.exe`.

```bash
# Read a compact markdown table without opening Excel.
<PY> <TOOL> read <file> [--sheet "Sheet1"] [--range "A1:D10"]

# Preserve fully empty display rows/columns when that context matters.
<PY> <TOOL> read <file> --sheet "Sheet1" --show-empty-rows --show-empty-cols

# Read via Excel COM when unsaved live workbook state or Excel-calculated values are required.
<PY> <TOOL> read <file> --backend com [--sheet "Sheet1"] [--range "A1:D10"]

# List/search without opening Excel.
<PY> <TOOL> list-sheets <file>
<PY> <TOOL> search <file> --query "keyword" [--sheet "Sheet1"]

# Direct closed-workbook edits.
<PY> <TOOL> edit <file> --sheet "Sheet1" --cell A1 --value "new value"
<PY> <TOOL> edit <file> --sheet "Sheet1" --range A1:D10 --clear

# Batch edits. Prefer this for large edits or PowerShell sessions where inline JSON quoting is fragile.
<PY> <TOOL> edit <file> --edits-file edits.json

# Force a type for a single-cell direct edit.
<PY> <TOOL> edit <file> --sheet "Sheet1" --cell A1 --value "2024-12-31" --value-type date

# Direct sheet ordering operations for closed workbooks.
<PY> <TOOL> add-sheet <file> "Support" --after "Main"
<PY> <TOOL> move-sheet <file> "Support" --before "Main"

# Use Excel when the workbook is open, protected, or needs Excel-native behavior.
<PY> <TOOL> edit <file> --engine com --sheet "Sheet1" --cell A1 --value "new value"
<PY> <TOOL> insert <file> --sheet "Sheet1" --axis columns --range D:D

# Validate formulas/cached values with Excel after direct formula edits.
<PY> <TOOL> validate <file> --full-calc --save --check-errors

# Copy whole sheets from another workbook through Excel COM.
<PY> <TOOL> copy-sheets <file> --source <source.xlsx> --sheets "Support Tab" "Old Name=New Name" --replace
```

## Notes

- `list-sheets`, `read`, and `search` use openpyxl by default and do not save the workbook.
- `read` omits fully empty rows and columns by default to preserve context budget. Use `--show-empty-rows` or `--show-empty-cols` when positional blanks matter.
- The direct edit engine refuses to write if the workbook is already open in Excel. Use `--engine com` for live workbooks.
- The direct edit engine refuses protected sheets unless `--allow-protected` is passed. Prefer COM for protected workbooks because it can preflight locked cells through Excel.
- Formula edits are marked for recalculation on next Excel open. Run `validate --full-calc --save` when cached formula results are needed for downstream reads.
- Use `--edits-file` for batch edits. JSON edits may use `"cell"` or `"range"`, `"value"` for scalars, `"values"` for row/column/matrix writes, and `null` to clear.
- Use `copy-sheets` when whole support sheets must be moved between workbooks while preserving formatting.
- Use `insert` for row/column insertions so Excel adjusts formulas, names, conditional formatting, and validation ranges.
- **NEVER use openpyxl or any pure-Python workbook library to save/write a workbook directly.** All workbook writes must go through `xlsx_tool.py`.
- COM operations (`--engine com`, `validate`, `insert`, `copy-sheets`) require Microsoft Excel to be installed on the machine. The OOXML engine and openpyxl-based reads work without Excel.

## Working files and ephemera

Do not write helper scripts, JSON specs, or intermediate artefacts into `${CLAUDE_PLUGIN_ROOT}` or its subdirectories — the plugin install directory is treated as read-only. For batch edit JSON or other transient data, use Python's `tempfile` module or write into the user's current working directory with an explicit, short-lived filename and delete it when done.
