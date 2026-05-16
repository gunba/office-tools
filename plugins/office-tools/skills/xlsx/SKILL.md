---
description: Create, read, search, or edit Microsoft Excel .xlsx workbooks using the Rust office-tools binary. Direct writes edit only targeted OOXML package parts and preserve unrelated XML such as conditional formatting, data validation, workbook metadata, and extension lists. Use Excel COM commands only when Excel-native behavior is required.
---

# Excel Workbook Tool

Use the Rust binary:

```bash
"${CLAUDE_PLUGIN_ROOT}/bin/office-tools.exe" xlsx <command>
```

For local development from the repo root, use:

```bash
cargo run -- xlsx <command>
```

## Common Commands

```bash
# Read/list/search without opening Excel or saving the workbook.
office-tools xlsx read <file.xlsx> [--sheet "Sheet1"] [--range A1:D20]
office-tools xlsx cells <file.xlsx> [--sheet "Sheet1"] [--range A1:D20] --json
office-tools xlsx list-sheets <file.xlsx> --json
office-tools xlsx search <file.xlsx> --query "keyword" [--sheet "Sheet1"]
office-tools xlsx inspect <file.xlsx> [--sheet "Sheet1"]
office-tools xlsx relationships <file.xlsx> [--sheet "Sheet1"] --json
office-tools xlsx formulas <file.xlsx> [--sheet "Sheet1"] --json
office-tools xlsx tables <file.xlsx> [--sheet "Sheet1"] --json
office-tools xlsx validations <file.xlsx> [--sheet "Sheet1"] --json
office-tools xlsx conditional-formatting <file.xlsx> [--sheet "Sheet1"] --json
office-tools xlsx hyperlinks <file.xlsx> [--sheet "Sheet1"] --json
office-tools xlsx comments <file.xlsx> [--sheet "Sheet1"] --json
office-tools xlsx defined-names <file.xlsx> [--name "TotalCell"] --json
office-tools xlsx merged-ranges <file.xlsx> [--sheet "Sheet1"] --json
office-tools xlsx auto-filters <file.xlsx> [--sheet "Sheet1"] --json
office-tools xlsx protections <file.xlsx> [--sheet "Sheet1"] --json

# Create a new workbook from structured rows and sparse cell edits.
office-tools xlsx create workbook.json --output workbook.xlsx

# Direct closed-workbook edits. Prefer --edits-file for scripts.
office-tools xlsx edit <file.xlsx> --sheet "Sheet1" --cell A1 --value "new value"
office-tools xlsx edit <file.xlsx> --sheet "Sheet1" --range A1:D10 --value "0"
office-tools xlsx edit <file.xlsx> --edits-file edits.json

# Direct sheet order operations.
office-tools xlsx add-sheet <file.xlsx> "Support" --after "Main"
office-tools xlsx move-sheet <file.xlsx> "Support" --before "Main"

# Excel-native operations through WinCOM on Windows.
office-tools xlsx rename-sheet <file.xlsx> --sheet "Support" --name "Analysis"
office-tools xlsx insert <file.xlsx> --sheet "Sheet1" --axis columns --range D:D
office-tools xlsx autofit <file.xlsx> --sheet "Sheet1" --axis columns
office-tools xlsx format <file.xlsx> --sheet "Sheet1" --range B:B --number-format "#,##0.00"
office-tools xlsx validate <file.xlsx> --full-calc --save --check-errors
office-tools xlsx copy-sheets <file.xlsx> --source source.xlsx --sheets "Old=New" --replace
```

## Batch Edit JSON

```json
[
  {"sheet": "Sheet1", "cell": "A1", "value": "text"},
  {"sheet": "Sheet1", "range": "B2:C3", "values": [[1, 2], [3, 4]]},
  {"sheet": "Sheet1", "cell": "D1", "value": "=SUM(B2:C3)"}
]
```

## Create JSON

```json
{
  "sheets": [
    {
      "name": "Data",
      "rows": [["Label", "Value"], ["Revenue", 123]],
      "cells": [{"cell": "C2", "value": "=B2*2"}]
    }
  ]
}
```

Formula edits mark the workbook for full recalculation on next Excel open and
remove stale `calcChain` parts.

## Notes

- Never use openpyxl or another workbook library to save/write a workbook from this project.
- Direct writes rewrite selected zip parts only. They do not load/save the whole workbook model.
- Protected sheets fail by default for direct writes. Use Excel COM for locked-cell preflight or pass `--allow-protected` only when you deliberately want package-level writes.
- The raw `office-tools package` commands are available for low-level OOXML inspection and targeted part replacement.
