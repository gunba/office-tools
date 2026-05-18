---
description: Create, read, search, or edit Microsoft Excel .xlsx workbooks using the Rust office-tools binary. Direct writes edit only targeted OOXML package parts and preserve unrelated XML such as conditional formatting, data validation, workbook metadata, and extension lists. Use Excel COM commands only when Excel-native behavior is required.
---

# Excel Workbook Tool

When the office-tools MCP server is loaded, use the `xlsx_*` tools directly. Otherwise invoke the binary:

```
"%LOCALAPPDATA%\Temp\office-tools\office-tools.exe" xlsx <command>
```

## When to pick which engine

| Engine | Use for | Side effects |
|--------|---------|--------------|
| Direct (`read`, `cells`, `search`, `inspect`, `edit`, `create`, `add-sheet`, `move-sheet`, `*-names`, `relationships`, `formulas`, `tables`, `validations`, `conditional-formatting`, `hyperlinks`, `comments`, `merged-ranges`, `auto-filters`, `protections`) | Closed-workbook inspection and surgical edits. Preserves unrelated XML. | None - rewrites only the package parts it touches. |
| Excel COM (`validate`, `insert`, `autofit`, `format`, `copy-sheets`, `rename-sheet`) | Anything that needs Excel-native behaviour: recalc, row/column insert with reference fix-up, copy a sheet with all its features, autofit by rendered width, rename so dependent refs update. | Spawns Excel.Application in the background. Windows + Office only. |

## Invariants

- **Never** use openpyxl, python-docx, xlwings, or any other library to write a workbook from this project. If a tool below doesn't cover a case, use `package_*` for raw OOXML.
- Formula edits via `edit` automatically remove stale `calcChain` parts and mark the workbook for full recalculation on next open.
- Protected sheets reject direct writes by default. Use Excel COM for locked-cell preflight, or pass `--allow-protected` only when you deliberately want package-level writes.
- Batch edits accept a JSON file: `[{"sheet":"S","cell":"A1","value":"x"}, {"sheet":"S","range":"B2:C3","values":[[1,2],[3,4]]}]`.
- `create` takes `{"sheets":[{"name":"Data","rows":[["h1","h2"],[1,2]],"cells":[{"cell":"C2","value":"=A2*2"}]}]}`.
