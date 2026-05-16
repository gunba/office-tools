# Completion Audit

Objective: rewrite `office-tools` in Rust with no Python requirement; provide
agent-friendly Office primitives and MCP tooling for `.xlsx`, `.docx`, `.pptx`,
and read-only Outlook; remove the legacy external corpus plugin; avoid
surfacing private firm branding data.

## Prompt-to-Artifact Checklist

| User objective phrase | Concrete deliverable | Current evidence |
| --- | --- | --- |
| "another pass" | A second-pass audit focused on agent scripting gaps, not just the initial Rust rewrite | This document's Agent-Surface Second Pass section plus integration tests for each newly surfaced primitive |
| "didn't leave out any functionality" | Office surfaces cover workbook read/create/edit/metadata, document read/replace/compose, deck read/notes/build, read-only Outlook, raw OOXML parts, and Excel/Word COM where native behavior matters | CLI/MCP commands listed below; `cargo run -- xlsx --help` shows the expanded workbook surface including `defined-names`, `merged-ranges`, `auto-filters`, and `protections`; `cargo test` exercises local primitives |
| "tools are suited for agent use" | Stable commands, JSON output, JSON input files/specs, MCP schemas, and low-level escape hatches | `xlsx cells`, metadata-listing commands, `xlsx create`, `xlsx edit --edits-file`, `docx compose`, `pptx build`, and `package_*` tools |
| "script the behaviour deterministically" | Closed-file OOXML reads/writes avoid whole-workbook save behavior; Excel COM is limited to operations where Excel-native reference updates are the deterministic owner | Direct writer tests verify preservation of validation/conditional formatting and calc-chain handling; Windows COM commands are separated and gated by `office doctor` and the smoke script |
| "as they currently like to do by manually writing xlwings/openpyxl scripts" | Common ad hoc script tasks are first-class commands: addressed cell reads, sheet IDs/visibility, package relationships, formulas, tables, validations, conditional formatting, hyperlinks, comments, defined names, merged ranges, auto-filters, protection settings, workbook creation, sparse/batch edits, sheet ordering, formatting/autofit/copy via COM | `src/xlsx/mod.rs`, `src/mcp.rs`, `tests/cli.rs`, `README.md`, and `plugins/office-tools/skills/xlsx/SKILL.md` |
| "consider what useful functionality is not surfaced" | Gap pass added missing narrow primitives where data existed only inside broad inspection output | `xlsx relationships`, `xlsx defined-names`, `xlsx merged-ranges`, `xlsx auto-filters`, `xlsx protections`, and MCP `xlsx_relationships` / `xlsx_defined_names` / `xlsx_merged_ranges` / `xlsx_auto_filters` / `xlsx_protections` added |

## Requirement Checklist

| Requirement | Artifact / Evidence | Status |
| --- | --- | --- |
| Project rewritten in Rust | `Cargo.toml`, `Cargo.lock`, `src/*.rs`, `src/xlsx/mod.rs`, `src/docx.rs`, `src/pptx.rs`, `src/outlook.rs`, `src/mcp.rs` | Implemented |
| No Python required | Deleted `plugins/office-tools/python-base`, `wheels`, `tools/*.py`, `bootstrap`, `requirements.txt`; `find . -path './target' -prune ... -iname '*.py' -o -iname '*.whl' -o -iname '*python*'` returns no files | Implemented |
| XLSX priority / deep support | `xlsx read`, `cells`, `list-sheets`, `search`, `inspect`, `relationships`, `formulas`, `tables`, `validations`, `conditional-formatting`, `hyperlinks`, `comments`, `defined-names`, `merged-ranges`, `auto-filters`, `protections`, `create`, direct `edit`, `add-sheet`, `move-sheet`, WinCOM `rename-sheet`, `validate`, `insert`, `autofit`, `format`, `copy-sheets`, raw `package` primitives | Implemented |
| Avoid openpyxl save mangling | Direct XLSX writes in `src/xlsx/mod.rs` rewrite only selected zip parts; tests assert `conditionalFormatting` and `dataValidations` survive direct edit | Verified locally |
| DOCX support | `docx read` with markdown/JSON output, direct/COM `replace`, MCP `docx_replace` engine selection, `compose`, footnotes, footer text, tables, rich text-ish segments, generic caller-provided brand values | Implemented |
| PPTX support | `pptx read` and `notes` with markdown/JSON output, `build`, notes slide generation, image-placeholder prompt notes, raw package primitives | Implemented |
| Agent scriptability | Stable CLI commands, JSON flags/outputs, `xlsx list-sheets --json` with sheet IDs and visibility state, addressed `xlsx cells`, `xlsx relationships`, `xlsx formulas`, `xlsx tables`, `xlsx validations`, `xlsx conditional-formatting`, `xlsx hyperlinks`, `xlsx comments`, `xlsx defined-names`, `xlsx merged-ranges`, `xlsx auto-filters`, `xlsx protections`, `xlsx create` from structured rows/cells, `xlsx edit --edits-file`, content-typed `package list-parts --json`, CLI `package read-part --base64`, `package write-part --base64`, and `package write-part --text`, base64/text MCP `package_read_part` and `package_write_part`, `package_delete_part`, and integration tests invoking the compiled binary | Verified locally |
| MCP server | `office-tools mcp serve`; `.mcp.json`; MCP tools for xlsx read/cells/list/search/inspect/relationships/formulas/tables/validations/conditional-formatting/hyperlinks/comments/defined-names/merged-ranges/auto-filters/protections/create/edit/add/move/rename/validate/insert/autofit/format/copy, docx read/replace with engine selection/compose, pptx read/notes/build, outlook, package, doctor | Verified locally via integration test and manual `tools/list` |
| WinCOM support | `src/wincom.rs`, `office-tools doctor`, Excel COM commands, Word COM replacement, Outlook COM read | Implemented; runtime validation external |
| Windows COM runtime smoke gate | `scripts/windows-wincom-smoke.ps1` creates temp Office files and exercises doctor, Excel validate/insert/rename/autofit/format/copy, Word COM replace, and Outlook read paths | Needs Windows + Office execution |
| Cross-platform Rust CI and release | `.github/workflows/rust.yml` runs fmt, clippy, tests, release builds on Linux and Windows runners, uploads the built binaries as `office-tools-linux-x64` and `office-tools-windows-x64` artifacts, and publishes both binaries plus `SHA256SUMS` to a GitHub Release when a `v*` tag is pushed | Implemented |
| Remove legacy external corpus plugin | Deleted its skill, command, launcher, and plugin registration | Implemented |
| Outlook read-only behavior | `src/outlook.rs` exposes search/filter/save attachments/save `.msg`; no send/reply/forward command; smoke grep found no Outlook send verbs | Implemented |
| Outlook reference coverage | Implements hours/days/since, search/sender/subject/to, include-read, full-body, folder scopes, attachment and `.msg` saving, meetings | Implemented |
| DOCX reference coverage without private firm data | Generic compose API covers title pages, headings, body/rich body, bullets/numbered lists, quote blocks, footnotes, tables, dividers, footer; no private firm strings in repo | Implemented |
| No private firm data surfaced | `cargo test --test audit` verifies this repository hygiene gate | Verified locally |

## Agent-Surface Second Pass

Objective audited: ensure the Rust tools did not leave out important
functionality, are suited for agent use, and expose deterministic scriptable
methods that replace ad hoc `xlwings`/`openpyxl` scripts.

| Requirement | Artifact / Evidence | Status |
| --- | --- | --- |
| Agents can create XLSX workbooks without Python | `office-tools xlsx create <spec.json> --output <file.xlsx>` and MCP `xlsx_create`; integration test `xlsx_cli_create_builds_workbook_from_json` verifies rows, formulas, multi-sheet workbook XML, and readback | Implemented |
| Agents can inspect sheet identity and visibility before scripting | `office-tools xlsx list-sheets --json` and MCP `xlsx_list_sheets` return sheet order, OOXML `sheetId`, and hidden/veryHidden state; integration test `xlsx_cli_list_sheets_reports_sheet_metadata` verifies hidden state and sheet ID extraction | Verified locally |
| Agents can map read values back to cell addresses | `office-tools xlsx cells` and MCP `xlsx_cells` return `{sheet, cell, value}` records with range filtering and optional empty cells; integration test `xlsx_cli_cells_lists_addressed_values` verifies addressed output | Verified locally |
| Agents can enumerate package relationships without raw XML scripts | `office-tools xlsx relationships` and MCP `xlsx_relationships` list workbook/worksheet relationship IDs, types, targets, and target modes; integration test `xlsx_cli_relationships_lists_sheet_targets` verifies hyperlink target and external target mode extraction | Verified locally |
| Agents can enumerate workbook formulas without XML scripts | `office-tools xlsx formulas` and MCP `xlsx_formulas` list sheet/cell/formula records; integration test `xlsx_cli_create_builds_workbook_from_json` verifies formula listing on a generated workbook | Verified locally |
| Agents can enumerate workbook tables without XML scripts | `office-tools xlsx tables` and MCP `xlsx_tables` list sheet/table/range records; integration test `xlsx_cli_tables_lists_table_metadata` verifies table name and range extraction | Verified locally |
| Agents can inspect data validation rules before editing | `office-tools xlsx validations` and MCP `xlsx_validations` list validation range/type/operator/formulas; integration test `xlsx_cli_validations_lists_validation_metadata` verifies a whole-number validation rule | Verified locally |
| Agents can inspect conditional formatting rules before editing | `office-tools xlsx conditional-formatting` and MCP `xlsx_conditional_formatting` list rule range/type/operator/formulas; integration test `xlsx_cli_conditional_formatting_lists_rule_metadata` verifies a cell rule | Verified locally |
| Agents can extract worksheet hyperlinks without XML scripts | `office-tools xlsx hyperlinks` and MCP `xlsx_hyperlinks` list link refs, targets, locations, display text, and tooltips; integration test `xlsx_cli_hyperlinks_lists_targets` verifies an external URL | Verified locally |
| Agents can extract worksheet comments without XML scripts | `office-tools xlsx comments` and MCP `xlsx_comments` list comment cell, author, text, and comments part; integration test `xlsx_cli_comments_lists_comment_text` verifies author and text extraction | Verified locally |
| Agents can extract workbook defined names without XML scripts | `office-tools xlsx defined-names` and MCP `xlsx_defined_names` list workbook and sheet-scoped defined names; integration test `xlsx_cli_defined_names_lists_scoped_names` verifies global and scoped names | Verified locally |
| Agents can extract merged ranges without XML scripts | `office-tools xlsx merged-ranges` and MCP `xlsx_merged_ranges` list sheet/range records; integration test `xlsx_cli_merged_ranges_lists_ranges` verifies merged range extraction | Verified locally |
| Agents can extract worksheet auto-filter ranges without XML scripts | `office-tools xlsx auto-filters` and MCP `xlsx_auto_filters` list sheet/range records; integration test `xlsx_cli_auto_filters_lists_filter_ranges` verifies range extraction | Verified locally |
| Agents can inspect sheet protection settings before editing | `office-tools xlsx protections` and MCP `xlsx_protections` list sheet protection attributes; integration test `xlsx_cli_protections_lists_sheet_protection_settings` verifies protected-sheet metadata extraction | Verified locally |
| Agents can still do sparse and batch XLSX edits deterministically | `xlsx edit --edits-file`, MCP `xlsx_edit`, matrix/range/cell JSON handling, formula recalc marking, preservation tests for data validation and conditional formatting | Verified locally |
| MCP XLSX edits preserve JSON scalar types | `xlsx_edit` MCP call path now builds a JSON edit payload instead of string-only CLI args; integration test `mcp_xlsx_edit_accepts_json_scalar_values` writes numeric `456` as a number cell | Verified locally |
| MCP exposes the same high-value primitives as the CLI | MCP tools now include `xlsx_cells`, `xlsx_relationships`, `xlsx_formulas`, `xlsx_tables`, `xlsx_validations`, `xlsx_conditional_formatting`, `xlsx_hyperlinks`, `xlsx_comments`, `xlsx_defined_names`, `xlsx_merged_ranges`, `xlsx_auto_filters`, `xlsx_protections`, `xlsx_create`, `xlsx_rename_sheet`, `xlsx_validate`, `xlsx_insert`, `xlsx_autofit`, `xlsx_format`, `xlsx_copy_sheets`, `docx_compose`, `pptx_build`, and `package_delete_part`; manual `tools/list` output confirms all names | Verified locally |
| MCP build/create and package-write tools are schema-guided for agents | `xlsx_create`, `docx_compose`, and `pptx_build` schemas expose expected spec fields such as `date1904`, `sheets`, `blocks`, and `slides`; `package_write_part` schema uses `oneOf` to require exactly one of `text` or `base64`; integration test asserts those schema hints are present in `tools/list` | Verified locally |
| MCP DOCX tools are executable as a round trip | `docx_compose`, `docx_replace`, and `docx_read` are exercised together through MCP; `docx_replace` schema accepts `engine: "direct" | "com"` and call handling passes that through to the shared DOCX replacement implementation; integration test `mcp_docx_replace_accepts_engine_argument` verifies compose/replace/read on the direct engine path | Verified locally |
| Agents have Excel-native fallbacks for common `xlwings` presentation/structure tasks | CLI and MCP expose `xlsx_rename_sheet` and `xlsx_autofit` through Excel COM, avoiding unsafe direct XML rewrites for reference-sensitive sheet renames and Excel-owned sizing | Implemented; runtime validation external |
| Agents can apply common Excel-native formatting without ad hoc COM scripts | CLI and MCP expose `xlsx_format` for range number formats, bold/italic, wrap text, fill color, and font color; Windows smoke script exercises it | Implemented; runtime validation external |
| Agents can deterministically diagnose COM availability | CLI `office-tools doctor` and MCP `office_doctor` return machine-readable JSON even on non-Windows hosts; integration tests `office_doctor_returns_machine_readable_json` and `mcp_office_doctor_returns_machine_readable_json` verify platform/wincom output shape | Verified locally; Windows COM runtime validation external |
| Agents can script raw OOXML bytes and text without temporary payload files | CLI `package read-part --base64`, `package write-part --base64`, `package write-part --text`, and MCP `package_read_part`/`package_write_part` support inline payloads; CLI and MCP package listings report declared content types for agent filtering; integration tests write/read/delete `custom/data.bin`, verify base64 CLI read/write, verify text CLI write, and verify CLI/MCP content type listing | Verified locally |
| MCP PPTX tools are executable, not only listed | `pptx_build`, `pptx_read`, and `pptx_notes` are exercised together through MCP; integration test `mcp_can_build_and_read_pptx_notes` verifies slide text and speaker-note round trip | Verified locally |
| COM commands are safe to expose over MCP framing | Excel COM helpers return JSON strings to callers; CLI prints the returned string, MCP wraps it in tool content; `rg` shows COM helpers no longer print directly from `src/wincom.rs` | Verified locally |
| Docs/skills advertise the agent-facing surface | `README.md`, `plugins/office-tools/skills/xlsx/SKILL.md`, and this audit mention `xlsx create` and expanded MCP primitives | Implemented |

## Local Verification Commands

These pass in the current Linux workspace:

```bash
cargo check
cargo check --target x86_64-pc-windows-gnu
cargo test --target x86_64-pc-windows-gnu --no-run
cargo clippy --all-targets -- -D warnings
cargo test
cargo build --release
cargo run -- xlsx create --help
```

Additional audits run:

```bash
find . -path './.git' -prune -o -path './target' -prune -o \
  \( -iname '*python*' -o -iname '*.py' -o -iname '*.pyc' -o -iname '*.whl' \) -print

cargo test --test audit
```

## Uncovered Gate

Actual WinCOM runtime behavior cannot be completed in this Linux workspace.
The remaining gate is:

```powershell
cargo build --release
.\scripts\windows-wincom-smoke.ps1
```

Run it on Windows with Microsoft Excel, Word, and Outlook installed and an
Outlook profile configured. This is the only known uncovered requirement.
