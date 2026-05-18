# office-tools

Rust Office document primitives for agents. The project no longer bundles or
requires Python, openpyxl, xlwings, python-docx, python-pptx, or pywin32.

| Area | Purpose |
| --- | --- |
| `xlsx` | Deepest support. Read/search with Rust, edit workbook XML package parts directly, preserve unrelated OOXML such as conditional formatting and data validation, and use Excel COM only when Excel-native behavior is required. |
| `docx` | Read as markdown, direct XML replacement, generic document composition, and optional Word COM replacement for formatting-sensitive edits. |
| `pptx` | Read slide text, extract notes, and build simple 16:9 decks from JSON specs. |
| `outlook` | Read-only Outlook COM mail search, meeting request extraction, `.msg` saving, and attachment saving. No send command exists. |
| `package` | Raw OOXML zip part list/read/write/delete primitives for agent scripts. |
| `mcp` | Stdio MCP server exposing the same primitives as tools. |

## Build

```bash
cargo build --release
```

The release binary is `target\release\office-tools.exe`. Most users do not need
to build — see [Install](#install) for the prebuilt path.

CI runs release builds on Linux and Windows and uploads the resulting binaries
as workflow artifacts, including `office-tools-windows-x64`. Pushing a `v*`
tag, such as `v0.1.0`, also publishes a GitHub Release with the Windows `.exe`,
the Linux binary, and `SHA256SUMS`.

## CLI Examples

```bash
office-tools xlsx list-sheets workbook.xlsx
office-tools xlsx list-sheets workbook.xlsx --json  # includes sheet_id and hidden/veryHidden state
office-tools xlsx read workbook.xlsx --sheet "Sheet1" --range A1:D20
office-tools xlsx cells workbook.xlsx --sheet "Sheet1" --range A1:D20 --json
office-tools xlsx search workbook.xlsx --query "revenue"
office-tools xlsx inspect workbook.xlsx --sheet "Sheet1"
office-tools xlsx relationships workbook.xlsx --sheet "Sheet1" --json
office-tools xlsx formulas workbook.xlsx --sheet "Sheet1" --json
office-tools xlsx tables workbook.xlsx --sheet "Sheet1" --json
office-tools xlsx validations workbook.xlsx --sheet "Sheet1" --json
office-tools xlsx conditional-formatting workbook.xlsx --sheet "Sheet1" --json
office-tools xlsx hyperlinks workbook.xlsx --sheet "Sheet1" --json
office-tools xlsx comments workbook.xlsx --sheet "Sheet1" --json
office-tools xlsx defined-names workbook.xlsx --json
office-tools xlsx merged-ranges workbook.xlsx --sheet "Sheet1" --json
office-tools xlsx auto-filters workbook.xlsx --sheet "Sheet1" --json
office-tools xlsx protections workbook.xlsx --sheet "Sheet1" --json
office-tools xlsx create workbook.json --output workbook.xlsx
office-tools xlsx edit workbook.xlsx --sheet "Sheet1" --cell B2 --value "123"
office-tools xlsx edit workbook.xlsx --edits-file edits.json
office-tools xlsx add-sheet workbook.xlsx Support --after Sheet1
office-tools xlsx move-sheet workbook.xlsx Support --before Sheet1
office-tools xlsx rename-sheet workbook.xlsx --sheet Support --name Analysis
office-tools xlsx autofit workbook.xlsx --sheet Analysis --axis columns
office-tools xlsx format workbook.xlsx --sheet Analysis --range B:B --number-format "#,##0.00"

office-tools docx read input.docx
office-tools docx replace input.docx --find old --replace new --output output.docx
office-tools docx replace input.docx --find old --replace new --engine com
office-tools docx compose spec.json --output output.docx

office-tools pptx read deck.pptx
office-tools pptx notes deck.pptx
office-tools pptx build deck.json --output deck.pptx

office-tools outlook --days 7 --include-read --search "invoice" --folder received

office-tools package list-parts workbook.xlsx --json  # includes content_type and sizes
office-tools package read-part workbook.xlsx xl/workbook.xml
office-tools package read-part workbook.xlsx xl/media/image1.png --base64
office-tools package write-part workbook.xlsx custom/info.xml --text "<root/>"
office-tools package write-part workbook.xlsx custom/data.bin --base64 "AAEC/w=="

office-tools doctor
```

## Windows COM Smoke Test

On a Windows machine with Microsoft Office and Outlook configured:

```powershell
cargo build --release
.\scripts\windows-wincom-smoke.ps1
```

The smoke test is non-destructive: it creates temporary Excel/Word files,
exercises Excel validate/insert/copy-sheets, Word COM replace, and one read-only
Outlook query, then deletes its temporary directory unless `-KeepTemp` is passed.

See [docs/completion-audit.md](docs/completion-audit.md) for the requirement
mapping and current verification status.

## MCP

The plugin registers:

```bash
office-tools mcp serve
```

The MCP server exposes tools for XLSX read/cells/list/search/inspect/relationships/formulas/tables,
validations, conditional formatting, hyperlinks, comments, defined names, merged
ranges, auto-filter ranges, and protection settings, create/edit, sheet add/move/rename,
validate/insert/autofit/format/copy, DOCX read/replace/compose with engine
selection, PPTX read/notes/build, read-only Outlook mail search, and raw OOXML
package part operations.

## Install

The plugin loads `office-tools.exe` from `%LOCALAPPDATA%\Temp\office-tools\` at
runtime. The repo no longer ships a `plugins\office-tools\bin\` directory or a
batch installer — the binary is staged from a GitHub release.

The default install path is `%LOCALAPPDATA%\Temp\office-tools` (i.e. `%TEMP%\office-tools`)
on purpose: on a typical corporate Windows image, Defender Attack Surface
Reduction rule "Block executable files… not from a trusted list" refuses to
exec unsigned binaries from most user-profile paths (`Desktop\`, `Documents\`,
`.cargo\bin\`, `AppData\Local\Programs\`, `C:\dev\`, etc.) but allows exec from
`%TEMP%`. Installing there keeps the plugin runnable on locked-down endpoints
without filing an IT ticket. Pass `-InstallDir <path>` to override on
unrestricted machines.

### Recommended: `install.ps1`

From the repo root on Windows:

```powershell
powershell -ExecutionPolicy Bypass -File .\install.ps1
```

This downloads `office-tools-windows-x64.zip` from the latest release, extracts
it into `%LOCALAPPDATA%\Temp\office-tools\`, verifies the binary runs, and
(when the CLIs are on `PATH`) registers the marketplace with Claude Code and
Codex.

Flags:

- `-Tag v0.1.0` — pin to a specific release tag.
- `-InstallDir <path>` — override binary location.
- `-SkipRegister` — only stage the binary; do not touch Claude/Codex configs.

### Manual install

If you prefer to wire things up by hand:

1. Download `office-tools-windows-x64.zip` from
   <https://github.com/gunba/office-tools/releases/latest> and extract to
   `%LOCALAPPDATA%\Temp\office-tools\` (so the path becomes
   `%LOCALAPPDATA%\Temp\office-tools\office-tools.exe`).
2. Register the marketplace:
   ```
   claude plugin marketplace add <path-to-this-repo>
   claude plugin install office-tools@office-tools
   ```
   ```
   codex plugin marketplace add <path-to-this-repo>
   ```
3. For Codex, add an MCP server entry to `~/.codex/config.toml`:
   ```toml
   [mcp_servers.office-tools]
   command = 'C:\Users\<you>\AppData\Local\Temp\office-tools\office-tools.exe'
   args = ["mcp", "serve"]
   enabled = true
   ```
4. For Claude Code, add the same server to `~/.claude.json` `mcpServers`:
   ```json
   "office-tools": {
     "type": "stdio",
     "command": "C:\\Users\\<you>\\AppData\\Local\\Temp\\office-tools\\office-tools.exe",
     "args": ["mcp", "serve"]
   }
   ```

### Endpoint security override

If you have an unrestricted machine and prefer a stable per-app dir over
`%TEMP%`, install with:

```powershell
powershell -ExecutionPolicy Bypass -File .\install.ps1 -InstallDir "$env:LOCALAPPDATA\office-tools"
```

You will then need to update the path in `plugins\office-tools\.mcp.json`, the
five `plugins\office-tools\skills\*\SKILL.md` files, and the matching Claude
and Codex MCP server entries to match. The default ships pointing at the
`%TEMP%` path because that is the path that is exec-permitted on the broadest
set of corporate endpoints.

### Uninstall

```powershell
powershell -ExecutionPolicy Bypass -File .\uninstall.ps1
```

Removes the binary directory and unregisters the marketplace from Claude Code
and Codex.

## Layout

```
office-tools/
├── Cargo.toml
├── src/                         # Rust CLI, library, and MCP server
├── install.ps1
├── uninstall.ps1
└── plugins/office-tools/        # pure metadata; no binaries
    ├── .claude-plugin/plugin.json
    ├── .mcp.json                # points at %LOCALAPPDATA%\Temp\office-tools\office-tools.exe
    └── skills/{xlsx,docx,pptx,outlook,ooxml}
```

The runtime binary lives outside the repo, at `%LOCALAPPDATA%\Temp\office-tools\office-tools.exe`.

## License

MIT. See [LICENSE](LICENSE).
