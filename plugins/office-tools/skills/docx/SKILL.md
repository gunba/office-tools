---
description: Read, replace, or compose Microsoft Word .docx documents using the Rust office-tools binary. Supports direct OOXML text replacement, generic branded document composition, and optional Word COM replacement on Windows.
---

# Word Document Tool

Use:

```bash
"%LOCALAPPDATA%\Temp\office-tools\office-tools.exe" docx <command>
```

Local development:

```bash
cargo run -- docx <command>
```

## Commands

```bash
# Read body paragraphs and tables as markdown.
office-tools docx read input.docx

# Direct XML text-node replacement. Good when matches are inside a single run.
office-tools docx replace input.docx --find "old" --replace "new" --output output.docx

# Word COM replacement. Use for formatting-sensitive documents or matches split across runs.
office-tools docx replace input.docx --find "old" --replace "new" --engine com

# MCP docx_replace accepts the same engine values: "direct" or "com".

# Compose a new generic document from JSON. Branding is supplied by the caller.
office-tools docx compose spec.json --output output.docx
```

## Compose Blocks

The JSON composer supports generic equivalents of the reference workflow:

- `title_page`
- `heading`
- `body`
- `body_rich`
- `bullet`
- `numbered`
- body/bullet/numbered `footnote` fields
- `quote`
- `quote_block`
- `table`
- `borderless_table`
- `divider`
- `spacer`
- `page_break`

Brand values such as fonts and colors are provided in the spec. No firm-specific
branding data is shipped in the repository.
