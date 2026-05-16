---
description: Inspect or surgically edit raw OOXML zip package parts in .xlsx, .docx, and .pptx files using the Rust office-tools package commands. Use when high-level commands do not cover a needed primitive.
---

# OOXML Package Primitives

Use:

```bash
"${CLAUDE_PLUGIN_ROOT}/bin/office-tools.exe" package <command>
```

Local development:

```bash
cargo run -- package <command>
```

## Commands

```bash
office-tools package list-parts file.xlsx --json
office-tools package read-part file.xlsx xl/workbook.xml
office-tools package read-part file.xlsx xl/media/image1.png --base64
office-tools package read-part file.docx word/document.xml --output document.xml
office-tools package write-part file.pptx ppt/slides/slide1.xml --input slide1.xml
office-tools package write-part file.xlsx custom/info.xml --text "<root/>"
office-tools package write-part file.xlsx custom/data.bin --base64 "AAEC/w=="
office-tools package delete-part file.xlsx xl/calcChain.xml
```

`list-parts --json` includes each part name, content type when declared by
`[Content_Types].xml`, compressed size, and uncompressed size.
Use `read-part --base64`, `write-part --base64`, and `write-part --text` for
deterministic stdout/stdin-style scripting without staging temporary payload
files.

These commands operate on exact zip package parts. They are intended for agent
scripts that need a primitive not yet exposed by the typed XLSX/DOCX/PPTX
commands.
