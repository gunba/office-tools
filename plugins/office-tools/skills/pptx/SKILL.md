---
description: Read, build, or extract notes from Microsoft PowerPoint .pptx decks using the Rust office-tools binary.
---

# PowerPoint Tool

Use:

```bash
"${CLAUDE_PLUGIN_ROOT}/bin/office-tools.exe" pptx <command>
```

Local development:

```bash
cargo run -- pptx <command>
```

## Commands

```bash
office-tools pptx read deck.pptx
office-tools pptx notes deck.pptx
office-tools pptx build spec.json --output deck.pptx
```

## Build Spec

The Rust builder accepts neutral JSON specs with caller-provided brand values.
Supported layouts include:

- `title`
- `section`
- `content`
- `two_column`
- `three_panel`
- `quote`
- `table`
- `image_placeholder`

The builder intentionally does not ship any private or firm-specific brand data.
Slide `notes` fields are written into notes-slide parts, and
`image_placeholder.image_prompt` is copied there for follow-up image generation.
For advanced edits, use the MCP/package primitives to inspect or replace exact
OOXML parts.
