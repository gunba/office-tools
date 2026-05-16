---
description: Read or edit Microsoft Word (.docx) documents while preserving all formatting, styles, headers, and footers. Reading uses python-docx; editing uses Word COM automation. Use whenever the user asks to read, summarise, or modify a .docx file, or to perform find/replace operations in Word.
---

# Word Document Tool (docx)

The tool is the Python script at `${CLAUDE_PLUGIN_ROOT}/tools/docx/docx_tool.py`, invoked through the plugin's bundled Python at `${CLAUDE_PLUGIN_ROOT}/python-base/python.exe`. Reading uses python-docx (pure Python). Editing uses Word COM automation, so all formatting, styles, headers/footers are preserved exactly.

On Codex the same paths resolve via this skill's directory: from the skill root, the tool is at `../../tools/docx/docx_tool.py` and the interpreter at `../../python-base/python.exe`.

## Usage

In every example, `<PY>` stands for `${CLAUDE_PLUGIN_ROOT}/python-base/python.exe` and `<TOOL>` for `${CLAUDE_PLUGIN_ROOT}/tools/docx/docx_tool.py`.

```bash
# Read a document as markdown
<PY> <TOOL> read <file>

# Single find and replace (in-place)
<PY> <TOOL> replace <file> --find "old text" --replace "new text"

# Find and replace, saving to a new file (copy + edit)
<PY> <TOOL> replace <file> --find "old" --replace "new" --output "new_file.docx"

# Batch replacements
<PY> <TOOL> replace <file> --replacements '[{"find":"old1","replace":"new1"},{"find":"old2","replace":"new2"}]'
```

## Notes

- `read` works on any platform; it parses the .docx package with python-docx and does not need Word installed.
- `replace` uses Word COM (pywin32). It requires Microsoft Word to be installed on the machine and only runs on Windows. The find-replace is dispatched through Word so formatting, headers, footers, styles, fields, tables, and comments are preserved exactly.

## Working files and ephemera

Do not write helper scripts, JSON specs, or intermediate artefacts into `${CLAUDE_PLUGIN_ROOT}` or its subdirectories — the plugin install directory is treated as read-only. For batch replacement JSON or other transient data, use Python's `tempfile` module or write into the user's current working directory with an explicit, short-lived filename and delete it when done.
