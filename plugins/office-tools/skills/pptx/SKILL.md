---
description: Build, read, or extract notes from PowerPoint (.pptx) decks via the office-tools pptx_tool.py. Build mode generates a 16:9 deck from a JSON spec with title/section/content/two_column/three_panel/image/image_placeholder/quote/table/big_text layouts. The visual style (colors, fonts, logo) is read from brand.toml — defaults to a neutral palette, override via the BRAND_TOML_PATH user config or by dropping a brand.toml next to the tool. Use whenever the user asks to create a slide deck, generate slides, draft a PowerPoint, or read/extract notes from an existing .pptx.
---

# PowerPoint Tool (pptx)

The tool is the Python script at `${CLAUDE_PLUGIN_ROOT}/tools/pptx/pptx_tool.py`, invoked through the plugin's bundled Python at `${CLAUDE_PLUGIN_ROOT}/python-base/python.exe`. It uses python-pptx under the hood and reads visual constants from `${CLAUDE_PLUGIN_ROOT}/tools/pptx/brand.py` (which in turn loads either `brand.toml` or `brand.default.toml`).

On Codex the same paths resolve via this skill's directory: from the skill root, the tool is at `../../tools/pptx/pptx_tool.py` and the interpreter at `../../python-base/python.exe`.

## Usage

In every example, `<PY>` stands for `${CLAUDE_PLUGIN_ROOT}/python-base/python.exe` and `<TOOL>` for `${CLAUDE_PLUGIN_ROOT}/tools/pptx/pptx_tool.py`.

```bash
# Build a deck from a JSON spec
<PY> <TOOL> build <spec.json> [--output deck.pptx]

# Dump deck contents (titles, body, tables, notes) as text
<PY> <TOOL> read <file.pptx>

# Speaker notes only
<PY> <TOOL> notes <file.pptx>
```

## Spec format

```json
{
  "meta": {
    "title": "Example deck title",
    "subtitle": "Subtitle line",
    "author": "Author name",
    "date": "Month Year",
    "logo": "C:/path/to/logo.png",
    "footer_text": "Footer line | Subtitle | Month Year"
  },
  "slides": [
    {"layout": "title",   "title": "...", "subtitle": "...", "notes": "..."},
    {"layout": "section", "number": "01", "title": "...", "subtitle": "..."},
    {"layout": "content", "title": "...", "bullets": ["...", {"text":"...","sub":["..."]}]},
    {"layout": "two_column", "title": "...",
        "left_title": "Left", "left_bullets": ["..."],
        "right_title": "Right", "right_bullets": ["..."]},
    {"layout": "three_panel", "title": "...",
        "panels": [
          {"title": "A", "bullets": ["..."]},
          {"title": "B", "bullets": ["..."]},
          {"title": "C", "bullets": ["..."]}
        ]},
    {"layout": "image", "title": "...", "image": "path/to.png", "caption": "..."},
    {"layout": "image_placeholder", "title": "...",
        "placeholder_text": "Architecture diagram",
        "image_prompt": "Detailed image prompt — appended to speaker notes"},
    {"layout": "quote", "quote": "...", "attribution": "..."},
    {"layout": "table", "title": "...", "headers": ["A","B"], "rows": [["1","2"]]},
    {"layout": "big_text", "title": "...", "body": "Single large statement"}
  ]
}
```

## Branding

Default palette is a neutral dark blue / grey 16:9 layout with no logo. To override:

- **Claude Code**: set the `BRAND_TOML_PATH` user config value for this plugin (via the plugin settings UI) to the absolute path of a brand.toml file you maintain anywhere on disk. The plugin reads it on the next tool invocation.
- **Codex / manual**: set the `BRAND_TOML_PATH` env var on the process before invoking the tool, e.g. `set BRAND_TOML_PATH=C:\path\to\mybrand.toml` in the same shell. Alternatively, drop a `brand.toml` directly next to `brand.default.toml` in the plugin install dir — but it will be overwritten on plugin update, so the env-var / userConfig approach is preferred.

The TOML schema is documented in `${CLAUDE_PLUGIN_ROOT}/tools/pptx/brand.default.toml`. Any subset of keys is valid; missing keys fall through to the shipped defaults.

## Notes

- All slides except `title` and `section` get the standard chrome: optional logo top-right, accent-color footer bar, footer text + page number.
- `image_placeholder` renders a dashed pale-blue placeholder box and copies the `image_prompt` into the slide's speaker notes — drop the generated image into the slide later.
- Bullets accept either plain strings or `{"text": "...", "sub": [...]}`.
- All writes go through python-pptx; this tool does not require PowerPoint to be installed.

## Working files and ephemera

Do not write helper scripts, JSON specs, or intermediate artefacts into `${CLAUDE_PLUGIN_ROOT}` or its subdirectories — the plugin install directory is treated as read-only. For deck spec JSON or other transient data, use Python's `tempfile` module or write into the user's current working directory with an explicit, short-lived filename and delete it when done.
