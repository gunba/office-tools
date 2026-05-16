# office-tools

A Claude Code / Codex plugin bundling four office-productivity capabilities for Windows agents:

| Skill        | Purpose |
|--------------|---------|
| `xlsx`       | Read, search, edit Excel workbooks. OOXML engine for closed workbooks, Excel COM for live ones. |
| `docx`       | Read and find/replace Word documents via python-docx + Word COM. |
| `pptx`       | Build 16:9 PowerPoint decks from a JSON spec. Branding is read from `brand.toml` — defaults are neutral; override via plugin user config. |
| `ato` (MCP)  | Search and retrieve the Australian Taxation Office legal corpus through the [ato-mcp](https://github.com/gunba/ato-mcp) Rust server. |

The repository **is** a local marketplace plus its single plugin. Both Claude Code and Codex accept the same `.claude-plugin/plugin.json` manifest, so one install pipeline covers both.

## Install

Clone or download a zip of this repo into a non-OneDrive path (`%LOCALAPPDATA%` is a good default; OneDrive-synced paths break the embedded Python and the ato-mcp corpus). Then:

```bat
:: From the repo root
install.bat
```

The script runs the appropriate `plugin marketplace add` + `plugin install` commands for whichever of `claude` / `codex` are on PATH. The plugin's Python dependencies (openpyxl, xlwings, python-docx, python-pptx, lxml, pywin32) lazily pip-install into the plugin's persistent data directory on the first tool call — there is nothing to install manually.

Equivalent manual install:

```bat
claude plugin marketplace add <repo-path>
claude plugin install office-tools@office-tools

codex plugin marketplace add <repo-path>
codex plugin add office-tools@office-tools
```

Codex versions prior to 0.129 do not ship `codex plugin add`; `install.bat` detects this and instructs the user to upgrade Codex.

## First-use caveat: ato-mcp

The `ato` MCP server is a Rust binary, downloaded lazily on first invocation into the plugin's persistent data directory. Two things to expect on a managed Windows machine:

1. **EDR sandbox (~20 minutes).** The binary is unsigned; Windows Defender / CrowdStrike / SentinelOne typically hold it in a sandbox queue before allowing execution. MCP calls fail until EDR releases the binary. This is normal — wait it out.
2. **Corpus install (~4 GB).** The binary itself does not ship the corpus. On first MCP call it asks the host: "ATO corpus not yet installed; run `ato-mcp update`". Do that once in a terminal, then restart the MCP client. See [skills/ato/SKILL.md](plugins/office-tools/skills/ato/SKILL.md) for the absolute path.

## Uninstall

```bat
uninstall.bat
```

The plugin's cache directory, skills, and MCP server registration are removed by the CLI. Persistent plugin data (cached corpus, pip-installed wheels) goes with the `plugin uninstall` step.

## Enterprise deployment

To push office-tools to many laptops without each user running `install.bat`, drop a managed config that pre-registers the marketplace and enables the plugin:

**Claude Code** — `C:\Program Files\ClaudeCode\managed-settings.json`:

```json
{
  "extraKnownMarketplaces": {
    "office-tools": {
      "source": { "source": "directory", "path": "\\\\fileserver\\share\\office-tools" }
    }
  },
  "enabledPlugins": {
    "office-tools@office-tools": true
  }
}
```

**Codex** — `%ProgramData%\OpenAI\Codex\managed_config.toml`:

```toml
[marketplaces.office-tools]
source_type = "local"
source      = "\\\\fileserver\\share\\office-tools"

[plugins."office-tools@office-tools"]
enabled = true
```

Distribute the repo tree to a file share or bake it into the laptop image. The plugin handles its own dependency install lazily on first tool call.

## Brand override (pptx)

The pptx tool's default palette is neutral dark blue and grey, no logo. To match a firm or team brand, set the `BRAND_TOML_PATH` plugin user config to the absolute path of a `brand.toml` you maintain elsewhere on disk. See [plugins/office-tools/tools/pptx/brand.default.toml](plugins/office-tools/tools/pptx/brand.default.toml) for the schema; any subset of keys overrides the corresponding defaults.

## Layout

```
office-tools/
├── .claude-plugin/marketplace.json        # marketplace manifest (Claude Code)
├── .agents/plugins/marketplace.json       # marketplace manifest (Codex)
├── install.bat
├── uninstall.bat
└── plugins/office-tools/                  # the plugin
    ├── .claude-plugin/plugin.json         # honoured by both Claude Code and Codex
    ├── .mcp.json                          # registers the ato MCP server
    ├── skills/{xlsx,docx,pptx,ato,hello}  # SKILL.md per skill
    ├── tools/{xlsx,docx,pptx}             # Python tool sources
    ├── bootstrap/                         # pip wrapper, ato-mcp launcher
    ├── python-base/                       # embedded Windows Python 3.13
    ├── wheels/                            # offline-install dependency wheels
    └── requirements.txt
```

## License

MIT — see [LICENSE](LICENSE).
