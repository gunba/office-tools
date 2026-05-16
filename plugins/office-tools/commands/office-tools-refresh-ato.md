---
description: Force-refresh the ato-mcp binary and corpus. Wipes the cached release tag so the next mcp__ato__* call re-downloads the latest gunba/ato-mcp Windows binary, then runs `ato-mcp update` to refresh the local ~4 GB corpus.
---

# Refresh ato-mcp

Run this when:
- A new ato-mcp release is published and you want this machine on the latest binary.
- The corpus is stale (the model warns "A newer ATO corpus index is available").
- The binary refuses to launch after an EDR / antivirus update and you want a clean re-install.

## What it does

1. Wipe the release-tag sentinel at `${CLAUDE_PLUGIN_DATA}\ato-mcp\release.tag` and remove the binary so the next MCP invocation triggers a fresh download from `github.com/gunba/ato-mcp/releases/latest`.
2. Invoke the installed `ato-mcp.exe update` to refresh the corpus (~4 GB download, 5-10 minutes on a typical connection, longer behind a TLS-inspecting corporate proxy).

## Run

Execute these two shell commands in order:

```bash
# 1. Clear the binary install so the next MCP launch fetches the latest tag.
del /Q "%CLAUDE_PLUGIN_DATA%\ato-mcp\release.tag" 2>nul
del /Q "%CLAUDE_PLUGIN_DATA%\ato-mcp\bin\ato-mcp.exe" 2>nul

# 2. Refresh the corpus via the installed binary's own update flow.
"%CLAUDE_PLUGIN_ROOT%\python-base\python.exe" "%CLAUDE_PLUGIN_ROOT%\bootstrap\ato_mcp_launcher.py" update
```

The second command downloads the binary if missing (silent, fast), then forwards `update` to it. Restart the MCP client after the corpus refresh completes so this server picks up the new index.
