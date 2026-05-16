---
description: Look up Australian tax law using the local ATO Legal Database via the ato-mcp server bundled with this plugin. Use when researching specific legislative provisions (e.g. section 355-210 ITAA 1997), ATO public rulings (TR/TD), interpretative decisions (ATO IDs), taxpayer alerts, practical compliance guidelines, law administration practice statements, edited private advice, or case law (AAT/Federal Court/High Court). Prefer this over web search for ATO law; corpus defaults exclude Edited Private Advice and pre-2000 non-legislation.
---

# ATO Legal Database (ato MCP server)

The ato-mcp server runs as an MCP server registered by this plugin. Its tools appear in your environment as `mcp__ato__search`, `mcp__ato__get_chunks`, `mcp__ato__get_doc_anchors`, `mcp__ato__get_definition`, `mcp__ato__get_asset`, `mcp__ato__fetch_external_doc`, and `mcp__ato__stats`. You do not call any local script — the host CLI launches the server through `${CLAUDE_PLUGIN_ROOT}/bootstrap/ato_mcp_launcher.py`, which lazily downloads the Rust binary on first use.

## Typical workflow

1. **Search first.** `mcp__ato__search` accepts a free-text query and returns chunk pointers (chunk IDs, doc IDs, titles, scores). Search hits are pointers, not authority.
2. **Fetch chunk bodies.** Pass the chunk IDs from step 1 to `mcp__ato__get_chunks` to read the actual text. Optionally include `before` / `after` neighbour chunks for context.
3. **Navigate within a document.** `mcp__ato__get_doc_anchors` returns in-doc anchors, related/history links, and reverse citations for a doc ID.
4. **Resolve statutory definitions.** `mcp__ato__get_definition` returns compact definitions for a defined term, with an ordinary-meaning fallback when no statutory definition exists.
5. **Verify install / corpus version.** `mcp__ato__stats` reports index version, counts, and the default search policy.
6. **Reach outside the corpus.** For specific `[doc:X]` links the local corpus does not contain, `mcp__ato__fetch_external_doc` retrieves them from ato.gov.au.

## Search hints

- Section number: `355-25`, `8-1`, `40-880`
- Topic keyword: `research and development`, `capital gains`, `fringe benefit`
- Case name: `TDS Biz`, `Moreton Resources`
- Ruling number: `TR 2021/3`, `TD 2024`
- ATO ID: `AID2010115`
- Default search excludes Edited Private Advice (`EV`) and pre-2000 non-legislation. Set `include_old=true` and `current_only=false` to widen.

## First-use caveats

- **Binary download.** The first invocation downloads `ato-mcp.exe` (~10 MB) into the plugin's persistent data dir. No action required.
- **EDR lockout.** The binary is unsigned. Windows Defender / CrowdStrike / SentinelOne typically sandbox it for ~20 minutes before allowing execution. During that window MCP calls fail with a connection error. Wait it out — the corpus install in the next step takes longer than the EDR delay anyway.
- **Corpus install.** The binary does NOT ship the corpus. On the first MCP call after the EDR allows execution, ato-mcp will tell the host: "ATO corpus is not yet installed on this machine. Tell the user to run `ato-mcp update`". Run that command once in a terminal — it downloads ~4 GB and takes 5-10 minutes on a typical connection (longer behind a TLS-inspecting corporate proxy). After it completes, restart the MCP client so this server picks up the corpus.

The corpus install command:

```bash
"%LOCALAPPDATA%\AnthropicClaude\plugins\cache\office-tools\office-tools\<version>\python-base\python.exe" "%LOCALAPPDATA%\AnthropicClaude\plugins\data\office-tools\ato-mcp\bin\ato-mcp.exe" update
```

(The bundled `ato-mcp.exe` is not on PATH; invoke it via its absolute path under the plugin's data dir, or use the `/office-tools-refresh-ato` slash command which wraps this.)

## Authority and verification

ato-mcp is retrieval infrastructure, not tax advice. Always verify cited ATO material against the canonical URL returned in search results, and apply professional judgment before relying on an answer.
