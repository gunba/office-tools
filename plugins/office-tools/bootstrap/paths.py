"""Resolve plugin root and persistent data directory in both host CLIs.

Both Claude Code and Codex set CLAUDE_PLUGIN_ROOT and CLAUDE_PLUGIN_DATA in the
environment when invoking plugin processes:
  - Claude Code: src/utils/plugins/mcpPluginIntegration.ts (resolvePluginMcpEnvironment)
                 src/utils/hooks.ts:890
  - Codex:       codex-rs/hooks/src/engine/discovery.rs:223-228 (PLUGIN_ROOT + CLAUDE_PLUGIN_ROOT)

We try the canonical env vars first and fall back to filesystem heuristics so
the same code works in interactive dev (running tools directly from a clone).
"""

from __future__ import annotations

import os
import sys
from pathlib import Path


def plugin_root() -> Path:
    """Absolute path to the plugin's install directory.

    Where the plugin's tools/, skills/, bootstrap/, wheels/, python-base/ live.
    Read-only once the plugin is installed — never write here.
    """
    env_value = os.environ.get("CLAUDE_PLUGIN_ROOT") or os.environ.get("PLUGIN_ROOT")
    if env_value:
        return Path(env_value).resolve()
    # Dev fallback: walk up from this file until we find .claude-plugin/plugin.json.
    here = Path(__file__).resolve()
    for ancestor in here.parents:
        if (ancestor / ".claude-plugin" / "plugin.json").is_file():
            return ancestor
    raise RuntimeError(
        "Cannot locate plugin root: neither CLAUDE_PLUGIN_ROOT nor PLUGIN_ROOT is "
        "set, and no .claude-plugin/plugin.json was found in ancestors of "
        f"{here}."
    )


def plugin_data_dir() -> Path:
    """Absolute path to the plugin's persistent data directory.

    Survives plugin version bumps and is where we install pip wheels, download
    the ato-mcp binary, and store the corpus.
    """
    env_value = os.environ.get("CLAUDE_PLUGIN_DATA") or os.environ.get("PLUGIN_DATA")
    if env_value:
        path = Path(env_value).resolve()
    else:
        # Dev fallback: drop alongside the plugin root so dev runs don't pollute
        # the user's actual cache.
        path = plugin_root().parent / "office-tools-data"
    path.mkdir(parents=True, exist_ok=True)
    return path


def site_packages_dir() -> Path:
    path = plugin_data_dir() / "site-packages"
    path.mkdir(parents=True, exist_ok=True)
    return path


def wheels_dir() -> Path:
    return plugin_root() / "wheels"


def requirements_file() -> Path:
    return plugin_root() / "requirements.txt"
