"""Lazy installer + launcher for the ato-mcp Rust binary.

Registered in the plugin's `.mcp.json` as the command Claude Code / Codex run
when an `mcp__ato__*` tool is invoked. On the first call (or after a release
tag change) this script:

  1. Queries gunba/ato-mcp's latest GitHub release.
  2. Downloads `ato-mcp-x86_64-pc-windows-msvc.zip` into
     `${CLAUDE_PLUGIN_DATA}/ato-mcp/bin/` and extracts ato-mcp.exe +
     onnxruntime.dll.
  3. Sets `ATO_MCP_DATA_DIR=${CLAUDE_PLUGIN_DATA}/ato-mcp/data` so the binary's
     corpus location lands inside our persistent plugin data dir (and crucially
     NOT in `%APPDATA%\\Roaming\\`, which is OneDrive-synced on the firm's
     enterprise tenant).
  4. Replaces the Python process with ato-mcp.exe.

The first launch of ato-mcp itself can take ~20 minutes on locked-down
Windows machines because Defender / CrowdStrike / SentinelOne sandbox the
unsigned Rust binary before allowing execution. This is expected; the launcher
does not preflight or timeout the binary.

The corpus is NOT downloaded here. The Windows release ships only the binary;
the corpus comes via `ato-mcp update`, which the binary itself prompts the
user to run on first MCP call when it sees no corpus on disk.
"""

from __future__ import annotations

import os
import platform
import sys
import zipfile
from pathlib import Path

if sys.platform == "win32":
    # Same platform.uname workaround as run_pip.py — endpoint security agents
    # stall the registry read inside platform.win32_ver() on unsigned Python.
    platform._uname_cache = platform.uname_result(  # type: ignore[attr-defined]
        system="Windows",
        node=os.environ.get("COMPUTERNAME", ""),
        release="",
        version="",
        machine=os.environ.get("PROCESSOR_ARCHITECTURE", "AMD64"),
    )

# Make the bootstrap package importable when invoked as a script.
_HERE = Path(__file__).resolve().parent
_PLUGIN_ROOT_FALLBACK = _HERE.parent
if str(_PLUGIN_ROOT_FALLBACK) not in sys.path:
    sys.path.insert(0, str(_PLUGIN_ROOT_FALLBACK))

from bootstrap import github_release, paths  # noqa: E402

REPO_OWNER = "gunba"
REPO_NAME = "ato-mcp"
WINDOWS_ASSET = "ato-mcp-x86_64-pc-windows-msvc.zip"


def _ato_root() -> Path:
    root = paths.plugin_data_dir() / "ato-mcp"
    root.mkdir(parents=True, exist_ok=True)
    return root


def _bin_path() -> Path:
    return _ato_root() / "bin" / "ato-mcp.exe"


def _data_root() -> Path:
    return _ato_root() / "data"


def _release_tag_sentinel() -> Path:
    return _ato_root() / "release.tag"


def _ensure_binary() -> Path:
    """Make sure ato-mcp.exe is on disk. Idempotent after first call."""
    bin_path = _bin_path()
    if bin_path.is_file():
        return bin_path

    print(
        f"[office-tools] ato-mcp.exe not yet installed; fetching latest "
        f"release from github.com/{REPO_OWNER}/{REPO_NAME}...",
        file=sys.stderr,
    )

    release = github_release.latest_release(REPO_OWNER, REPO_NAME)
    tag = release.get("tag_name", "unknown")
    url = github_release.asset_url(release, WINDOWS_ASSET)
    if not url:
        raise RuntimeError(
            f"Release {tag} does not include {WINDOWS_ASSET}. Available assets: "
            f"{[a.get('name') for a in release.get('assets', [])]}."
        )

    bin_dir = bin_path.parent
    bin_dir.mkdir(parents=True, exist_ok=True)
    zip_path = bin_dir / WINDOWS_ASSET

    print(f"[office-tools] downloading {url} -> {zip_path}", file=sys.stderr)
    github_release.download_to(url, zip_path)

    print(f"[office-tools] extracting {zip_path}", file=sys.stderr)
    with zipfile.ZipFile(zip_path) as zf:
        zf.extractall(bin_dir)
    zip_path.unlink(missing_ok=True)

    if not bin_path.is_file():
        # Some zips put files in a top-level directory. Flatten one level if needed.
        for candidate in bin_dir.rglob("ato-mcp.exe"):
            candidate.rename(bin_path)
            break
        for candidate in bin_dir.rglob("onnxruntime.dll"):
            target = bin_dir / "onnxruntime.dll"
            if candidate != target:
                candidate.rename(target)
            break

    if not bin_path.is_file():
        raise RuntimeError(
            f"Extraction completed but {bin_path} is still missing. "
            f"Inspect {bin_dir} for the layout."
        )

    _release_tag_sentinel().write_text(tag)
    print(
        f"[office-tools] installed ato-mcp {tag}. The binary is unsigned; on "
        f"Windows the EDR sandbox typically holds it for ~20 minutes before "
        f"first execution succeeds. This is normal.",
        file=sys.stderr,
    )
    return bin_path


def _exec_serve(bin_path: Path) -> int:
    """Hand off to ato-mcp.exe serve, replacing the Python process."""
    data_dir = _data_root()
    data_dir.mkdir(parents=True, exist_ok=True)

    env = os.environ.copy()
    env["ATO_MCP_DATA_DIR"] = str(data_dir)

    forwarded_args = sys.argv[1:]
    if not forwarded_args:
        forwarded_args = ["serve"]

    if sys.platform == "win32":
        # os.execvpe doesn't behave consistently with .exe on Windows; use
        # subprocess and forward stdio so the MCP host's stdio remains hooked.
        import subprocess

        completed = subprocess.run([str(bin_path), *forwarded_args], env=env, check=False)
        return completed.returncode

    # Non-Windows path (developer testing only — production target is Windows).
    os.execvpe(str(bin_path), [str(bin_path), *forwarded_args], env)
    return 0  # unreachable


def main() -> int:
    bin_path = _ensure_binary()
    return _exec_serve(bin_path)


if __name__ == "__main__":
    sys.exit(main())
