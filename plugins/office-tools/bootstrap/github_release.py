"""Minimal anonymous GitHub Releases API client.

Stdlib-only — runs against the embedded Python before pip has installed
anything. Handles TLS-inspecting corporate proxies by falling back to
unverified SSL when the verified attempt fails. The same trade-off rdtaxtools
makes today at gitsync.py:14-19 — necessary for the firm's network, dangerous
in general, OK here because the URLs we hit are public GitHub release assets
whose contents we verify by SHA elsewhere if needed.
"""

from __future__ import annotations

import json
import ssl
import sys
import urllib.error
import urllib.request
from pathlib import Path


def _opener(insecure: bool = False) -> urllib.request.OpenerDirector:
    if insecure:
        ctx = ssl._create_unverified_context()
    else:
        ctx = ssl.create_default_context()
    handler = urllib.request.HTTPSHandler(context=ctx)
    opener = urllib.request.build_opener(handler)
    opener.addheaders = [
        ("User-Agent", "office-tools-plugin"),
        ("Accept", "application/vnd.github+json"),
    ]
    return opener


def _request(url: str, *, allow_insecure: bool = True) -> bytes:
    """GET a URL with retry under unverified SSL on certificate failure."""
    try:
        with _opener(insecure=False).open(url, timeout=60) as resp:
            return resp.read()
    except urllib.error.URLError as e:
        # Many corporate proxies break the TLS chain. Retry without verification.
        if allow_insecure and isinstance(getattr(e, "reason", None), ssl.SSLError):
            print(
                f"[office-tools] SSL verify failed for {url}, retrying without verification "
                f"(corporate proxy interception). Reason: {e.reason}",
                file=sys.stderr,
            )
            with _opener(insecure=True).open(url, timeout=60) as resp:
                return resp.read()
        raise


def latest_release(owner: str, repo: str) -> dict:
    """Fetch the latest release metadata for a public repo (no auth)."""
    url = f"https://api.github.com/repos/{owner}/{repo}/releases/latest"
    data = _request(url)
    return json.loads(data)


def asset_url(release: dict, asset_name: str) -> str | None:
    """Return the browser_download_url for a named asset, or None."""
    for asset in release.get("assets", []):
        if asset.get("name") == asset_name:
            return asset.get("browser_download_url")
    return None


def download_to(url: str, dest: Path) -> None:
    """Stream a URL to a file path. Atomic via temp + rename."""
    dest.parent.mkdir(parents=True, exist_ok=True)
    tmp = dest.with_suffix(dest.suffix + ".part")
    try:
        # Stream rather than read all into memory; release assets can be hundreds of MB.
        try:
            req = _opener(insecure=False).open(url, timeout=300)
        except urllib.error.URLError as e:
            if isinstance(getattr(e, "reason", None), ssl.SSLError):
                print(
                    f"[office-tools] SSL verify failed for {url}, retrying insecure.",
                    file=sys.stderr,
                )
                req = _opener(insecure=True).open(url, timeout=300)
            else:
                raise
        with req as resp, tmp.open("wb") as out:
            while True:
                chunk = resp.read(1 << 20)  # 1 MiB
                if not chunk:
                    break
                out.write(chunk)
        tmp.replace(dest)
    except BaseException:
        tmp.unlink(missing_ok=True)
        raise
