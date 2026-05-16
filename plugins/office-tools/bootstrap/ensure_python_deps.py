"""Idempotently install the plugin's Python deps into the persistent data dir.

Called by every tool's entrypoint before any third-party import. The first call
(or any call after requirements.txt changes) runs pip install. Subsequent calls
are a hash comparison and an immediate return.

Deps land in CLAUDE_PLUGIN_DATA/site-packages/ so they survive plugin version
bumps. The plugin install directory (CLAUDE_PLUGIN_ROOT) is treated as
read-only; we never write to wheels/, python-base/Lib/site-packages/, etc.
"""

from __future__ import annotations

import hashlib
import os
import subprocess
import sys
from pathlib import Path

from . import paths


def _requirements_hash() -> str:
    return hashlib.sha256(paths.requirements_file().read_bytes()).hexdigest()


def _hash_sentinel() -> Path:
    return paths.site_packages_dir() / ".requirements.sha256"


def _python_executable() -> str:
    """The interpreter we should invoke pip with.

    Prefer the currently running interpreter (sys.executable). If that's blank
    (rare; happens with frozen / embedded interpreters that omit it), fall back
    to a python.exe alongside our plugin's python-base.
    """
    if sys.executable:
        return sys.executable
    candidate = paths.plugin_root() / "python-base" / "python.exe"
    if candidate.is_file():
        return str(candidate)
    raise RuntimeError("Cannot determine Python interpreter to invoke pip with.")


def _pip_install() -> None:
    target = paths.site_packages_dir()
    wheels = paths.wheels_dir()
    requirements = paths.requirements_file()

    cmd = [
        _python_executable(),
        "-m",
        "bootstrap.run_pip",
        "install",
        "--target",
        str(target),
        "--no-index",
        "--find-links",
        str(wheels),
        "--upgrade",
        "-r",
        str(requirements),
    ]

    env = os.environ.copy()
    # Ensure the child process can find the bootstrap package.
    pythonpath_parts = [str(paths.plugin_root())]
    if env.get("PYTHONPATH"):
        pythonpath_parts.append(env["PYTHONPATH"])
    env["PYTHONPATH"] = os.pathsep.join(pythonpath_parts)

    print(
        f"[office-tools] Installing Python dependencies into {target} (one-time)...",
        file=sys.stderr,
    )
    completed = subprocess.run(cmd, env=env, check=False)
    if completed.returncode != 0:
        raise RuntimeError(
            f"pip install failed with exit code {completed.returncode}. "
            "Re-run install.bat or report the failure."
        )


def run() -> None:
    """Ensure third-party deps are importable; install lazily if needed.

    Idempotent and cheap on the fast path. Adds the site-packages directory to
    sys.path either way.
    """
    target = paths.site_packages_dir()
    sentinel = _hash_sentinel()
    current_hash = _requirements_hash()

    needs_install = True
    if sentinel.is_file():
        try:
            if sentinel.read_text().strip() == current_hash:
                needs_install = False
        except OSError:
            needs_install = True

    if needs_install:
        _pip_install()
        sentinel.write_text(current_hash)

    target_str = str(target)
    if target_str not in sys.path:
        sys.path.insert(0, target_str)


if __name__ == "__main__":
    run()
