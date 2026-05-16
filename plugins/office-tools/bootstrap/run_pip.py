"""Invoke pip from a wheel via zipimport, with the platform.uname() workaround.

Background (from rdtaxtools/autorego/autorego/pip_install.py):
On corporate Windows endpoints where the embedded Python is unsigned, pip's
internal ``user_agent()`` helper calls ``platform.system()``, which dispatches
to ``platform.uname()`` -> ``platform.win32_ver()``. That last call reads
``HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion`` from the registry,
and the endpoint security agent stalls that read indefinitely, making
``pip install`` hang at "Verifying project dependencies". Pre-populating
``platform._uname_cache`` short-circuits ``uname()`` so ``win32_ver()`` is
never called and pip runs normally.

Why zipimport instead of running pip as a subprocess: the embedded Python
distribution doesn't ship pip, and we don't want to install it into the plugin
root (read-only, replaced on plugin upgrade). Treating the pip wheel as a
sys.path entry lets Python import pip directly from
``wheels/pip-*.whl`` without any prior install step.
"""

from __future__ import annotations

import glob
import os
import platform
import sys

if sys.platform == "win32":
    platform._uname_cache = platform.uname_result(  # type: ignore[attr-defined]
        system="Windows",
        node=os.environ.get("COMPUTERNAME", ""),
        release="",
        version="",
        machine=os.environ.get("PROCESSOR_ARCHITECTURE", "AMD64"),
    )


def _add_pip_wheel_to_path() -> None:
    from . import paths

    candidates = sorted(glob.glob(str(paths.wheels_dir() / "pip-*.whl")))
    if not candidates:
        raise RuntimeError(
            f"No pip wheel found in {paths.wheels_dir()}. The plugin install is "
            "incomplete — re-run install.bat."
        )
    pip_wheel = candidates[-1]
    if pip_wheel not in sys.path:
        sys.path.insert(0, pip_wheel)


def main(argv: list[str] | None = None) -> int:
    _add_pip_wheel_to_path()
    from pip._internal.cli.main import main as pip_main

    return pip_main(list(argv) if argv is not None else sys.argv[1:])


if __name__ == "__main__":
    sys.exit(main())
