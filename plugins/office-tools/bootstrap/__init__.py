"""office-tools bootstrap package.

Each tool's main script calls `bootstrap.ensure_python_deps.run()` before
importing any third-party library. The function lazily pip-installs all
declared dependencies into the plugin's persistent data dir on first run,
then noops on subsequent runs.
"""
