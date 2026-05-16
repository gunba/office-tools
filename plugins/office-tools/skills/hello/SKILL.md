---
description: Verify the office-tools plugin install is healthy. Confirms the embedded Python is reachable, the bootstrap can locate the plugin root and data dir, and the dependency wheels are present.
---

# office-tools:hello

Diagnostic skill for the office-tools plugin. Run this after install to confirm the plugin is wired up correctly.

## What to do

Run this command exactly:

```bash
"${CLAUDE_PLUGIN_ROOT}/python-base/python.exe" -c "import sys, os; sys.path.insert(0, os.environ['CLAUDE_PLUGIN_ROOT']); from bootstrap import paths; print('plugin_root  =', paths.plugin_root()); print('plugin_data  =', paths.plugin_data_dir()); print('wheels found =', len(list(paths.wheels_dir().glob('*.whl'))))"
```

A healthy install reports an absolute plugin_root under the CLI's plugin cache, an absolute plugin_data under the CLI's plugin DATA dir, and at least 10 wheels found. Anything else means the plugin is mis-wired:

- "wheels found = 0" → `wheels/` is missing; re-run install.bat.
- ModuleNotFoundError for `bootstrap` → `CLAUDE_PLUGIN_ROOT` did not expand. Make sure the plugin is registered with `claude plugin marketplace add` and enabled.
- python.exe not found → `python-base/` is missing; the install was incomplete.
