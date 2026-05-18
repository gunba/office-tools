---
description: Check whether a newer office-tools release is available and recommend the upgrade step.
---

Run the following in a shell and report the result to the user:

```powershell
$exe = Join-Path $env:LOCALAPPDATA 'Temp\office-tools\office-tools.exe'
if (-not (Test-Path $exe)) {
    Write-Host "office-tools is not installed at $exe."
    Write-Host "Install with: powershell -ExecutionPolicy Bypass -File <marketplace>\install.ps1"
    return
}

$current = ((& $exe --version 2>&1) -replace 'office-tools\s+', '').Trim()
try {
    $latest = (Invoke-RestMethod 'https://api.github.com/repos/gunba/office-tools/releases/latest').tag_name.TrimStart('v')
} catch {
    Write-Host "Could not reach api.github.com to look up the latest release: $($_.Exception.Message)"
    return
}

Write-Host "Current: $current"
Write-Host "Latest:  $latest"
Write-Host ""
if ([version]$current -ge [version]$latest) {
    Write-Host "office-tools is up to date."
} else {
    Write-Host "An update is available. To upgrade, re-run install.ps1 from your office-tools marketplace directory:"
    Write-Host "  powershell -ExecutionPolicy Bypass -File <marketplace>\install.ps1 -Tag v$latest"
    Write-Host ""
    Write-Host "After install, restart Claude so the MCP server picks up the new binary."
}
```

Notes for the user if the install.ps1 step fails:
- A freshly-released binary may be temporarily blocked by Defender ASR ("Block executable files... not from a trusted list") until Microsoft's cloud reputation accumulates. Wait or ask IT for a path/SHA exclusion.
- The installer respects `-InstallDir <path>` if the default path is rejected by endpoint security.
