# office-tools uninstaller (Windows, PowerShell)
#
# Reverses install.ps1: removes the binary directory and unregisters
# the plugin from Claude Code and Codex if they are on PATH.

[CmdletBinding()]
param(
    [string]$InstallDir = (Join-Path $env:LOCALAPPDATA 'Temp\office-tools')
)

$ErrorActionPreference = 'Stop'

Write-Host "office-tools uninstaller"
Write-Host ""

if (Get-Command claude -ErrorAction SilentlyContinue) {
    Write-Host "Removing from Claude Code..."
    & claude plugin uninstall office-tools 2>&1 | Out-Null
    & claude plugin marketplace remove office-tools 2>&1 | Out-Null
    Write-Host "  done."
} else {
    Write-Host "claude not on PATH; skipping."
}

Write-Host ""
if (Get-Command codex -ErrorAction SilentlyContinue) {
    Write-Host "Removing from Codex..."
    & codex plugin marketplace remove office-tools 2>&1 | Out-Null
    Write-Host "  done (remember to remove any [mcp_servers.office-tools] block from ~/.codex/config.toml)."
} else {
    Write-Host "codex not on PATH; skipping."
}

Write-Host ""
if (Test-Path $InstallDir) {
    Write-Host "Removing $InstallDir ..."
    Remove-Item $InstallDir -Recurse -Force
    Write-Host "  done."
} else {
    Write-Host "$InstallDir does not exist; nothing to remove."
}

Write-Host ""
Write-Host "Uninstall complete."
