# office-tools installer (Windows, PowerShell)
#
# Downloads the latest office-tools release binary, drops it under
# %LOCALAPPDATA%\office-tools\, then registers the plugin marketplace
# with whichever of Claude Code or Codex CLI is on PATH.
#
# Usage (run from the repo root):
#   powershell -ExecutionPolicy Bypass -File .\install.ps1
#
# Optional flags:
#   -Tag <tag>           Install a specific tag (e.g. v0.1.0). Default: latest.
#   -InstallDir <path>   Override binary install location.
#                        Default: $env:LOCALAPPDATA\office-tools.
#   -SkipRegister        Skip claude/codex marketplace registration.

[CmdletBinding()]
param(
    [string]$Tag = 'latest',
    [string]$InstallDir = (Join-Path $env:LOCALAPPDATA 'office-tools'),
    [switch]$SkipRegister
)

$ErrorActionPreference = 'Stop'
$repo = 'gunba/office-tools'
$marketplaceRoot = $PSScriptRoot

Write-Host "office-tools installer"
Write-Host "  marketplace: $marketplaceRoot"
Write-Host "  install dir: $InstallDir"
Write-Host ""

# --- Resolve release ---------------------------------------------------------
$apiUrl = if ($Tag -eq 'latest') {
    "https://api.github.com/repos/$repo/releases/latest"
} else {
    "https://api.github.com/repos/$repo/releases/tags/$Tag"
}
Write-Host "Fetching release metadata from $apiUrl"
$release = Invoke-RestMethod -Uri $apiUrl -UseBasicParsing
$asset = $release.assets | Where-Object { $_.name -eq 'office-tools-windows-x64.zip' } | Select-Object -First 1
if (-not $asset) {
    throw "Release $($release.tag_name) does not have an office-tools-windows-x64.zip asset."
}
Write-Host "  -> $($release.tag_name) / $($asset.name) ($([math]::Round($asset.size/1MB,2)) MB)"

# --- Download + extract ------------------------------------------------------
$tmpZip = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), [Guid]::NewGuid().ToString('N') + '.zip')
try {
    Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $tmpZip -UseBasicParsing
    New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
    Expand-Archive -Path $tmpZip -DestinationPath $InstallDir -Force
} finally {
    if (Test-Path $tmpZip) { Remove-Item $tmpZip -Force -ErrorAction SilentlyContinue }
}

# --- Verify exec -------------------------------------------------------------
$exe = Join-Path $InstallDir 'office-tools.exe'
if (-not (Test-Path $exe)) { throw "office-tools.exe not found at $exe after extract." }
Write-Host ""
Write-Host "Verifying binary..."
try {
    $version = & $exe --version 2>&1
    Write-Host "  $version"
} catch {
    Write-Warning "Could not exec $exe : $($_.Exception.Message)"
    Write-Warning "If your endpoint policy blocks exec from $InstallDir, install to a permitted path with -InstallDir <path> and update plugins/office-tools/.mcp.json accordingly."
    throw
}

# --- Register marketplaces ---------------------------------------------------
if ($SkipRegister) {
    Write-Host ""
    Write-Host "Skipping marketplace registration (-SkipRegister)."
} else {
    Write-Host ""
    if (Get-Command claude -ErrorAction SilentlyContinue) {
        Write-Host "Registering with Claude Code..."
        & claude plugin marketplace add $marketplaceRoot
        & claude plugin install office-tools@office-tools
    } else {
        Write-Host "claude not on PATH; skipping Claude registration."
    }
    Write-Host ""
    if (Get-Command codex -ErrorAction SilentlyContinue) {
        Write-Host "Registering with Codex..."
        & codex plugin marketplace add $marketplaceRoot
        Write-Host "  (codex does not auto-install plugins from marketplaces — add an [mcp_servers.office-tools] block to ~/.codex/config.toml. See README.)"
    } else {
        Write-Host "codex not on PATH; skipping Codex registration."
    }
}

Write-Host ""
Write-Host "============================================================"
Write-Host "Install complete."
Write-Host "  binary:      $exe"
Write-Host "  marketplace: $marketplaceRoot"
Write-Host "============================================================"
