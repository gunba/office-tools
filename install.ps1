# office-tools installer (Windows, PowerShell)
#
# Downloads the latest office-tools release binary, drops it under
# %LOCALAPPDATA%\Temp\office-tools\, then registers the plugin marketplace
# with whichever of Claude Code or Codex CLI is on PATH.
#
# Why %LOCALAPPDATA%\Temp (a.k.a. %TEMP%): on a typical corporate Windows
# image, Defender Attack Surface Reduction rule "Block executable files...
# not from a trusted list" refuses exec of unsigned binaries from most of
# the user profile (Desktop, Documents, .cargo\bin, AppData\Local\Programs,
# etc.) but allows exec from %TEMP%. Installing here keeps the default
# install runnable on enterprise endpoints; pass -InstallDir to override
# for unrestricted machines.
#
# Usage (run from the repo root):
#   powershell -ExecutionPolicy Bypass -File .\install.ps1
#
# Optional flags:
#   -Tag <tag>           Install a specific tag (e.g. v0.1.0). Default: latest.
#   -InstallDir <path>   Override binary install location.
#                        Default: $env:LOCALAPPDATA\Temp\office-tools.
#   -SkipRegister        Skip claude/codex marketplace registration.

[CmdletBinding()]
param(
    [string]$Tag = 'latest',
    [string]$InstallDir = (Join-Path $env:LOCALAPPDATA 'Temp\office-tools'),
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

# Newer releases ship the binary directly as office-tools-windows-x64.exe.
# v0.1.0 packaged it inside office-tools-windows-x64.zip. Accept either.
$exeAsset = $release.assets | Where-Object { $_.name -eq 'office-tools-windows-x64.exe' } | Select-Object -First 1
$zipAsset = $release.assets | Where-Object { $_.name -eq 'office-tools-windows-x64.zip' } | Select-Object -First 1
if (-not $exeAsset -and -not $zipAsset) {
    throw "Release $($release.tag_name) does not have an office-tools-windows-x64.exe or .zip asset."
}
$asset = if ($exeAsset) { $exeAsset } else { $zipAsset }
Write-Host "  -> $($release.tag_name) / $($asset.name) ($([math]::Round($asset.size/1MB,2)) MB)"

# --- Download + place --------------------------------------------------------
New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
if ($asset.name -like '*.zip') {
    $tmpZip = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), [Guid]::NewGuid().ToString('N') + '.zip')
    try {
        Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $tmpZip -UseBasicParsing
        Expand-Archive -Path $tmpZip -DestinationPath $InstallDir -Force
    } finally {
        if (Test-Path $tmpZip) { Remove-Item $tmpZip -Force -ErrorAction SilentlyContinue }
    }
} else {
    $exePath = Join-Path $InstallDir 'office-tools.exe'
    Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $exePath -UseBasicParsing
}

# Strip Mark-of-the-Web so SmartScreen/AppLocker doesn't reject the file.
Get-ChildItem $InstallDir -File -Recurse | ForEach-Object { Unblock-File -LiteralPath $_.FullName -ErrorAction SilentlyContinue }

# --- Verify exec -------------------------------------------------------------
$exe = Join-Path $InstallDir 'office-tools.exe'
if (-not (Test-Path $exe)) { throw "office-tools.exe not found at $exe after install." }
Write-Host ""
Write-Host "Verifying binary..."
try {
    $version = & $exe --version 2>&1
    Write-Host "  $version"
} catch {
    Write-Warning "Could not exec $exe : $($_.Exception.Message)"
    Write-Warning ""
    Write-Warning "This often means the host's Defender Attack Surface Reduction (ASR) rule"
    Write-Warning "'Block executable files... not from a trusted list' is refusing the new"
    Write-Warning "binary because it has no Microsoft Cloud reputation yet. Options:"
    Write-Warning "  - Wait a day or so for Microsoft's cloud reputation to clear it, then retry."
    Write-Warning "  - Ask IT to add a Defender ASR exclusion for $InstallDir, or to add the"
    Write-Warning "    binary's SHA-256 to the trusted hash list."
    Write-Warning "  - Build from source: cargo build --release."
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
