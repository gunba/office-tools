@echo off
setlocal enabledelayedexpansion

REM office-tools installer
REM Registers the marketplace at the current directory and installs the office-tools plugin
REM under Claude Code and Codex. Both registrations are independent; one failing doesn't stop the other.

set "REPO_DIR=%~dp0"
if "%REPO_DIR:~-1%"=="\" set "REPO_DIR=%REPO_DIR:~0,-1%"

echo.
echo office-tools installer
echo Marketplace path: %REPO_DIR%
echo.

set "INSTALLED_ANY=0"

REM ---- Claude Code ---------------------------------------------------------
where claude >nul 2>&1
if %errorlevel% equ 0 (
    echo Registering with Claude Code...
    claude plugin marketplace add "%REPO_DIR%"
    if !errorlevel! equ 0 (
        claude plugin install office-tools@office-tools
        if !errorlevel! equ 0 (
            echo   Claude Code: registered + installed.
            set "INSTALLED_ANY=1"
        ) else (
            echo   Claude Code: marketplace added, but plugin install failed.
        )
    ) else (
        echo   Claude Code: marketplace add failed.
    )
    echo.
) else (
    echo Claude Code CLI not found on PATH. Skipping.
    echo.
)

REM ---- Codex ---------------------------------------------------------------
where codex >nul 2>&1
if %errorlevel% equ 0 (
    echo Registering with Codex...
    codex plugin marketplace add "%REPO_DIR%"
    if !errorlevel! equ 0 (
        codex plugin add office-tools@office-tools 2>nul
        if !errorlevel! equ 0 (
            echo   Codex: registered + installed.
            set "INSTALLED_ANY=1"
        ) else (
            echo   Codex: marketplace added, but `codex plugin add` is not available
            echo   in this version of Codex. Run `codex update` to upgrade, then re-run
            echo   this installer. The marketplace registration above is harmless and
            echo   can stay in place.
        )
    ) else (
        echo   Codex: marketplace add failed.
    )
    echo.
) else (
    echo Codex CLI not found on PATH. Skipping.
    echo.
)

if "%INSTALLED_ANY%"=="0" (
    echo No CLI installs succeeded. Confirm at least one of `claude` or `codex` is on PATH.
    exit /b 1
)

echo ============================================================
echo Install complete.
echo.
echo On first use of the `ato` MCP tools, the ato-mcp binary (~tens of MB)
echo and the ATO legal corpus (~4 GB) will download into the plugin's
echo persistent data directory. Windows Defender / CrowdStrike / SentinelOne
echo typically holds the unsigned ato-mcp binary in an EDR sandbox for
echo about 20 minutes before allowing execution. This is normal — wait it
echo out and the MCP server will start working without further action.
echo ============================================================

endlocal
