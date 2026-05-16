@echo off
setlocal enabledelayedexpansion

REM office-tools uninstaller
REM Reverses what install.bat did. Each CLI removal is independent.

echo.
echo office-tools uninstaller
echo.

REM ---- Claude Code ---------------------------------------------------------
where claude >nul 2>&1
if %errorlevel% equ 0 (
    echo Removing from Claude Code...
    claude plugin uninstall office-tools@office-tools 2>nul
    claude plugin marketplace remove office-tools 2>nul
    echo   Claude Code: uninstalled.
    echo.
) else (
    echo Claude Code CLI not found. Skipping.
    echo.
)

REM ---- Codex ---------------------------------------------------------------
where codex >nul 2>&1
if %errorlevel% equ 0 (
    echo Removing from Codex...
    codex plugin remove office-tools@office-tools 2>nul
    codex plugin marketplace remove office-tools 2>nul
    echo   Codex: uninstalled.
    echo.
) else (
    echo Codex CLI not found. Skipping.
    echo.
)

echo Uninstall complete. Persistent plugin data (cached corpus, pip-installed
echo wheels) was removed by the CLI's plugin uninstall step.

endlocal
