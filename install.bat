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

where cargo >nul 2>&1
if %errorlevel% neq 0 (
    echo Rust Cargo was not found on PATH.
    echo Install Rust from https://rustup.rs/ or place a prebuilt office-tools.exe at:
    echo   %REPO_DIR%\plugins\office-tools\bin\office-tools.exe
    exit /b 1
)

echo Building Rust office-tools binary...
cargo build --release --manifest-path "%REPO_DIR%\Cargo.toml"
if %errorlevel% neq 0 (
    echo Rust build failed.
    exit /b 1
)

if not exist "%REPO_DIR%\plugins\office-tools\bin" mkdir "%REPO_DIR%\plugins\office-tools\bin"
copy /Y "%REPO_DIR%\target\release\office-tools.exe" "%REPO_DIR%\plugins\office-tools\bin\office-tools.exe" >nul
if %errorlevel% neq 0 (
    echo Could not copy office-tools.exe into the plugin bin directory.
    exit /b 1
)
echo   Built: plugins\office-tools\bin\office-tools.exe
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
echo The plugin now uses the Rust binary above. No bundled Python runtime,
echo Python wheels, openpyxl, xlwings, python-docx, or python-pptx are used.
echo WinCOM commands still require Microsoft Office on Windows.
echo ============================================================

endlocal
