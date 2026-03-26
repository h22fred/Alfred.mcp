@echo off
title AlFred.mcp — Installer

set "REPO_URL=https://github.com/h22fred/Alfred.mcp.git"

echo.
echo   What is your role?
echo.
echo     1) SC / SSC / Manager  — Solution Consulting
echo     2) Sales               — Account Executive / Manager
echo.
set /p VARIANT_CHOICE="   Enter 1 or 2 (default: 1): "

if "%VARIANT_CHOICE%"=="2" (
    set "ALFRED_VARIANT=sales"
    set "INSTALL_DIR=%USERPROFILE%\Documents\alfred.sales"
    echo     Installing Alfred Sales to %USERPROFILE%\Documents\alfred.sales
) else (
    set "ALFRED_VARIANT=sc"
    set "INSTALL_DIR=%USERPROFILE%\Documents\alfred.sc"
    echo     Installing Alfred SC to %USERPROFILE%\Documents\alfred.sc
)

echo.
echo ==================================================
echo   AlFred.mcp — Installer
echo ==================================================
echo.

:: ------------------------------------------------------------
:: 1. Check Git
:: ------------------------------------------------------------
where git >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo   ERROR: Git is not installed or not in PATH.
    echo.
    echo   Please install Git for Windows from:
    echo   https://git-scm.com/download/win
    echo.
    echo   Then re-run this script.
    echo.
    pause
    exit /b 1
)

:: ------------------------------------------------------------
:: 2. Clone or update the repo
:: ------------------------------------------------------------
if exist "%INSTALL_DIR%\.git" (
    echo   Updating existing installation...
    git -C "%INSTALL_DIR%" fetch origin >nul 2>&1
    git -C "%INSTALL_DIR%" reset --hard origin/main >nul 2>&1
    echo   Updated to latest
) else (
    echo   Cloning alfred.mcp...
    git clone "%REPO_URL%" "%INSTALL_DIR%"
    echo   Cloned to %INSTALL_DIR%
)

:: ------------------------------------------------------------
:: 3. Run setup
:: ------------------------------------------------------------
echo.
powershell -ExecutionPolicy Bypass -File "%INSTALL_DIR%\setup.ps1"

echo.
pause
