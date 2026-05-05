@echo off
chcp 65001 >nul 2>&1
setlocal enabledelayedexpansion
title AlFred.mcp - Installer

set "REPO_URL=https://github.com/h22fred/Alfred.mcp.git"

echo.
echo   What is your role?
echo.
echo     1) SC / SSC / Manager           - Solution Consulting
echo     2) Sales / Specialist / Manager - Account Executive
echo.
set /p VARIANT_CHOICE="   Enter 1 or 2 (default: 1): "

if "!VARIANT_CHOICE!"=="2" (
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
echo   AlFred.mcp - Installer
echo ==================================================
echo.

:: ------------------------------------------------------------
:: 1. Check / install Git
:: ------------------------------------------------------------
where git >nul 2>&1
if !ERRORLEVEL! neq 0 (
    echo   Git not found - installing automatically...
    echo.

    rem Try winget first (built into Windows 10 1709+ and Windows 11)
    where winget >nul 2>&1
    if !ERRORLEVEL! equ 0 (
        echo   Installing Git via winget...
        winget install --id Git.Git -e --source winget --accept-package-agreements --accept-source-agreements
    ) else (
        rem Fallback: download Git installer directly via curl (built-in Windows 10+)
        echo   Downloading Git for Windows...
        set "GIT_INSTALLER=%TEMP%\git-installer.exe"
        curl -fL -o "!GIT_INSTALLER!" "https://github.com/git-for-windows/git/releases/latest/download/Git-2.49.0-64-bit.exe"
        if not exist "!GIT_INSTALLER!" (
            echo   Download failed. Please install Git manually from:
            echo   https://git-scm.com/download/win
            echo   Then re-run this script.
            pause
            exit /b 1
        )
        echo   Installing Git (this may take a minute)...
        "!GIT_INSTALLER!" /VERYSILENT /NORESTART /NOCANCEL /SP- /CLOSEAPPLICATIONS /RESTARTAPPLICATIONS /COMPONENTS="icons,ext\reg\shellhere,assoc,assoc_sh"
        del "!GIT_INSTALLER!" >nul 2>&1
    )

    rem Refresh PATH from registry so git is available in this session
    for /f "tokens=2*" %%a in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v Path 2^>nul') do set "SYS_PATH=%%b"
    for /f "tokens=2*" %%a in ('reg query "HKCU\Environment" /v Path 2^>nul') do set "USR_PATH=%%b"
    set "PATH=!SYS_PATH!;!USR_PATH!"

    where git >nul 2>&1
    if !ERRORLEVEL! neq 0 (
        echo.
        echo   Git installation finished but git is not in PATH yet.
        echo   Please close this window and re-run Setup_Windows.bat.
        echo.
        pause
        exit /b 1
    )
    echo   Git installed successfully
)

:: ------------------------------------------------------------
:: 2. Clone or update the repo
:: ------------------------------------------------------------
if exist "!INSTALL_DIR!\.git" (
    echo   Updating existing installation...
    git -C "!INSTALL_DIR!" fetch origin 2>&1
    git -C "!INSTALL_DIR!" reset --hard origin/main 2>&1
    if !ERRORLEVEL! neq 0 (
        echo   Update failed - doing fresh install...
        rmdir /s /q "!INSTALL_DIR!" 2>nul
        git clone "!REPO_URL!" "!INSTALL_DIR!"
    ) else (
        echo   Updated to latest
    )
) else (
    echo   Cloning alfred.mcp...
    git clone "!REPO_URL!" "!INSTALL_DIR!"
)
if not exist "!INSTALL_DIR!\setup\setup.ps1" (
    echo.
    echo   ERROR: Installation failed - setup.ps1 not found.
    echo   Please check your internet connection and try again.
    echo.
    pause
    exit /b 1
)

:: ------------------------------------------------------------
:: 3. Run setup
:: ------------------------------------------------------------
echo.
powershell -ExecutionPolicy Bypass -File "!INSTALL_DIR!\setup\setup.ps1"

echo.
pause
endlocal
\r