# ============================================================
# AlFred.mcp - Windows Bootstrap
# Downloaded and run by Setup_Windows.bat via iex
# Handles: role selection, git install, repo clone, then
# launches setup\setup.ps1 from the cloned repo.
# ============================================================
$ErrorActionPreference = "Stop"

$RepoUrl  = "https://github.com/h22fred/Alfred.mcp.git"

Write-Host ""
Write-Host "  What is your role?"
Write-Host ""
Write-Host "    1) SC / SSC / Manager           - Solution Consulting"
Write-Host "    2) Sales / Specialist / Manager - Account Executive"
Write-Host ""
$Choice = Read-Host "   Enter 1 or 2 (default: 1)"
if ($Choice -eq "2") {
    $env:ALFRED_VARIANT = "sales"
    $InstallDir = "$env:USERPROFILE\Documents\alfred.sales"
    Write-Host "    Installing Alfred Sales to $InstallDir"
} else {
    $env:ALFRED_VARIANT = "sc"
    $InstallDir = "$env:USERPROFILE\Documents\alfred.sc"
    Write-Host "    Installing Alfred SC to $InstallDir"
}

Write-Host ""
Write-Host "=================================================="
Write-Host "  AlFred.mcp - Installer"
Write-Host "=================================================="
Write-Host ""

# ── 1. Check / install Git ────────────────────────────────
$GitOk = $false
try { & git --version 2>$null | Out-Null; $GitOk = $true } catch {}

if (-not $GitOk) {
    Write-Host "  Git not found - installing automatically..."
    Write-Host ""

    $WingetOk = $false
    try { & winget --version 2>$null | Out-Null; $WingetOk = $true } catch {}

    $GitInstalledViaWinget = $false
    if ($WingetOk) {
        Write-Host "  Installing Git via winget..."
        & winget install --id Git.Git -e --source winget --accept-package-agreements --accept-source-agreements
        if ($LASTEXITCODE -eq 0) { $GitInstalledViaWinget = $true }
        else { Write-Host "  Winget failed (code $LASTEXITCODE) - falling back to direct installer..." }
    }

    if (-not $GitInstalledViaWinget) {
        Write-Host "  Downloading Git installer..."
        $GitInstaller = "$env:TEMP\git-installer.exe"
        try {
            (New-Object Net.WebClient).DownloadFile(
                "https://github.com/git-for-windows/git/releases/latest/download/Git-2.49.0-64-bit.exe",
                $GitInstaller
            )
        } catch {
            Write-Host "  Download failed: $_"
            Write-Host "  Please install Git manually from https://git-scm.com/download/win then re-run."
            Read-Host "  Press Enter to exit"; exit 1
        }
        if (-not (Test-Path $GitInstaller)) {
            Write-Host "  Download failed. Install Git from https://git-scm.com/download/win then re-run."
            Read-Host "  Press Enter to exit"; exit 1
        }
        Write-Host "  Installing Git (this may take a minute)..."
        & $GitInstaller /VERYSILENT /NORESTART /NOCANCEL /SP- /CLOSEAPPLICATIONS /RESTARTAPPLICATIONS
        Remove-Item $GitInstaller -ErrorAction SilentlyContinue
    }

    # Refresh PATH so git is usable in this session
    $env:PATH = [System.Environment]::GetEnvironmentVariable("PATH", "Machine") + ";" +
                [System.Environment]::GetEnvironmentVariable("PATH", "User")

    try { & git --version 2>$null | Out-Null; $GitOk = $true } catch {}
    if (-not $GitOk) {
        Write-Host ""
        Write-Host "  Git install completed but not detected in PATH."
        Write-Host "  Please close this window, open a new one, and re-run Setup_Windows.bat."
        Read-Host "  Press Enter to exit"; exit 1
    }
    Write-Host "  Git installed successfully"
}

# ── 2. Clone or update ────────────────────────────────────
if (Test-Path (Join-Path $InstallDir ".git")) {
    Write-Host "  Updating existing installation..."
    & git -C $InstallDir fetch origin 2>&1 | Out-Null
    & git -C $InstallDir reset --hard origin/main 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) {
        Write-Host "  Updated to latest"
    } else {
        Write-Host "  Update failed - doing fresh install..."
        Remove-Item $InstallDir -Recurse -Force -ErrorAction SilentlyContinue
        & git clone $RepoUrl $InstallDir
    }
} else {
    Write-Host "  Cloning alfred.mcp..."
    & git clone $RepoUrl $InstallDir
}

# ── 3. Run setup ─────────────────────────────────────────
$SetupScript = Join-Path $InstallDir "setup\setup.ps1"
if (-not (Test-Path $SetupScript)) {
    Write-Host ""
    Write-Host "  ERROR: Installation failed - setup.ps1 not found."
    Write-Host "  Please check your internet connection and try again."
    Read-Host "  Press Enter to exit"; exit 1
}

Write-Host ""
& powershell -ExecutionPolicy Bypass -File $SetupScript

Write-Host ""
Read-Host "  Press Enter to exit"
