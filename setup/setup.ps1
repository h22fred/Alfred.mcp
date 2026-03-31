# ============================================================
# SC Engagement MCP - Setup Script (Windows PowerShell)
# ============================================================
# Run this once to install everything and configure Claude Desktop.
# Requirements: Windows 10+, Google Chrome, Claude Desktop
# Called by: Setup_Windows.bat (sets $env:ALFRED_VARIANT)

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoDir = Split-Path -Parent $ScriptDir
$ConfigPath = "$env:USERPROFILE\.alfred-config.json"

# --- Helper: read / write config JSON ---
function Read-AlfredConfig {
    if (Test-Path $ConfigPath) {
        try {
            return Get-Content $ConfigPath -Raw | ConvertFrom-Json
        } catch {
            return [PSCustomObject]@{}
        }
    }
    return [PSCustomObject]@{}
}

function Write-AlfredConfig {
    param([PSCustomObject]$Config)
    $Config | ConvertTo-Json -Depth 5 | Set-Content $ConfigPath -Encoding UTF8
}

# --- Helper: ensure a property exists on the config object ---
function Ensure-Property {
    param([PSCustomObject]$Obj, [string]$Name, $Default = "")
    if (-not ($Obj.PSObject.Properties.Name -contains $Name)) {
        $Obj | Add-Member -NotePropertyName $Name -NotePropertyValue $Default
    }
}

# --- Helper: day name to DayOfWeek enum ---
function ConvertTo-DayOfWeek {
    param([string]$Day)
    switch ($Day.ToLower()) {
        "mon"       { return "Monday" }
        "monday"    { return "Monday" }
        "tue"       { return "Tuesday" }
        "tuesday"   { return "Tuesday" }
        "wed"       { return "Wednesday" }
        "wednesday" { return "Wednesday" }
        "thu"       { return "Thursday" }
        "thursday"  { return "Thursday" }
        "fri"       { return "Friday" }
        "friday"    { return "Friday" }
        "sat"       { return "Saturday" }
        "saturday"  { return "Saturday" }
        "sun"       { return "Sunday" }
        "sunday"    { return "Sunday" }
        default     { return $null }
    }
}

# --- Helper: day name to schtasks 3-letter abbreviation ---
function ConvertTo-SchtasksDay {
    param([string]$Day)
    switch ($Day) {
        "Monday"    { return "MON" }
        "Tuesday"   { return "TUE" }
        "Wednesday" { return "WED" }
        "Thursday"  { return "THU" }
        "Friday"    { return "FRI" }
        "Saturday"  { return "SAT" }
        "Sunday"    { return "SUN" }
        default     { return "MON" }
    }
}

Write-Host ""
Write-Host "=================================================="
Write-Host "  SC Engagement MCP - Setup"
Write-Host "=================================================="
Write-Host ""

# ------------------------------------------------------------
# 1. Check / install Node.js
# ------------------------------------------------------------
Write-Host "[>] Checking Node.js..."

$NodePath = $null
$SearchPaths = @(
    "C:\Program Files\nodejs\node.exe",
    "C:\Program Files (x86)\nodejs\node.exe"
)

foreach ($p in $SearchPaths) {
    if (Test-Path $p) {
        $NodePath = $p
        break
    }
}

if (-not $NodePath) {
    $cmd = Get-Command node -ErrorAction SilentlyContinue
    if ($cmd) {
        $NodePath = $cmd.Source
    }
}

if (-not $NodePath) {
    Write-Host ""
    Write-Host "   Node.js not found."
    Write-Host ""
    Write-Host "   Please install Node.js LTS from: https://nodejs.org"
    Write-Host "   Then re-run this script."
    Write-Host ""
    exit 1
}

$NodeDir = Split-Path -Parent $NodePath
Write-Host "   Node.js found at $NodePath"

# Ensure npm is on PATH for this session
if ($env:PATH -notlike "*$NodeDir*") {
    $env:PATH = "$NodeDir;$env:PATH"
}

# ------------------------------------------------------------
# 2. Install dependencies and build
# ------------------------------------------------------------
Write-Host ""
Write-Host "[>] Checking dependencies..."

$LockFile = Join-Path $RepoDir "package-lock.json"
$CurrentHash = (Get-FileHash $LockFile -Algorithm MD5).Hash

$Config = Read-AlfredConfig
$StoredHash = if ($Config.PSObject.Properties.Name -contains "lockSum") { $Config.lockSum } else { "" }
$NodeModulesLock = Join-Path $RepoDir "node_modules\.package-lock.json"

if (($CurrentHash -eq $StoredHash) -and (Test-Path $NodeModulesLock)) {
    Write-Host "   Dependencies up to date - skipping reinstall"
} else {
    Write-Host "   Installing dependencies..."
    & npm --prefix "$RepoDir" ci --no-fund
    if ($LASTEXITCODE -ne 0) { throw "npm ci failed with exit code $LASTEXITCODE" }

    $Config = Read-AlfredConfig
    Ensure-Property $Config "lockSum" $CurrentHash
    $Config.lockSum = $CurrentHash
    Write-AlfredConfig $Config
    Write-Host "   Dependencies installed"
}

Write-Host ""
Write-Host "[>] Building MCP server..."
& npm --prefix "$RepoDir" run build
if ($LASTEXITCODE -ne 0) { throw "npm run build failed with exit code $LASTEXITCODE" }
Write-Host "   Build complete"

# Save installed version SHA for update checks
try {
    $InstalledSHA = & git -C $RepoDir rev-parse --short HEAD 2>$null
    if ($InstalledSHA) {
        $Config = Read-AlfredConfig
        Ensure-Property $Config "installedVersion" ""
        $Config.installedVersion = $InstalledSHA
        Write-AlfredConfig $Config
        Write-Host "   Installed version: $InstalledSHA"
    }
} catch {
    # Not a git repo or git not available - skip
}

# ------------------------------------------------------------
# 3. Dynamics URL
# ------------------------------------------------------------
Write-Host ""
Write-Host "[>] Dynamics 365 instance..."

$Config = Read-AlfredConfig
$ExistingDynamicsUrl = if ($Config.PSObject.Properties.Name -contains "dynamicsUrl") { $Config.dynamicsUrl } else { "" }

if ($ExistingDynamicsUrl) {
    Write-Host "   Dynamics URL already set: $ExistingDynamicsUrl"
    $NewCompany = Read-Host "   Change company name? (press Enter to keep)"
    if ($NewCompany) {
        $NewDynamicsUrl = "https://$NewCompany.crm.dynamics.com"
    } else {
        $NewDynamicsUrl = $ExistingDynamicsUrl
    }
} else {
    $NewCompany = Read-Host "   What is your company name? (press Enter for 'servicenow')"
    if (-not $NewCompany) { $NewCompany = "servicenow" }
    $NewDynamicsUrl = "https://$NewCompany.crm.dynamics.com"
}

$Config = Read-AlfredConfig
Ensure-Property $Config "dynamicsUrl" ""
$Config.dynamicsUrl = $NewDynamicsUrl
Write-AlfredConfig $Config
Write-Host "   Dynamics URL set to: $NewDynamicsUrl"

# ------------------------------------------------------------
# 4. Configure Claude Desktop
# ------------------------------------------------------------
Write-Host ""
Write-Host "[>] Configuring Claude Desktop..."

$Variant = if ($env:ALFRED_VARIANT) { $env:ALFRED_VARIANT } else { "sc" }
$DistPath = Join-Path $RepoDir "dist\$Variant\index.js"
$ClaudeConfigDir = Join-Path $env:APPDATA "Claude"
$ClaudeConfigPath = Join-Path $ClaudeConfigDir "claude_desktop_config.json"

if (-not (Test-Path $ClaudeConfigDir)) {
    New-Item -ItemType Directory -Path $ClaudeConfigDir -Force | Out-Null
}

if (Test-Path $ClaudeConfigPath) {
    try {
        $ClaudeConfig = Get-Content $ClaudeConfigPath -Raw | ConvertFrom-Json
    } catch {
        $ClaudeConfig = [PSCustomObject]@{}
    }
} else {
    $ClaudeConfig = [PSCustomObject]@{}
}

Ensure-Property $ClaudeConfig "mcpServers" ([PSCustomObject]@{})

# Remove old entry if present
if ($ClaudeConfig.mcpServers.PSObject.Properties.Name -contains "sc-engagement") {
    $ClaudeConfig.mcpServers.PSObject.Properties.Remove("sc-engagement")
}

$AlfredEntry = [PSCustomObject]@{
    command = $NodePath
    args    = @($DistPath)
}

if ($ClaudeConfig.mcpServers.PSObject.Properties.Name -contains "alfred") {
    $ClaudeConfig.mcpServers.alfred = $AlfredEntry
} else {
    $ClaudeConfig.mcpServers | Add-Member -NotePropertyName "alfred" -NotePropertyValue $AlfredEntry
}

$ClaudeConfig | ConvertTo-Json -Depth 5 | Set-Content $ClaudeConfigPath -Encoding UTF8
Write-Host "   Claude Desktop config updated"

# ------------------------------------------------------------
# 5. Create Alfred.bat on Desktop
# ------------------------------------------------------------
Write-Host ""
Write-Host "[>] Creating Alfred.bat on Desktop..."

$AlfredBatPath = Join-Path $env:USERPROFILE "Desktop\Alfred.bat"

$BatContent = @"
@echo off
:: Alfred - Launch Chrome with debug profile and open key tabs
:: Auto-generated by setup.ps1

:: Check if Alfred Chrome profile is already running
wmic process where "name='chrome.exe' and commandline like '%%alfred-profile%%'" get processid 2>nul | findstr /r "[0-9]" >nul 2>&1
if %ERRORLEVEL% equ 0 (
    echo Alfred is already running.
    timeout /t 3 >nul
    exit /b 0
)

:: Find Chrome
set "CHROME_EXE="
if exist "C:\Program Files\Google\Chrome\Application\chrome.exe" (
    set "CHROME_EXE=C:\Program Files\Google\Chrome\Application\chrome.exe"
)
if exist "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" (
    set "CHROME_EXE=C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
)

if "%CHROME_EXE%"=="" (
    echo ERROR: Google Chrome not found in Program Files.
    echo Please install Chrome from https://www.google.com/chrome/
    pause
    exit /b 1
)

:: Create profile directory
if not exist "%USERPROFILE%\.alfred-profile" mkdir "%USERPROFILE%\.alfred-profile"

:: Launch Chrome with Alfred profile
start "" "%CHROME_EXE%" ^
    --remote-debugging-port=9222 ^
    --user-data-dir="%USERPROFILE%\.alfred-profile" ^
    --no-first-run ^
    --no-default-browser-check ^
    --disable-extensions ^
    --disable-sync ^
    --disable-default-apps ^
    "$NewDynamicsUrl" ^
    "https://outlook.office.com" ^
    "https://teams.microsoft.com/v2/"

:: Background update check
start /b powershell -WindowStyle Hidden -ExecutionPolicy Bypass -Command ^"& { ^
    try { ^
        `$cfg = Get-Content '%USERPROFILE%\.alfred-config.json' -Raw | ConvertFrom-Json; ^
        `$installed = `$cfg.installedVersion; ^
        if (-not `$installed) { exit }; ^
        `$resp = Invoke-RestMethod -Uri 'https://api.github.com/repos/h22fred/Alfred.mcp/commits/main' -TimeoutSec 5 -ErrorAction Stop; ^
        `$latest = `$resp.sha.Substring(0,7); ^
        if (`$installed -ne `$latest) { ^
            try { ^
                [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] | Out-Null; ^
                `$template = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText02); ^
                `$textNodes = `$template.GetElementsByTagName('text'); ^
                `$textNodes.Item(0).AppendChild(`$template.CreateTextNode('Alfred Update Available')) | Out-Null; ^
                `$textNodes.Item(1).AppendChild(`$template.CreateTextNode('A new version of Alfred is available. Ask Claude: update Alfred')) | Out-Null; ^
                `$notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier('Alfred'); ^
                `$notifier.Show([Windows.UI.Notifications.ToastNotification]::new(`$template)); ^
            } catch { ^
                Write-Host 'Alfred: A new version is available. Ask Claude to update Alfred.'; ^
            } ^
        } ^
    } catch {} ^
}^"
"@

Set-Content -Path $AlfredBatPath -Value $BatContent -Encoding ASCII
Write-Host "   Alfred.bat created on Desktop"
Write-Host "   First launch: double-click Alfred.bat, then log into Dynamics, Outlook and Teams"

# ------------------------------------------------------------
# 6. Teams webhook config
# ------------------------------------------------------------
Write-Host ""
Write-Host "[>] Setting up Teams webhook for hygiene sweep notifications..."

$Config = Read-AlfredConfig
$ExistingWebhook = if ($Config.PSObject.Properties.Name -contains "teamsWebhook") { $Config.teamsWebhook } else { "" }

if ($ExistingWebhook) {
    Write-Host "   Teams webhook already configured"
    Write-Host "      $ExistingWebhook"
    Write-Host ""
    $NewWebhook = Read-Host "   Replace it? (press Enter to keep, or paste a new URL)"
    if ($NewWebhook) {
        $Config = Read-AlfredConfig
        Ensure-Property $Config "teamsWebhook" ""
        $Config.teamsWebhook = $NewWebhook
        Write-AlfredConfig $Config
        Write-Host "   Webhook updated"
    }
} else {
    Write-Host ""
    Write-Host "   To get a webhook URL:"
    Write-Host "   Teams > any channel > ... > Connectors > Incoming Webhook > configure > copy URL"
    Write-Host ""
    $NewWebhook = Read-Host "   Paste your Teams incoming webhook URL (or press Enter to skip)"
    if ($NewWebhook) {
        $Config = Read-AlfredConfig
        Ensure-Property $Config "teamsWebhook" ""
        $Config.teamsWebhook = $NewWebhook
        Write-AlfredConfig $Config
        Write-Host "   Webhook saved to $ConfigPath"
    } else {
        Write-Host "   Skipped - hygiene sweep will run without Teams notifications"
        Write-Host "      Run setup again anytime to add it, or ask Claude to configure_teams_webhook"
    }
}

# ------------------------------------------------------------
# 7. Role selection
# ------------------------------------------------------------
Write-Host ""

if ($env:ALFRED_VARIANT -eq "sales") {
    Write-Host "[>] What is your Sales role?"
    Write-Host ""
    Write-Host "   1) AE      - Account Executive (you own accounts and opportunities)"
    Write-Host "   2) Manager - Sales Manager (you oversee a team of AEs)"
    Write-Host ""
    $RoleChoice = Read-Host "   Enter 1 or 2 (default: 1)"
    switch ($RoleChoice) {
        "2" {
            $UserRole = "sales_manager"
            Write-Host "   Role set to Sales Manager - Alfred will show territory-wide pipeline"
        }
        default {
            $UserRole = "sales"
            Write-Host "   Role set to AE - Alfred will default to your own pipeline"
        }
    }
} else {
    Write-Host "[>] What is your SC role?"
    Write-Host ""
    Write-Host "   1) SC      - Solution Consultant (you have assigned opportunities in Dynamics)"
    Write-Host "   2) SSC     - Sales Support Consultant (you support SCs, no assigned pipeline)"
    Write-Host "   3) Manager - SC Manager (you want to see your team's pipeline)"
    Write-Host ""
    $RoleChoice = Read-Host "   Enter 1, 2 or 3 (default: 1)"
    switch ($RoleChoice) {
        "2" {
            $UserRole = "ssc"
            Write-Host "   Role set to SSC - Alfred will search all accounts by default"
        }
        "3" {
            $UserRole = "manager"
            Write-Host "   Role set to Manager - Alfred will browse by territory/SC by default"
        }
        default {
            $UserRole = "sc"
            Write-Host "   Role set to SC - Alfred will default to your own pipeline"
        }
    }
}

$Config = Read-AlfredConfig
Ensure-Property $Config "role" ""
$Config.role = $UserRole
Write-AlfredConfig $Config

# ------------------------------------------------------------
# 8. Engagement types
# ------------------------------------------------------------
Write-Host ""
Write-Host "[>] Which milestones do you track on your opportunities?"
Write-Host "   (Press Enter to keep all defaults, or enter numbers separated by spaces)"
Write-Host ""

if ($env:ALFRED_VARIANT -eq "sales") {
    $AllTypes = @(
        "Discovery", "Opportunity Summary", "Mutual Plan",
        "Budget", "Implementation Plan", "Stakeholder Alignment"
    )
    Write-Host "    1) Discovery               4) Budget"
    Write-Host "    2) Opportunity Summary      5) Implementation Plan"
    Write-Host "    3) Mutual Plan              6) Stakeholder Alignment"
} else {
    $AllTypes = @(
        "Business Case", "Customer Business Review", "Demo", "Discovery", "EBC",
        "Post Sale Engagement", "POV", "RFx", "Technical Win", "Workshop"
    )
    Write-Host "    1) Business Case            6) Post Sale Engagement"
    Write-Host "    2) Customer Business Review  7) POV"
    Write-Host "    3) Demo                      8) RFx"
    Write-Host "    4) Discovery                 9) Technical Win"
    Write-Host "    5) EBC                      10) Workshop"
}

Write-Host ""
$TypeSelection = Read-Host "   Your selection (e.g. 1 2 3), or Enter for all"

if ($TypeSelection.Trim()) {
    $Indices = $TypeSelection.Trim() -split '\s+' | ForEach-Object {
        $num = 0
        if ([int]::TryParse($_, [ref]$num) -and $num -ge 1 -and $num -le $AllTypes.Length) {
            $num - 1
        }
    }
    if ($Indices.Count -gt 0) {
        $SelectedTypes = $Indices | ForEach-Object { $AllTypes[$_] }
    } else {
        $SelectedTypes = $AllTypes
    }
} else {
    $SelectedTypes = $AllTypes
}

$Config = Read-AlfredConfig
Ensure-Property $Config "engagementTypes" @()
$Config.engagementTypes = @($SelectedTypes)
Write-AlfredConfig $Config
Write-Host "   Milestones: $($SelectedTypes -join ', ')"

# ------------------------------------------------------------
# 9. Scheduled tasks (replaces cron)
# ------------------------------------------------------------
Write-Host ""
Write-Host "[>] Automated jobs..."
Write-Host ""

$HygieneScheduleDesc = ""
$MeetingScheduleDesc = ""

# --- Hygiene sweep ---
$InstallHygiene = Read-Host "   Install hygiene sweep (flags missing engagements on your pipeline)? [Y/n]"

if ($InstallHygiene -notmatch "^[nN]") {
    $HygieneDay = Read-Host "   Run on which day?       [Monday]"
    if (-not $HygieneDay) { $HygieneDay = "Monday" }
    $HygieneDow = ConvertTo-DayOfWeek $HygieneDay
    while (-not $HygieneDow) {
        $HygieneDay = Read-Host "   Unknown day - try again [Monday]"
        if (-not $HygieneDay) { $HygieneDay = "Monday" }
        $HygieneDow = ConvertTo-DayOfWeek $HygieneDay
    }

    $HygieneTime = Read-Host "   Run at what time? (HH:MM 24h) [09:30]"
    if (-not $HygieneTime) { $HygieneTime = "09:30" }
    $HygieneParts = $HygieneTime -split ":"
    $HygieneHour = [int]$HygieneParts[0]
    $HygieneMin  = [int]$HygieneParts[1]

    $HygieneScheduleDesc = "$HygieneDay at $HygieneTime"
    $HygieneTimeStr = "{0:D2}:{1:D2}" -f $HygieneHour, $HygieneMin
    $HygieneLogFile = Join-Path $env:USERPROFILE ".alfred-hygiene.log"
    $HygieneScript = Join-Path $RepoDir "scripts\hygiene-sweep.mjs"

    # Build the action: run node with the hygiene script, redirect output to log
    $HygieneAction = New-ScheduledTaskAction `
        -Execute $NodePath `
        -Argument "`"$HygieneScript`"" `
        -WorkingDirectory $RepoDir

    $HygieneTrigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $HygieneDow -At $HygieneTimeStr
    $HygieneSettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

    # Remove existing task if present, then register
    Unregister-ScheduledTask -TaskName "Alfred-HygieneSweep" -Confirm:$false -ErrorAction SilentlyContinue
    try {
        Register-ScheduledTask `
            -TaskName "Alfred-HygieneSweep" `
            -Action $HygieneAction `
            -Trigger $HygieneTrigger `
            -Settings $HygieneSettings `
            -Description "Alfred CRM hygiene sweep - flags missing engagements" | Out-Null
    } catch {
        Write-Host "   Note: Using schtasks.exe (no admin required)"
        $SchtasksDow = ConvertTo-SchtasksDay $HygieneDow
        schtasks /create /tn "Alfred-HygieneSweep" /tr "`"$NodePath`" `"$HygieneScript`"" /sc weekly /d $SchtasksDow /st $HygieneTimeStr /f 2>$null
    }

    Write-Host "   Hygiene sweep scheduled: $HygieneScheduleDesc"
} else {
    Write-Host "   Hygiene sweep skipped"
}

# --- Meeting review ---
$InstallMeeting = Read-Host "   Install meeting review (matches this week's meetings to open opps)? [Y/n]"

if ($InstallMeeting -notmatch "^[nN]") {
    $MeetingDay = Read-Host "   Run on which day?       [Friday]"
    if (-not $MeetingDay) { $MeetingDay = "Friday" }
    $MeetingDow = ConvertTo-DayOfWeek $MeetingDay
    while (-not $MeetingDow) {
        $MeetingDay = Read-Host "   Unknown day - try again [Friday]"
        if (-not $MeetingDay) { $MeetingDay = "Friday" }
        $MeetingDow = ConvertTo-DayOfWeek $MeetingDay
    }

    $MeetingTime = Read-Host "   Run at what time? (HH:MM 24h) [14:00]"
    if (-not $MeetingTime) { $MeetingTime = "14:00" }
    $MeetingParts = $MeetingTime -split ":"
    $MeetingHour = [int]$MeetingParts[0]
    $MeetingMin  = [int]$MeetingParts[1]

    $MeetingScheduleDesc = "$MeetingDay at $MeetingTime"
    $MeetingTimeStr = "{0:D2}:{1:D2}" -f $MeetingHour, $MeetingMin
    $MeetingScript = Join-Path $RepoDir "scripts\post-meeting-sweep.mjs"

    $MeetingAction = New-ScheduledTaskAction `
        -Execute $NodePath `
        -Argument "`"$MeetingScript`"" `
        -WorkingDirectory $RepoDir

    $MeetingTrigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $MeetingDow -At $MeetingTimeStr
    $MeetingSettings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

    Unregister-ScheduledTask -TaskName "Alfred-MeetingReview" -Confirm:$false -ErrorAction SilentlyContinue
    try {
        Register-ScheduledTask `
            -TaskName "Alfred-MeetingReview" `
            -Action $MeetingAction `
            -Trigger $MeetingTrigger `
            -Settings $MeetingSettings `
            -Description "Alfred weekly meeting review - matches meetings to open opportunities" | Out-Null
    } catch {
        Write-Host "   Note: Using schtasks.exe (no admin required)"
        $SchtasksDow = ConvertTo-SchtasksDay $MeetingDow
        schtasks /create /tn "Alfred-MeetingReview" /tr "`"$NodePath`" `"$MeetingScript`"" /sc weekly /d $SchtasksDow /st $MeetingTimeStr /f 2>$null
    }

    Write-Host "   Meeting review scheduled: $MeetingScheduleDesc"
} else {
    Write-Host "   Meeting review skipped"
}

# ------------------------------------------------------------
# 10. Done
# ------------------------------------------------------------
Write-Host ""
Write-Host "=================================================="
Write-Host "  Setup complete!"
Write-Host "=================================================="
Write-Host ""
Write-Host "Next steps:"
Write-Host "  1. Double-click Alfred.bat on your Desktop"
Write-Host "  2. Log into Dynamics, Outlook and Teams in that window"
Write-Host "  3. Restart Claude Desktop"
Write-Host "  4. Ask Claude anything - opportunities, calendar, hygiene sweep!"
Write-Host ""

if ($HygieneScheduleDesc -or $MeetingScheduleDesc) {
    Write-Host "Automated jobs:"
    if ($HygieneScheduleDesc) { Write-Host "  * $HygieneScheduleDesc - CRM hygiene sweep" }
    if ($MeetingScheduleDesc) { Write-Host "  * $MeetingScheduleDesc - Weekly meeting review" }
    if ($ExistingWebhook -or $NewWebhook) {
        Write-Host "Results will be posted to your Teams channel."
    } else {
        Write-Host "Run setup again to add a Teams webhook for automated notifications."
    }
    Write-Host ""
}
