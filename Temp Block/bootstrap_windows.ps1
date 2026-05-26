# Alfred MCP — Windows bootstrap
# Downloads Setup_Windows.bat to a temp file and runs it.
#
# Safe one-liner to paste in PowerShell (does NOT use IEX or DownloadString):
#
#   $f="$env:TEMP\Setup_Windows.bat"; Invoke-WebRequest -Uri "https://raw.githubusercontent.com/h22fred/Alfred.mcp/main/Setup_Windows.bat" -OutFile $f; Start-Process cmd.exe -ArgumentList "/c `"$f`"" -Wait -NoNewWindow
#
# Or just download Setup_Windows.bat from the README and double-click it — that is the preferred method.

$ErrorActionPreference = "Stop"

$installerUrl = "https://raw.githubusercontent.com/h22fred/Alfred.mcp/main/Setup_Windows.bat"
$dest = Join-Path $env:TEMP "Alfred_Setup_Windows.bat"

Write-Host "Downloading Alfred installer..."
Invoke-WebRequest -Uri $installerUrl -OutFile $dest -UseBasicParsing

Write-Host "Running installer..."
Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$dest`"" -Wait -NoNewWindow
