@echo off
powershell -NoProfile -ExecutionPolicy Bypass -Command "iex ((New-Object Net.WebClient).DownloadString('https://raw.githubusercontent.com/h22fred/Alfred.mcp/main/setup/bootstrap_windows.ps1'))"
