@echo off
rem Starts the graphical interface of the Sophos Firewall export in the browser.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Export-SophosFirewallConfig.ps1" -Gui %*
