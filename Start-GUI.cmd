@echo off
rem Startet die grafische Oberflaeche des Sophos-Firewall-Exports im Browser.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Export-SophosFirewallConfig.ps1" -Gui %*
