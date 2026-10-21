# Changelog

## [2.0.0] - 2026-10-21

### Graphical interface
- New `-Gui` mode, or double-click **`Start-GUI.cmd`**: the interface opens in your browser
- Test the connection (reachability + login), pick sections with a click, choose the format (Automatic/Excel/JSON/XML), HTML report, target folder and timeout
- Live progress per module with cancel; afterwards open the HTML report and the folder directly
- Light/dark mode and mobile layout
- Secure: runs on `127.0.0.1` only, random session token, credentials are not stored
- New parameters: `-Gui`, `-GuiPort`, `-NoBrowser`

![GUI after an export](https://raw.githubusercontent.com/adblocker-1/Sophos-Firewall-API-Documentation/main/docs/screenshots/gui-export-done.png)

### Everything in English
- Console output, GUI, HTML report, Excel overview and README are now in English
- Renamed output files: `Sophos_Firewall_Configuration.html`, `_Overview.json`, Excel sheet `Overview`

### Bug fix
- The HTML report failed whenever the ImportExcel module was installed ("needs the PowerHTML module"). Excel and HTML are now created together as intended.

All existing parameters keep working.

**Full Changelog**: https://github.com/adblocker-1/Sophos-Firewall-API-Documentation/compare/v1.0.0...v2.0.0

## [1.0.0]

- Initial release: export of the Sophos Firewall configuration via the XML API to Excel, JSON and XML, with HTML report
