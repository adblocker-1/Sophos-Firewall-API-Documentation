# Sophos Firewall – Configuration Export (PowerShell + XML API)

`Export-SophosFirewallConfig.ps1` connects via HTTPS to the **XML API of a
local Sophos Firewall (SFOS)** and exports firewall rules and settings either
to an **Excel file** (one worksheet per module, with all fields as columns)
or – if the Excel export is not possible – automatically to
**JSON and XML files per module**.

In addition, an **HTML report** (`Sophos_Firewall_Configuration.html`) with
navigation, summary and one table per module is created automatically at the
end of every export – so by default you get **Excel and HTML** in a single
run. Everything is contained in **one single script**.

**New in version 2.0:** a **graphical interface** (`-Gui` or double-click
`Start-GUI.cmd`) with connection test, point-and-click section selection,
format selection and live progress per module.

## Features

- **Graphical interface** in the browser (`-Gui`): test the connection, click
  the sections, choose the format, follow the progress live, open the result
  folder and the HTML report directly – including dark mode
- Interactive prompt for the **firewall IP address** and **API credentials**
  (the password is entered hidden)
- Reachability and **login check** before the export (up to 3 attempts with
  wrong credentials)
- **Selection menu**: export only specific configuration sections or
  everything (multiple selection with commas and ranges, e.g. `1,3,5-8`,
  `A` = everything)
- **Excel export** with the PowerShell module [ImportExcel](https://www.powershellgallery.com/packages/ImportExcel)
  (installed automatically if needed):
  - worksheet **"Overview"** with section, module, record count and status
  - **one worksheet per API module** (e.g. `FirewallRule`, `IPHost`, …)
  - nested XML fields are flattened into columns
    (`NetworkPolicy.SourceZones.Zone`), lists into comma-separated values
    (`HTTP, HTTPS, DNS`)
  - auto filter, frozen and bold header row
- **Automatic fallback**: if ImportExcel is not available, a JSON file
  (processed table data) **and** an XML file (raw API response) are created
  per module, plus `_Overview.json`
- **HTML report**: after every export, a clear HTML page
  (`Sophos_Firewall_Configuration.html`) with table of contents, summary and
  one table per module is created as well (can be disabled with `-NoHtml`)
- Fully **scriptable** via parameters (for automation / Task Scheduler)
- Works with **Windows PowerShell 5.1** and **PowerShell 7+**
  (Windows/Linux/macOS); self-signed firewall certificates are accepted

## Requirements

1. **PowerShell 5.1 or newer** (recommended: PowerShell 7)
2. **Enable API access on the Sophos Firewall:**
   - Sign in to the web admin console (default: `https://<firewall-ip>:4444`)
   - **Backup & firmware → API**
   - **Enable API configuration**
   - Add the **IP address of the computer running the script** under
     *"Allowed IP address"*
3. An **administrator account** of the firewall (username + password) – the
   XML API uses the regular admin credentials
4. For the Excel export: internet access to the PowerShell Gallery **or** the
   ImportExcel module already installed
   (`Install-Module ImportExcel -Scope CurrentUser`)

> **Note:** Excel itself does **not** need to be installed – the `.xlsx` file
> is created directly by the ImportExcel module.

## Graphical interface (GUI)

**Windows:** double-click **`Start-GUI.cmd`** – or in PowerShell:

```powershell
.\Export-SophosFirewallConfig.ps1 -Gui
```

The script starts a small web server **on this computer only**
(`127.0.0.1`) and opens the interface in the default browser. The
PowerShell window must stay open meanwhile; it additionally shows a log.

1. **Connection:** enter firewall IP, port, username and password →
   *Test connection* (checks reachability and login)
2. **Configuration sections:** click the sections you want (*All* / *None*)
3. **Output:** choose the format (*Automatic*, *Excel*, *JSON*, *XML*), HTML
   report, target folder and timeout
4. **Start export:** progress and record count per module appear live;
   afterwards you can open the created files, the HTML report and the folder
   directly. *Quit* stops the interface.

Other parameters are used to pre-fill the form (except the password):

```powershell
.\Export-SophosFirewallConfig.ps1 -Gui -FirewallIP 192.168.1.1 -Username admin -OutputFormat Json
```

> **Security:** The GUI server listens on `127.0.0.1` only, requires a
> random session token with every request (included in the address the
> script opens) and checks the Host header. Credentials are kept in memory
> only and are not stored.

## Quick start (console, interactive)

```powershell
# allow running scripts if necessary (once, current session only):
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

.\Export-SophosFirewallConfig.ps1
```

Workflow:

1. Enter the **firewall IP address** (e.g. `192.168.1.1`)
2. Enter the **API username** and **API password** (password hidden)
3. Choose the sections in the **selection menu**:

```
=============================================================
  Which configuration sections should be exported?
=============================================================
  [ 1] Firewall rules                 (FirewallRule, FirewallRuleGroup)
  [ 2] NAT rules                      (NATRule)
  [ 3] Network objects (hosts)        (IPHost, IPHostGroup, FQDNHost, ...)
  [ 4] Services                       (Services, ServiceGroup)
  ...
  [ A] Export EVERYTHING
  [ Q] Cancel

Multiple selection with commas and ranges, e.g.: 1,3,5-8
Selection: A
```

4. The export is saved in the subfolder `SophosExport_<IP>_<timestamp>`.

## Usage with parameters (automation)

```powershell
# Export everything to Excel:
.\Export-SophosFirewallConfig.ps1 -FirewallIP 192.168.1.1 `
    -Username admin -Password 'Secret123!' -Sections All

# Only firewall rules and hosts, as JSON:
.\Export-SophosFirewallConfig.ps1 -FirewallIP 192.168.1.1 `
    -Username admin -Password 'Secret123!' `
    -Sections FirewallRules,Hosts -OutputFormat Json

# Save raw XML API responses (e.g. as backup/documentation):
.\Export-SophosFirewallConfig.ps1 -FirewallIP 192.168.1.1 `
    -Username admin -Password 'Secret123!' -Sections All -OutputFormat Xml

# List the available section keys:
.\Export-SophosFirewallConfig.ps1 -ListSections
```

### Parameters

| Parameter          | Description                                                                  | Default  |
|--------------------|------------------------------------------------------------------------------|----------|
| `-FirewallIP`      | IP address or hostname of the firewall (otherwise prompted)                  | –        |
| `-Port`            | HTTPS port of the admin console/API                                          | `4444`   |
| `-Username`        | API username (otherwise prompted)                                            | –        |
| `-Password`        | API password in plain text – **for automation only** (otherwise hidden prompt) | –      |
| `-Sections`        | Sections (`All` or keys, see `-ListSections`)                                | menu     |
| `-OutputFormat`    | `Auto` (Excel, otherwise fallback), `Excel`, `Json`, `Xml`                   | `Auto`   |
| `-OutputDirectory` | Target directory                                                             | `SophosExport_<IP>_<time>` |
| `-NoHtml`          | Suppresses the additional HTML report                                        | –        |
| `-ListSections`    | Lists all section keys and exits                                             | –        |
| `-TimeoutSec`      | Timeout per API request (seconds)                                            | `60`     |
| `-Gui`             | Starts the graphical interface in the browser                                | –        |
| `-GuiPort`         | Local port of the interface (`0` = pick a free port automatically)           | `0`      |
| `-NoBrowser`       | With `-Gui`: do not open a browser, only show the address in the console     | –        |

### Exportable sections

| No | Key (`-Sections`)  | Section                     | API modules (Sophos XML API)                                 |
|----|--------------------|-----------------------------|--------------------------------------------------------------|
| 1  | `FirewallRules`    | Firewall rules              | `FirewallRule`, `FirewallRuleGroup`                          |
| 2  | `NatRules`         | NAT rules                   | `NATRule`                                                    |
| 3  | `Hosts`            | Network objects (hosts)     | `IPHost`, `IPHostGroup`, `FQDNHost`, `FQDNHostGroup`, `MACHost` |
| 4  | `Services`         | Services                    | `Services`, `ServiceGroup`                                   |
| 5  | `Zones`            | Zones                       | `Zone`                                                       |
| 6  | `Interfaces`       | Interfaces and network      | `Interface`, `VLAN`, `LAG`, `BridgePair`                     |
| 7  | `Routing`          | Routing (static routes)     | `UnicastRoute`                                               |
| 8  | `DnsDhcp`          | DNS and DHCP                | `DNS`, `DHCPServer`                                          |
| 9  | `Users`            | Users and groups            | `User`, `Group`                                              |
| 10 | `WebFilter`        | Web filter                  | `WebFilterPolicy`, `WebFilterURLGroup`                       |
| 11 | `AppFilter`        | Application filter          | `ApplicationFilterPolicy`                                    |
| 12 | `IPS`              | IPS policies                | `IPSPolicy`                                                  |
| 13 | `VPN`              | VPN (IPsec / SSL)           | `VPNIPSecConnection`, `SSLVPNPolicy`                         |
| 14 | `Wireless`         | Wireless                    | `WirelessNetworks`, `AccessPoint`                            |
| 15 | `Certificates`     | Certificates                | `Certificate`, `CertificateAuthority`                        |
| 16 | `Schedules`        | Schedules                   | `Schedule`                                                   |
| 17 | `System`           | System and administration   | `AdminSettings`, `Notification`, `SyslogServers`, `SNMPAgentConfiguration`, `BackupRestore` |

Modules without configured objects report
`Status 526 - Record does not exist` in the overview – this is normal and not
an error.

## Output

**Excel mode** (`-OutputFormat Auto`/`Excel`):

```
SophosExport_192.168.1.1_2026-07-21_18-00-00/
├── Sophos_Firewall_Export_2026-07-21_18-00-00.xlsx
│   ├── Sheet "Overview"     (all modules with count + status)
│   ├── Sheet "FirewallRule" (one row per rule, all fields as columns)
│   ├── Sheet "IPHost"
│   └── ... (one sheet per module with data)
└── Sophos_Firewall_Configuration.html   (HTML report)
```

**File mode** (`-OutputFormat Json`/`Xml` or automatic fallback):

```
SophosExport_192.168.1.1_2026-07-21_18-00-00/
├── FirewallRules_FirewallRule.json   (processed table data)
├── FirewallRules_FirewallRule.xml    (raw API response)
├── Hosts_IPHost.json
├── ...
├── _Overview.json
└── Sophos_Firewall_Configuration.html   (HTML report)
```

## HTML report

The HTML report is built into `Export-SophosFirewallConfig.ps1` and is
created automatically with every export (can be disabled with `-NoHtml`).
It contains:

- a header with device (firewall IP), export and creation date
- side navigation with the record count per module
- a summary table of all sections
- one section with a table per module; empty artifact columns are removed,
  very long values are truncated

The file is self-contained (no internet/JavaScript required) and can be
opened directly in the browser or shared.

## Technical background

For every module the script sends a `Get` request to the XML API endpoint of
the firewall:

```
POST https://<firewall-ip>:4444/webconsole/APIController
     reqxml=<Request>
              <Login><Username>...</Username><Password>...</Password></Login>
              <Get><FirewallRule/></Get>
            </Request>
```

The XML response is parsed, every record is recursively flattened into a
table row (nested levels joined with `.`, lists combined with `, `) and then
exported.

## Troubleshooting

| Message / problem                                             | Cause and solution |
|---------------------------------------------------------------|--------------------|
| `Login failed: wrong username or password`                    | Check the credentials (admin account of the firewall). |
| `API access denied: API operation not permitted from …`       | The IP of the script computer is not allowed as API IP on the firewall → add it under **Backup & firmware → API**. |
| `… is not reachable (TCP timeout)`                            | Check IP/port (default 4444) and the firewall rules/zone access (admin services HTTPS) for the source network. |
| `Excel export not possible (module "ImportExcel" missing)`    | Without internet access: run `Install-Module ImportExcel` on a computer with internet and copy the module – or use `-OutputFormat Json`/`Xml`. |
| Sheet missing in the Excel file                               | Modules without records are only listed in the overview (status 526). |
| Script does not start (`... cannot be loaded`)                | `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass` |
| GUI: browser does not open                                    | Open the address shown in the console (`http://127.0.0.1:<port>/?t=…`) manually in the browser. |
| GUI: "No connection to the export script"                     | The PowerShell window was closed or *Quit* was pressed – start the GUI again. |

## Security notes

- The `-Password` parameter (plain text) is intended for automation only –
  interactively the password is prompted **hidden** and not stored.
- Certificate validation is disabled for the firewall connection because
  Sophos Firewalls use self-signed certificates by default.
- The exported files contain the complete configuration (e.g. rule set,
  objects) – treat them as confidential.

## Tested

- Version 2.0: GUI with PowerShell 7.4 (Linux) and Chromium against a
  simulated Sophos XML API: connection test (correct/wrong password),
  section selection, export as Excel (incl. HTML report) and JSON, module
  with an invalid response, quit; light/dark and mobile layout
- PowerShell 7.6 (Linux) against a simulated Sophos XML API
  (endpoint `/webconsole/APIController`, SFOS response format), including:
  - interactive workflow (IP → credentials → menu, selection `1,3-5` and `A`)
  - wrong credentials (3 attempts, clean abort)
  - invalid menu input, unreachable firewall, status code 526
  - Excel, JSON and XML output incl. automatic fallback
