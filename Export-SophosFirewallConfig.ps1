<#
.SYNOPSIS
    Exports the configuration of a Sophos Firewall (SFOS) via the XML API
    to an Excel file (one worksheet per module) or alternatively to
    JSON/XML files per module.

.DESCRIPTION
    The script connects via HTTPS to the XML API of a local Sophos
    Firewall (default port 4444, endpoint /webconsole/APIController).

    Workflow:
      1. Prompt for the firewall IP address
      2. Prompt for the API credentials (username + password)
      3. Selection menu: export specific configuration sections or everything
      4. Export to Excel (module "ImportExcel", one worksheet per module).
         If the Excel export is not possible, JSON and XML files are
         created per module automatically.
      5. In addition, an HTML report is always created
         (Sophos_Firewall_Configuration.html), can be disabled with -NoHtml.

    With -Gui a graphical interface starts in the browser instead of the
    console prompts (connection, section selection, output and progress).

    Requirements on the firewall:
      - Enable API access under: Backup & firmware > API
      - The IP address of the computer running this script must be added
        there as an allowed API IP address.

.PARAMETER FirewallIP
    IP address (or hostname) of the Sophos Firewall. Prompted
    interactively if omitted.

.PARAMETER Port
    HTTPS port of the web admin console / API (default: 4444).

.PARAMETER Username
    API username. Prompted interactively if omitted.

.PARAMETER Password
    API password in plain text (intended for automation only!).
    Prompted interactively and hidden if omitted.

.PARAMETER Sections
    List of sections to export (keys, see -ListSections) or 'All' for
    everything. If omitted, an interactive selection menu is shown.

.PARAMETER OutputFormat
    Auto  = Excel, automatic fallback to JSON+XML on problems (default)
    Excel = force Excel file
    Json  = one JSON file per module
    Xml   = one raw XML file (API response) per module

.PARAMETER OutputDirectory
    Target directory for the export. Default: subfolder
    "SophosExport_<IP>_<timestamp>" in the current directory.

.PARAMETER NoHtml
    Suppresses the additional HTML report
    (Sophos_Firewall_Configuration.html).

.PARAMETER ListSections
    Only lists the available section keys and exits.

.PARAMETER TimeoutSec
    Timeout per API request in seconds (default: 60).

.PARAMETER Gui
    Starts the graphical interface in the default browser. All other
    parameters (except -Password) are used to pre-fill the form.
    The interface is served by a local web server that only listens on
    127.0.0.1 and requires a random session token.

.PARAMETER GuiPort
    Local port for the interface (default: 0 = pick a free port).

.PARAMETER NoBrowser
    Does not open a browser when the GUI starts; the address is only
    shown in the console.

.EXAMPLE
    .\Export-SophosFirewallConfig.ps1
    Fully interactive: prompts for IP, credentials and sections.

.EXAMPLE
    .\Export-SophosFirewallConfig.ps1 -FirewallIP 192.168.1.1 -Username apiadmin -Password 'Secret123!' -Sections All
    Non-interactive: exports all sections to Excel.

.EXAMPLE
    .\Export-SophosFirewallConfig.ps1 -FirewallIP 192.168.1.1 -Username apiadmin -Password 'Secret123!' -Sections FirewallRules,Hosts -OutputFormat Json

.EXAMPLE
    .\Export-SophosFirewallConfig.ps1 -Gui
    Starts the graphical interface in the default browser.

.NOTES
    Version 2.0.0
    Tested with PowerShell 5.1 and PowerShell 7.x.
    The connection accepts the firewall's self-signed certificates.
#>

#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]$FirewallIP,
    [ValidateRange(1, 65535)]
    [int]$Port = 4444,
    [string]$Username,
    [string]$Password,
    [string[]]$Sections,
    [ValidateSet('Auto', 'Excel', 'Json', 'Xml')]
    [string]$OutputFormat = 'Auto',
    [string]$OutputDirectory,
    [switch]$NoHtml,
    [switch]$ListSections,
    [ValidateRange(5, 600)]
    [int]$TimeoutSec = 60,
    [switch]$Gui,
    [ValidateRange(0, 65535)]
    [int]$GuiPort = 0,
    [switch]$NoBrowser
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Definition of the exportable configuration sections.
# "Modules" are the XML API entities of the Sophos Firewall (Get requests).
# ---------------------------------------------------------------------------
$script:SectionDefinitions = @(
    [pscustomobject]@{ Nr =  1; Key = 'FirewallRules';   Name = 'Firewall rules';            Modules = @('FirewallRule', 'FirewallRuleGroup') }
    [pscustomobject]@{ Nr =  2; Key = 'NatRules';        Name = 'NAT rules';                 Modules = @('NATRule') }
    [pscustomobject]@{ Nr =  3; Key = 'Hosts';           Name = 'Network objects (hosts)';   Modules = @('IPHost', 'IPHostGroup', 'FQDNHost', 'FQDNHostGroup', 'MACHost') }
    [pscustomobject]@{ Nr =  4; Key = 'Services';        Name = 'Services';                  Modules = @('Services', 'ServiceGroup') }
    [pscustomobject]@{ Nr =  5; Key = 'Zones';           Name = 'Zones';                     Modules = @('Zone') }
    [pscustomobject]@{ Nr =  6; Key = 'Interfaces';      Name = 'Interfaces and network';    Modules = @('Interface', 'VLAN', 'LAG', 'BridgePair') }
    [pscustomobject]@{ Nr =  7; Key = 'Routing';         Name = 'Routing (static routes)';   Modules = @('UnicastRoute') }
    [pscustomobject]@{ Nr =  8; Key = 'DnsDhcp';         Name = 'DNS and DHCP';              Modules = @('DNS', 'DHCPServer') }
    [pscustomobject]@{ Nr =  9; Key = 'Users';           Name = 'Users and groups';          Modules = @('User', 'Group') }
    [pscustomobject]@{ Nr = 10; Key = 'WebFilter';       Name = 'Web filter';                Modules = @('WebFilterPolicy', 'WebFilterURLGroup') }
    [pscustomobject]@{ Nr = 11; Key = 'AppFilter';       Name = 'Application filter';        Modules = @('ApplicationFilterPolicy') }
    [pscustomobject]@{ Nr = 12; Key = 'IPS';             Name = 'IPS policies';              Modules = @('IPSPolicy') }
    [pscustomobject]@{ Nr = 13; Key = 'VPN';             Name = 'VPN (IPsec / SSL)';         Modules = @('VPNIPSecConnection', 'SSLVPNPolicy') }
    [pscustomobject]@{ Nr = 14; Key = 'Wireless';        Name = 'Wireless';                  Modules = @('WirelessNetworks', 'AccessPoint') }
    [pscustomobject]@{ Nr = 15; Key = 'Certificates';    Name = 'Certificates';              Modules = @('Certificate', 'CertificateAuthority') }
    [pscustomobject]@{ Nr = 16; Key = 'Schedules';       Name = 'Schedules';                 Modules = @('Schedule') }
    [pscustomobject]@{ Nr = 17; Key = 'System';          Name = 'System and administration'; Modules = @('AdminSettings', 'Notification', 'SyslogServers', 'SNMPAgentConfiguration', 'BackupRestore') }
)

$script:Version = '2.0.0'
$script:IsCoreEdition = $PSVersionTable.PSEdition -eq 'Core'

# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------
function Write-Info    { param([string]$Text) Write-Host $Text -ForegroundColor Cyan }
function Write-Ok      { param([string]$Text) Write-Host $Text -ForegroundColor Green }
function Write-WarnMsg { param([string]$Text) Write-Host $Text -ForegroundColor Yellow }
function Write-ErrMsg  { param([string]$Text) Write-Host $Text -ForegroundColor Red }

function Enable-InsecureTls {
    # Sophos Firewalls usually use self-signed certificates.
    # Windows PowerShell 5.1 requires disabling certificate validation
    # globally; PowerShell 7+ uses -SkipCertificateCheck instead.
    if ($script:IsCoreEdition) { return }
    [System.Net.ServicePointManager]::SecurityProtocol = `
        [System.Net.SecurityProtocolType]::Tls12 -bor [System.Net.SecurityProtocolType]::Tls11
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
}

function ConvertFrom-SecureStringPlain {
    param([securestring]$SecureString)
    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
    try {
        return [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    }
    finally {
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }
}

function Read-PasswordInput {
    # Hidden password input. With redirected input (e.g. pipe/script) the
    # console does not support SecureString input - read plain text then.
    param([string]$Prompt)
    if ([System.Console]::IsInputRedirected) {
        return (Read-Host $Prompt)
    }
    $secure = Read-Host $Prompt -AsSecureString
    return (ConvertFrom-SecureStringPlain -SecureString $secure)
}

function Invoke-SophosApiRequest {
    <#
        Sends a Get request for an API entity and returns an object with
        RawXml (response as text), Records (XmlElement list) and
        StatusText (note, e.g. "no records").
    #>
    param(
        [Parameter(Mandatory)][string]$EntityName
    )

    $user = [System.Security.SecurityElement]::Escape($script:ApiUser)
    $pass = [System.Security.SecurityElement]::Escape($script:ApiPassword)
    $requestXml = "<Request><Login><Username>$user</Username><Password>$pass</Password></Login><Get><$EntityName/></Get></Request>"
    $uri = "https://$($script:TargetHost):$($script:TargetPort)/webconsole/APIController"

    $invokeParams = @{
        Uri             = $uri
        Method          = 'Post'
        Body            = @{ reqxml = $requestXml }
        TimeoutSec      = $script:RequestTimeout
        UseBasicParsing = $true
        ErrorAction     = 'Stop'
    }
    if ($script:IsCoreEdition) { $invokeParams.SkipCertificateCheck = $true }

    $response = Invoke-WebRequest @invokeParams
    $content  = $response.Content
    if ($content -is [byte[]]) { $content = [System.Text.Encoding]::UTF8.GetString($content) }
    $content = $content.TrimStart([char]0xFEFF).Trim()

    try {
        [xml]$xml = $content
    }
    catch {
        throw "Firewall response is not valid XML (entity '$EntityName'): $($_.Exception.Message)"
    }

    # Login check: with wrong credentials the API returns
    # <Login>Authentication Failure</Login>.
    $loginNode = $xml.SelectSingleNode('/Response/Login')
    if ($loginNode -and $loginNode.InnerText -match 'Authentication Failure') {
        throw 'AUTH_FAILURE'
    }

    $records = @($xml.SelectNodes("/Response/$EntityName"))
    $statusText = $null

    # A single record containing only a <Status> element is a status
    # message of the API (e.g. code 526 = "Record does not exist").
    if ($records.Count -eq 1) {
        $elementChildren = @($records[0].ChildNodes | Where-Object { $_.NodeType -eq [System.Xml.XmlNodeType]::Element })
        if ($elementChildren.Count -eq 1 -and $elementChildren[0].Name -eq 'Status') {
            $code = $elementChildren[0].GetAttribute('code')
            $text = $elementChildren[0].InnerText
            $statusText = if ($code) { "Status $code - $text" } else { $text }
            $records = @()
        }
    }

    # Global status message of the API (e.g. "API operation not permitted from ...")
    if ($records.Count -eq 0 -and -not $statusText) {
        $globalStatus = $xml.SelectSingleNode('/Response/Status')
        if ($globalStatus) {
            $text = $globalStatus.InnerText
            if ($text -match 'not permitted|No\. of') {
                throw "API access denied: $text (allow the IP of this computer under 'Backup & firmware > API'!)"
            }
            $statusText = $text
        }
        else {
            $statusText = 'No records found'
        }
    }

    [pscustomobject]@{
        Entity     = $EntityName
        RawXml     = $content
        Records    = $records
        StatusText = $statusText
    }
}

function Add-FlatValue {
    param(
        [System.Collections.Specialized.OrderedDictionary]$Row,
        [string]$Key,
        [string]$Value
    )
    if ($Row.Contains($Key)) {
        $existing = [string]$Row[$Key]
        if ([string]::IsNullOrEmpty($existing)) {
            $Row[$Key] = $Value
        }
        elseif (-not [string]::IsNullOrEmpty($Value)) {
            $Row[$Key] = "$existing, $Value"
        }
    }
    else {
        $Row[$Key] = $Value
    }
}

function Add-FlatElement {
    # Recursively converts a (nested) XML element into flat column=value
    # pairs. Nested levels are joined with '.', lists (repeated identical
    # elements) are combined with ', '.
    param(
        [System.Collections.Specialized.OrderedDictionary]$Row,
        [System.Xml.XmlElement]$Element,
        [string]$Prefix
    )
    $children = @($Element.ChildNodes | Where-Object { $_.NodeType -eq [System.Xml.XmlNodeType]::Element })
    if ($children.Count -eq 0) {
        $key = if ([string]::IsNullOrEmpty($Prefix)) { $Element.Name } else { $Prefix }
        Add-FlatValue -Row $Row -Key $key -Value $Element.InnerText.Trim()
        return
    }
    foreach ($child in $children) {
        $childKey = if ([string]::IsNullOrEmpty($Prefix)) { $child.Name } else { "$Prefix.$($child.Name)" }
        Add-FlatElement -Row $Row -Element $child -Prefix $childKey
    }
}

function ConvertTo-FlatRows {
    # Converts the XML records of a module into table rows and makes sure
    # that all rows have the same columns.
    param([object[]]$Records)

    $rawRows = foreach ($record in $Records) {
        $row = [ordered]@{}
        Add-FlatElement -Row $row -Element $record -Prefix ''
        $row
    }
    $rawRows = @($rawRows)
    if ($rawRows.Count -eq 0) { return @() }

    $allKeys = [System.Collections.Generic.List[string]]::new()
    foreach ($row in $rawRows) {
        foreach ($key in $row.Keys) {
            if (-not $allKeys.Contains($key)) { [void]$allKeys.Add($key) }
        }
    }

    $normalized = foreach ($row in $rawRows) {
        $obj = [ordered]@{}
        foreach ($key in $allKeys) {
            $obj[$key] = if ($row.Contains($key)) { $row[$key] } else { $null }
        }
        [pscustomobject]$obj
    }
    return @($normalized)
}

function Get-SafeSheetName {
    # Excel worksheet names: max. 31 characters, none of []:*?/\
    param(
        [string]$Name,
        [System.Collections.Generic.HashSet[string]]$UsedNames
    )
    $clean = $Name -replace '[\[\]\:\*\?\/\\]', '_'
    if ($clean.Length -gt 31) { $clean = $clean.Substring(0, 31) }
    $candidate = $clean
    $suffix = 1
    while (-not $UsedNames.Add($candidate)) {
        $suffix++
        $base = $clean
        $tail = "_$suffix"
        if (($base.Length + $tail.Length) -gt 31) { $base = $base.Substring(0, 31 - $tail.Length) }
        $candidate = "$base$tail"
    }
    return $candidate
}

function Initialize-ImportExcel {
    # Checks whether the module "ImportExcel" is available and otherwise
    # tries to install it from the PowerShell Gallery. Returns $true/$false.
    if (Get-Module -Name ImportExcel) { return $true }
    if (Get-Module -ListAvailable -Name ImportExcel) {
        Import-Module ImportExcel -ErrorAction Stop
        return $true
    }
    Write-Info 'Module "ImportExcel" not found - trying to install it from the PowerShell Gallery ...'
    try {
        Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Import-Module ImportExcel -ErrorAction Stop
        Write-Ok 'Module "ImportExcel" has been installed.'
        return $true
    }
    catch {
        Write-WarnMsg "Installing ImportExcel failed: $($_.Exception.Message)"
        return $false
    }
}

# ---------------------------------------------------------------------------
# HTML report
# ---------------------------------------------------------------------------
$script:HtmlMaxLen = 800

function Convert-HtmlSanitize {
    param([string]$s)
    if ($null -eq $s) { return '' }
    $s = [regex]::Replace($s, "[\x00-\x08\x0b\x0c\x0e-\x1f]", ' ')
    $s = [regex]::Replace($s, "[ \t]{3,}", '  ')
    $s = $s.Trim()
    if ($s.Length -gt $script:HtmlMaxLen) { $s = $s.Substring(0, $script:HtmlMaxLen) + ' ... [truncated]' }
    return $s
}

function Convert-HtmlValue {
    param($v)
    if ($null -eq $v) { return '' }
    if ($v -is [System.Collections.IEnumerable] -and -not ($v -is [string])) {
        $items = @()
        foreach ($x in $v) { if ($null -ne $x) { $items += (Convert-HtmlValue $x) } }
        return (Convert-HtmlSanitize ($items -join ', '))
    }
    if ($v -is [System.Management.Automation.PSCustomObject]) {
        $parts = @()
        foreach ($p in $v.PSObject.Properties) { $parts += ('{0}: {1}' -f $p.Name, (Convert-HtmlValue $p.Value)) }
        return (Convert-HtmlSanitize ($parts -join '; '))
    }
    return (Convert-HtmlSanitize ([string]$v))
}

function Get-HtmlEscape {
    param($s)
    $t = [string]$s
    $t = $t.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;').Replace("'", '&#x27;')
    return $t
}

function ConvertTo-HtmlReportTable {
    # Prepares the table rows of a module for the HTML report:
    # position artifact columns (ending in ".Name") and completely empty
    # columns are removed, values are sanitized and truncated.
    param($records)
    $cols = New-Object System.Collections.Generic.List[string]
    foreach ($r in $records) {
        if ($r -is [System.Management.Automation.PSCustomObject]) {
            foreach ($p in $r.PSObject.Properties) {
                if (-not $cols.Contains($p.Name)) { [void]$cols.Add($p.Name) }
            }
        }
    }
    $cols = @($cols | Where-Object { -not ($_.EndsWith('.Name') -and $_ -ne 'Name') })

    $rows = @()
    foreach ($r in $records) {
        if ($r -is [System.Management.Automation.PSCustomObject]) {
            $row = @()
            foreach ($c in $cols) {
                $prop = $r.PSObject.Properties[$c]
                if ($prop) { $row += (Convert-HtmlValue $prop.Value) } else { $row += '' }
            }
            $rows += ,$row
        } else {
            $rows += ,@((Convert-HtmlValue $r))
        }
    }
    # Remove completely empty columns
    $keep = @()
    for ($i = 0; $i -lt $cols.Count; $i++) {
        $has = $false
        foreach ($row in $rows) {
            if ($i -lt $row.Count -and $row[$i].Trim().Length -gt 0) { $has = $true; break }
        }
        if ($has) { $keep += $i }
    }
    $ncols = @()
    foreach ($i in $keep) { $ncols += $cols[$i] }
    $nrows = @()
    foreach ($row in $rows) {
        $nr = @()
        foreach ($i in $keep) { if ($i -lt $row.Count) { $nr += $row[$i] } else { $nr += '' } }
        $nrows += ,$nr
    }
    return [PSCustomObject]@{ Cols = $ncols; Rows = $nrows }
}

function New-HtmlReport {
    # Creates an HTML page from the already retrieved module results
    # with navigation, summary and one section per module.
    param(
        [Parameter(Mandatory)]$Results,
        [Parameter(Mandatory)][string]$Device,
        [Parameter(Mandatory)][string]$ExportDate,
        [Parameter(Mandatory)][string]$OutFile
    )

    $sections = @()
    foreach ($r in $Results) {
        $flat = ConvertTo-HtmlReportTable @($r.Rows)
        $sections += [PSCustomObject]@{
            Section = Convert-HtmlValue $r.Section
            Module  = Convert-HtmlValue $r.Module
            Count   = $r.Count
            Status  = Convert-HtmlValue $r.Status
            Cols    = @($flat.Cols)
            Rows    = @($flat.Rows)
        }
    }
    $today = (Get-Date).ToString('yyyy-MM-dd')

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append(@"
<!DOCTYPE html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Sophos Firewall &#8211; Configuration $(Get-HtmlEscape $Device)</title>
<style>
:root{--nav:#1f3864;--nav2:#2e5496;--line:#d0d7e2;--alt:#f2f6fc;--muted:#6b7280;}
*{box-sizing:border-box}
body{margin:0;font-family:'Segoe UI',Arial,sans-serif;color:#1a1a1a;background:#f4f6f9;}
header{background:linear-gradient(135deg,#1f3864,#2e5496);color:#fff;padding:28px 40px;}
header h1{margin:0 0 6px;font-size:24px;}
header .meta{font-size:13px;opacity:.9;}
.wrap{display:flex;align-items:flex-start;gap:24px;max-width:1500px;margin:0 auto;padding:24px;}
nav{position:sticky;top:16px;flex:0 0 260px;background:#fff;border:1px solid var(--line);border-radius:10px;padding:14px;max-height:90vh;overflow:auto;}
nav h2{font-size:12px;text-transform:uppercase;letter-spacing:.05em;color:var(--muted);margin:0 0 8px;}
nav a{display:flex;justify-content:space-between;gap:8px;padding:5px 8px;border-radius:6px;color:#233;text-decoration:none;font-size:13px;}
nav a:hover{background:var(--alt);}
nav a .cnt{color:var(--muted);font-variant-numeric:tabular-nums;}
nav a.empty{color:#aeb4bd;}
main{flex:1;min-width:0;}
section{background:#fff;border:1px solid var(--line);border-radius:10px;padding:20px 22px;margin-bottom:22px;}
section h2{margin:0 0 2px;font-size:18px;color:var(--nav);}
section .cat{font-size:12px;color:var(--muted);text-transform:uppercase;letter-spacing:.04em;margin-bottom:12px;}
.tblwrap{overflow-x:auto;border:1px solid var(--line);border-radius:8px;}
table{border-collapse:collapse;width:100%;font-size:13px;}
th{background:var(--nav2);color:#fff;text-align:left;padding:8px 10px;position:sticky;top:0;white-space:nowrap;}
td{padding:7px 10px;border-top:1px solid var(--line);vertical-align:top;}
tr:nth-child(even) td{background:var(--alt);}
.empty-note{color:var(--muted);font-style:italic;font-size:13px;}
.pill{display:inline-block;background:#e8eefb;color:#1f3864;border-radius:12px;padding:1px 9px;font-size:12px;font-weight:600;}
.summary{width:100%;border-collapse:collapse;font-size:13px;}
.summary th,.summary td{border:1px solid var(--line);padding:7px 10px;text-align:left;}
.summary th{background:var(--nav);color:#fff;}
.summary td.n{text-align:center;font-variant-numeric:tabular-nums;}
.summary tr.zero td{color:#aeb4bd;}
footer{text-align:center;color:var(--muted);font-size:12px;padding:20px;}
</style></head><body>
<header>
<h1>Sophos Firewall &#8211; Configuration Overview</h1>
<div class="meta">Device $(Get-HtmlEscape $Device) &nbsp;&#183;&nbsp; Export $(Get-HtmlEscape $ExportDate) &nbsp;&#183;&nbsp; created $today</div>
</header>
<div class="wrap">
<nav><h2>Contents</h2>
"@)

    # Navigation
    for ($i = 0; $i -lt $sections.Count; $i++) {
        $s = $sections[$i]
        $cls = ''
        if ($s.Count -eq 0) { $cls = ' class=empty' }
        [void]$sb.Append(('<a href="#s{0}"{1}><span>{2}</span><span class="cnt">{3}</span></a>' -f $i, $cls, (Get-HtmlEscape $s.Module), $s.Count))
    }
    [void]$sb.Append('</nav><main>')

    # Summary
    [void]$sb.Append('<section id="overview"><h2>Summary</h2><div class="cat">All configuration sections</div>')
    [void]$sb.Append('<table class="summary"><thead><tr><th>Section</th><th>Module</th><th>Records</th><th>Status</th></tr></thead><tbody>')
    $last = $null
    foreach ($s in $sections) {
        $zc = ''
        if ($s.Count -eq 0) { $zc = ' class=zero' }
        $bshow = ''
        if ($s.Section -ne $last) { $bshow = (Get-HtmlEscape $s.Section) }
        $last = $s.Section
        if ($s.Count -eq 0) { $st = 'empty' } else { $st = (Get-HtmlEscape $s.Status) }
        [void]$sb.Append(('<tr{0}><td>{1}</td><td>{2}</td><td class="n">{3}</td><td>{4}</td></tr>' -f $zc, $bshow, (Get-HtmlEscape $s.Module), $s.Count, $st))
    }
    [void]$sb.Append('</tbody></table></section>')

    # Sections
    for ($i = 0; $i -lt $sections.Count; $i++) {
        $s = $sections[$i]
        [void]$sb.Append(('<section id="s{0}"><h2>{1} <span class="pill">{2}</span></h2><div class="cat">{3}</div>' -f $i, (Get-HtmlEscape $s.Module), $s.Count, (Get-HtmlEscape $s.Section)))
        if ($s.Count -eq 0 -or $s.Cols.Count -eq 0) {
            [void]$sb.Append('<p class="empty-note">No entries.</p>')
        } else {
            [void]$sb.Append('<div class="tblwrap"><table><thead><tr>')
            foreach ($c in $s.Cols) { [void]$sb.Append(('<th>{0}</th>' -f (Get-HtmlEscape $c))) }
            [void]$sb.Append('</tr></thead><tbody>')
            foreach ($row in $s.Rows) {
                [void]$sb.Append('<tr>')
                foreach ($v in $row) { [void]$sb.Append(('<td>{0}</td>' -f (Get-HtmlEscape $v))) }
                [void]$sb.Append('</tr>')
            }
            [void]$sb.Append('</tbody></table></div>')
        }
        [void]$sb.Append('</section>')
    }

    [void]$sb.Append('</main></div><footer>Generated from the Sophos export &#183; for internal documentation only</footer></body></html>')

    # Write (UTF-8 without BOM)
    $enc = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($OutFile, $sb.ToString(), $enc)
}

function Test-TcpPort {
    param(
        [string]$TargetHost,
        [int]$TargetPort,
        [int]$TimeoutMs = 5000
    )
    $client = New-Object System.Net.Sockets.TcpClient
    try {
        $async = $client.BeginConnect($TargetHost, $TargetPort, $null, $null)
        if (-not $async.AsyncWaitHandle.WaitOne($TimeoutMs, $false)) { return $false }
        $client.EndConnect($async)
        return $true
    }
    catch {
        return $false
    }
    finally {
        $client.Close()
    }
}

function Show-SectionMenu {
    # Shows the selection menu and returns the selected section objects.
    Write-Host ''
    Write-Host '=============================================================' -ForegroundColor White
    Write-Host '  Which configuration sections should be exported?'          -ForegroundColor White
    Write-Host '=============================================================' -ForegroundColor White
    foreach ($section in $script:SectionDefinitions) {
        $line = '  [{0,2}] {1,-30} ({2})' -f $section.Nr, $section.Name, ($section.Modules -join ', ')
        Write-Host $line
    }
    Write-Host '  [ A] Export EVERYTHING' -ForegroundColor Green
    Write-Host '  [ Q] Cancel'            -ForegroundColor Yellow
    Write-Host ''
    Write-Host 'Multiple selection with commas and ranges, e.g.: 1,3,5-8' -ForegroundColor DarkGray

    for ($attempt = 1; $attempt -le 5; $attempt++) {
        $answer = Read-Host 'Selection'
        if ([string]::IsNullOrWhiteSpace($answer)) { continue }
        $answer = $answer.Trim()

        if ($answer -match '^[Qq]$') {
            Write-WarnMsg 'Cancelled.'
            exit 0
        }
        if ($answer -match '^([Aa](ll)?|0)$') {
            return @($script:SectionDefinitions)
        }

        $numbers = [System.Collections.Generic.List[int]]::new()
        $valid = $true
        foreach ($part in ($answer -split '[,;\s]+')) {
            if ([string]::IsNullOrWhiteSpace($part)) { continue }
            if ($part -match '^(\d+)-(\d+)$') {
                $from = [int]$Matches[1]; $to = [int]$Matches[2]
                if ($from -gt $to) { $valid = $false; break }
                foreach ($n in $from..$to) { [void]$numbers.Add($n) }
            }
            elseif ($part -match '^\d+$') {
                [void]$numbers.Add([int]$part)
            }
            else {
                $valid = $false; break
            }
        }

        if ($valid -and $numbers.Count -gt 0) {
            $selected = @($script:SectionDefinitions | Where-Object { $numbers -contains $_.Nr })
            $unknown  = @($numbers | Where-Object { $_ -lt 1 -or $_ -gt $script:SectionDefinitions.Count } | Select-Object -Unique)
            if ($unknown.Count -gt 0) {
                Write-WarnMsg "Invalid number(s): $($unknown -join ', '). Please try again."
                continue
            }
            if ($selected.Count -gt 0) { return $selected }
        }
        Write-WarnMsg 'Input not understood. Examples: "1,3", "5-8", "A" for everything.'
    }
    throw 'Too many invalid inputs - aborting.'
}

function Resolve-SectionsFromParameter {
    param([string[]]$Requested)
    # Also accept "A,B" as a single argument (e.g. with pwsh -File)
    $Requested = @($Requested | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    if ($Requested -contains 'All') {
        return @($script:SectionDefinitions)
    }
    $resolved = [System.Collections.Generic.List[object]]::new()
    foreach ($name in $Requested) {
        $match = $script:SectionDefinitions | Where-Object { $_.Key -eq $name }
        if (-not $match) {
            $validKeys = ($script:SectionDefinitions.Key -join ', ')
            throw "Unknown section '$name'. Valid values: All, $validKeys"
        }
        [void]$resolved.Add($match)
    }
    return @($resolved)
}

# ---------------------------------------------------------------------------
# Export workflow (shared by console and GUI)
# ---------------------------------------------------------------------------

function Set-SophosTarget {
    param(
        [Parameter(Mandatory)][string]$TargetHost,
        [int]$TargetPort = 4444,
        [int]$Timeout = 60
    )
    $script:TargetHost     = $TargetHost
    $script:TargetPort     = $TargetPort
    $script:RequestTimeout = $Timeout
}

function Test-SophosLogin {
    # Checks the credentials with a simple Get request.
    # Throws 'AUTH_FAILURE' on wrong credentials, otherwise the error message.
    param(
        [string]$User,
        [string]$Pass
    )
    $script:ApiUser     = $User
    $script:ApiPassword = $Pass
    [void](Invoke-SophosApiRequest -EntityName 'Zone')
}

function Get-SophosModuleResult {
    # Retrieves an API module and returns a result object for the export.
    # Errors of individual modules are recorded in the status; only a
    # login failure ('AUTH_FAILURE') is passed on.
    param(
        [Parameter(Mandatory)]$Section,
        [Parameter(Mandatory)][string]$Module
    )
    try {
        $apiResult = Invoke-SophosApiRequest -EntityName $Module
        $rows = ConvertTo-FlatRows -Records $apiResult.Records
        $status = if ($apiResult.Records.Count -gt 0) { 'OK' } else { $apiResult.StatusText }
        return [pscustomobject]@{
            Section    = $Section.Name
            SectionKey = $Section.Key
            Module     = $Module
            Count      = $apiResult.Records.Count
            Status     = $status
            Failed     = $false
            Rows       = $rows
            RawXml     = $apiResult.RawXml
        }
    }
    catch {
        if ($_.Exception.Message -eq 'AUTH_FAILURE') { throw }
        return [pscustomobject]@{
            Section    = $Section.Name
            SectionKey = $Section.Key
            Module     = $Module
            Count      = 0
            Status     = "ERROR: $($_.Exception.Message)"
            Failed     = $true
            Rows       = @()
            RawXml     = $null
        }
    }
}

function New-ExportDirectory {
    param(
        [string]$OutputDirectory,
        [Parameter(Mandatory)][string]$TargetHost,
        [Parameter(Mandatory)][string]$Timestamp
    )
    if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
        $safeHost = $TargetHost -replace '[^\w\.\-]', '_'
        $OutputDirectory = Join-Path -Path (Get-Location).Path -ChildPath "SophosExport_${safeHost}_$Timestamp"
    }
    if (-not (Test-Path -LiteralPath $OutputDirectory)) {
        [void](New-Item -ItemType Directory -Path $OutputDirectory -Force)
    }
    return (Resolve-Path -LiteralPath $OutputDirectory).ProviderPath
}

function Export-SophosResults {
    # Writes the retrieved module results as Excel or JSON/XML and
    # optionally as an HTML report. Returns the created files.
    param(
        [Parameter(Mandatory)]$Results,
        [Parameter(Mandatory)][string]$OutputFormat,
        [Parameter(Mandatory)][string]$OutputDirectory,
        [Parameter(Mandatory)][string]$Timestamp,
        [switch]$NoHtml
    )

    $overview = @($Results | Select-Object `
        @{n = 'Section'; e = { $_.Section } },
        @{n = 'Module'; e = { $_.Module } },
        @{n = 'Records'; e = { $_.Count } },
        @{n = 'Status'; e = { $_.Status } })

    $notes = [System.Collections.Generic.List[string]]::new()
    $useExcel = $false
    if ($OutputFormat -in @('Auto', 'Excel')) {
        $useExcel = Initialize-ImportExcel
        if (-not $useExcel) {
            if ($OutputFormat -eq 'Excel') {
                throw 'Excel export not possible (module "ImportExcel" missing). Please install it manually: Install-Module ImportExcel'
            }
            $msg = 'Excel export not possible - falling back to JSON and XML files per module.'
            Write-WarnMsg $msg
            [void]$notes.Add($msg)
        }
    }

    $createdFiles = [System.Collections.Generic.List[string]]::new()

    if ($useExcel) {
        $excelPath = Join-Path $OutputDirectory "Sophos_Firewall_Export_$Timestamp.xlsx"
        $usedSheetNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

        $exportExcel = {
            param($Data, $SheetName)
            try {
                $Data | Export-Excel -Path $excelPath -WorksheetName $SheetName -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter
            }
            catch {
                # -AutoSize can fail on systems without a graphics library
                $Data | Export-Excel -Path $excelPath -WorksheetName $SheetName -FreezeTopRow -BoldTopRow -AutoFilter
            }
        }

        $overviewSheet = Get-SafeSheetName -Name 'Overview' -UsedNames $usedSheetNames
        & $exportExcel $overview $overviewSheet

        foreach ($result in $Results) {
            if ($result.Count -eq 0) { continue }
            $sheetName = Get-SafeSheetName -Name $result.Module -UsedNames $usedSheetNames
            & $exportExcel $result.Rows $sheetName
        }
        [void]$createdFiles.Add($excelPath)
    }
    else {
        $wantJson = $OutputFormat -in @('Auto', 'Json')
        $wantXml  = $OutputFormat -in @('Auto', 'Xml')

        foreach ($result in $Results) {
            $baseName = "$($result.SectionKey)_$($result.Module)"
            if ($wantJson -and $result.Count -gt 0) {
                $jsonPath = Join-Path $OutputDirectory "$baseName.json"
                $result.Rows | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $jsonPath -Encoding UTF8
                [void]$createdFiles.Add($jsonPath)
            }
            if ($wantXml -and $result.RawXml) {
                $xmlPath = Join-Path $OutputDirectory "$baseName.xml"
                $result.RawXml | Set-Content -LiteralPath $xmlPath -Encoding UTF8
                [void]$createdFiles.Add($xmlPath)
            }
        }

        $summaryPath = Join-Path $OutputDirectory '_Overview.json'
        $overview | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $summaryPath -Encoding UTF8
        [void]$createdFiles.Add($summaryPath)
    }

    $htmlPath = $null
    if (-not $NoHtml) {
        try {
            $htmlPath = Join-Path $OutputDirectory 'Sophos_Firewall_Configuration.html'
            Write-Info 'Creating HTML report ...'
            New-HtmlReport -Results $Results -OutFile $htmlPath `
                -Device $script:TargetHost -ExportDate (Get-Date -Format 'yyyy-MM-dd HH:mm')
            Write-Ok "HTML created: $htmlPath"
            [void]$createdFiles.Add($htmlPath)
        }
        catch {
            $msg = "HTML report could not be created: $($_.Exception.Message)"
            Write-WarnMsg $msg
            [void]$notes.Add($msg)
            $htmlPath = $null
        }
    }

    [pscustomobject]@{
        Overview  = $overview
        Files     = @($createdFiles)
        UsedExcel = $useExcel
        HtmlPath  = $htmlPath
        Notes     = @($notes)
    }
}

# ---------------------------------------------------------------------------
# Graphical interface (-Gui)
# A small HTTP server on 127.0.0.1 serves a web interface that is opened
# in the default browser. It only accepts connections from this computer;
# every request must contain the random session token.
# ---------------------------------------------------------------------------

function Open-GuiTarget {
    # Opens a URL, file or folder with the default application.
    param([Parameter(Mandatory)][string]$Target)
    if (-not $script:IsCoreEdition -or $IsWindows) {
        Start-Process -FilePath $Target
    }
    elseif ($IsMacOS) {
        Start-Process -FilePath 'open' -ArgumentList $Target
    }
    else {
        Start-Process -FilePath 'xdg-open' -ArgumentList $Target
    }
}

function Send-GuiResponse {
    param(
        [Parameter(Mandatory)][System.IO.Stream]$Stream,
        [int]$Status = 200,
        [string]$ContentType = 'application/json; charset=utf-8',
        [byte[]]$Body = @()
    )
    $reason = switch ($Status) {
        200 { 'OK' } 400 { 'Bad Request' } 403 { 'Forbidden' } 404 { 'Not Found' }
        405 { 'Method Not Allowed' } default { 'Internal Server Error' }
    }
    $head = "HTTP/1.1 $Status $reason`r`n" +
            "Content-Type: $ContentType`r`n" +
            "Content-Length: $($Body.Length)`r`n" +
            "Cache-Control: no-store`r`n" +
            "X-Content-Type-Options: nosniff`r`n" +
            "X-Frame-Options: DENY`r`n" +
            "Referrer-Policy: no-referrer`r`n" +
            "Connection: close`r`n`r`n"
    $headBytes = [System.Text.Encoding]::ASCII.GetBytes($head)
    $Stream.Write($headBytes, 0, $headBytes.Length)
    if ($Body.Length -gt 0) { $Stream.Write($Body, 0, $Body.Length) }
    $Stream.Flush()
}

function ConvertTo-GuiJson {
    param($Data)
    $json = ConvertTo-Json -InputObject $Data -Depth 8 -Compress
    return [System.Text.Encoding]::UTF8.GetBytes($json)
}

function Read-GuiRequest {
    # Reads an HTTP request (start line, headers, body) from the stream.
    param([Parameter(Mandatory)][System.IO.Stream]$Stream)

    $buffer = New-Object System.IO.MemoryStream
    $chunk = New-Object byte[] 8192
    $headerEnd = -1
    while ($headerEnd -lt 0) {
        $n = $Stream.Read($chunk, 0, $chunk.Length)
        if ($n -le 0) { return $null }
        $buffer.Write($chunk, 0, $n)
        if ($buffer.Length -gt 1MB) { return $null }
        # ASCII decoding preserves byte positions (1 character = 1 byte)
        $headerEnd = [System.Text.Encoding]::ASCII.GetString($buffer.ToArray()).IndexOf("`r`n`r`n")
    }
    $all = $buffer.ToArray()
    $lines = [System.Text.Encoding]::ASCII.GetString($all, 0, $headerEnd) -split "`r`n"
    $startLine = $lines[0] -split ' '
    if ($startLine.Count -lt 2) { return $null }

    $headers = @{}
    foreach ($line in ($lines | Select-Object -Skip 1)) {
        $idx = $line.IndexOf(':')
        if ($idx -gt 0) { $headers[$line.Substring(0, $idx).Trim().ToLowerInvariant()] = $line.Substring($idx + 1).Trim() }
    }

    $contentLength = 0
    if ($headers.ContainsKey('content-length')) { [void][int]::TryParse($headers['content-length'], [ref]$contentLength) }
    if ($contentLength -gt 1MB) { return $null }
    $body = New-Object System.IO.MemoryStream
    $bodyStart = $headerEnd + 4
    if ($all.Length -gt $bodyStart) { $body.Write($all, $bodyStart, $all.Length - $bodyStart) }
    while ($body.Length -lt $contentLength) {
        $n = $Stream.Read($chunk, 0, $chunk.Length)
        if ($n -le 0) { break }
        $body.Write($chunk, 0, $n)
    }

    $target = $startLine[1]
    $path = $target
    $query = @{}
    $qIdx = $target.IndexOf('?')
    if ($qIdx -ge 0) {
        $path = $target.Substring(0, $qIdx)
        foreach ($pair in ($target.Substring($qIdx + 1) -split '&')) {
            if (-not $pair) { continue }
            $kv = $pair -split '=', 2
            $value = if ($kv.Count -gt 1) { [uri]::UnescapeDataString($kv[1]) } else { '' }
            $query[[uri]::UnescapeDataString($kv[0])] = $value
        }
    }

    [pscustomobject]@{
        Method  = $startLine[0].ToUpperInvariant()
        Path    = $path
        Query   = $query
        Headers = $headers
        Body    = [System.Text.Encoding]::UTF8.GetString($body.ToArray())
    }
}

function Get-GuiProperty {
    # Reads a field from the JSON body (missing fields => default value).
    param($Object, [string]$Name, $Default = $null)
    if ($null -eq $Object) { return $Default }
    $prop = $Object.PSObject.Properties[$Name]
    if ($null -eq $prop -or $null -eq $prop.Value) { return $Default }
    return $prop.Value
}

function Invoke-GuiApi {
    # Handles an API call of the interface and returns an object that is
    # sent back as JSON.
    param(
        [Parameter(Mandatory)][string]$Name,
        $Data
    )
    $state = $script:GuiState

    switch ($Name) {
        'info' {
            return @{
                ok       = $true
                version  = $script:Version
                cwd      = (Get-Location).Path
                sections = @($script:SectionDefinitions | ForEach-Object {
                    @{ nr = $_.Nr; key = $_.Key; name = $_.Name; modules = @($_.Modules) }
                })
                defaults = $state.Defaults
            }
        }

        'connect' {
            $targetHost = ([string](Get-GuiProperty $Data 'host' '')).Trim()
            $port       = [int](Get-GuiProperty $Data 'port' 4444)
            $user       = [string](Get-GuiProperty $Data 'user' '')
            $pass       = [string](Get-GuiProperty $Data 'password' '')
            $timeout    = [int](Get-GuiProperty $Data 'timeout' 60)
            $state.Connected = $false

            if (-not $targetHost) { return @{ ok = $false; error = 'Please enter the IP address or hostname of the firewall.' } }
            if ($port -lt 1 -or $port -gt 65535) { return @{ ok = $false; error = 'The port must be between 1 and 65535.' } }
            if ($timeout -lt 5 -or $timeout -gt 600) { return @{ ok = $false; error = 'The timeout must be between 5 and 600 seconds.' } }
            if (-not $user -or -not $pass) { return @{ ok = $false; error = 'Please enter username and password.' } }

            Set-SophosTarget -TargetHost $targetHost -TargetPort $port -Timeout $timeout
            Write-Info "[GUI] Checking connection to ${targetHost}:${port} ..."
            $reachable = Test-TcpPort -TargetHost $targetHost -TargetPort $port
            if (-not $reachable) {
                Write-WarnMsg "[GUI] ${targetHost}:${port} is not reachable."
                return @{ ok = $false; reachable = $false; error = "${targetHost}:${port} is not reachable (TCP timeout). Check IP, port and access to the admin services." }
            }
            try {
                Test-SophosLogin -User $user -Pass $pass
            }
            catch {
                if ($_.Exception.Message -eq 'AUTH_FAILURE') {
                    Write-ErrMsg '[GUI] Login failed.'
                    return @{ ok = $false; reachable = $true; error = 'Login failed: wrong username or password.' }
                }
                Write-ErrMsg "[GUI] Connection failed: $($_.Exception.Message)"
                return @{ ok = $false; reachable = $true; error = "Connection failed: $($_.Exception.Message)" }
            }
            $state.Connected = $true
            Write-Ok "[GUI] Login to ${targetHost}:${port} successful."
            return @{ ok = $true; reachable = $true; host = $targetHost; port = $port }
        }

        'start' {
            if (-not $state.Connected) { return @{ ok = $false; error = 'Please connect to the firewall first.' } }
            $keys = @(Get-GuiProperty $Data 'sections' @())
            if ($keys.Count -eq 0) { return @{ ok = $false; error = 'Please select at least one section.' } }
            $format = [string](Get-GuiProperty $Data 'format' 'Auto')
            if ($format -notin @('Auto', 'Excel', 'Json', 'Xml')) { return @{ ok = $false; error = "Unknown output format '$format'." } }

            $selected = Resolve-SectionsFromParameter -Requested $keys
            $modules = [System.Collections.Generic.List[object]]::new()
            foreach ($section in $selected) {
                foreach ($module in $section.Modules) {
                    [void]$modules.Add([pscustomobject]@{ Section = $section; Module = $module })
                }
            }
            $state.Job = [pscustomobject]@{
                Modules         = $modules
                Results         = New-Object object[] $modules.Count
                Format          = $format
                Html            = [bool](Get-GuiProperty $Data 'html' $true)
                OutputDirectory = ([string](Get-GuiProperty $Data 'outputDirectory' '')).Trim()
                Timestamp       = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
            }
            $state.Export = $null
            Write-Info ("[GUI] Export started: " + (($selected | ForEach-Object { $_.Name }) -join ', '))
            return @{
                ok      = $true
                total   = $modules.Count
                modules = @($modules | ForEach-Object { @{ section = $_.Section.Name; module = $_.Module } })
            }
        }

        'fetch' {
            $job = $state.Job
            if ($null -eq $job) { return @{ ok = $false; error = 'No export in progress.' } }
            $index = [int](Get-GuiProperty $Data 'index' -1)
            if ($index -lt 0 -or $index -ge $job.Modules.Count) { return @{ ok = $false; error = 'Invalid module index.' } }
            $entry = $job.Modules[$index]
            try {
                $result = Get-SophosModuleResult -Section $entry.Section -Module $entry.Module
            }
            catch {
                $state.Connected = $false
                return @{ ok = $false; auth = $true; error = 'Login expired/failed during the export.' }
            }
            $job.Results[$index] = $result
            if ($result.Failed) { Write-ErrMsg ("  {0,-28} {1}" -f $result.Module, $result.Status) }
            elseif ($result.Count -gt 0) { Write-Ok ("  {0,-28} {1,4} records" -f $result.Module, $result.Count) }
            else { Write-Host ("  {0,-28} {1}" -f $result.Module, $result.Status) -ForegroundColor DarkGray }
            return @{
                ok      = $true
                module  = $result.Module
                section = $result.Section
                count   = $result.Count
                status  = [string]$result.Status
                failed  = $result.Failed
            }
        }

        'finish' {
            $job = $state.Job
            if ($null -eq $job) { return @{ ok = $false; error = 'No export in progress.' } }
            $results = @($job.Results | Where-Object { $null -ne $_ })
            if ($results.Count -eq 0) { return @{ ok = $false; error = 'No data has been retrieved.' } }

            $dir = New-ExportDirectory -OutputDirectory $job.OutputDirectory -TargetHost $script:TargetHost -Timestamp $job.Timestamp
            Write-Info "[GUI] Writing export to: $dir"
            $export = Export-SophosResults -Results $results -OutputFormat $job.Format -OutputDirectory $dir `
                -Timestamp $job.Timestamp -NoHtml:(-not $job.Html)
            $state.Export = [pscustomobject]@{ Directory = $dir; HtmlPath = $export.HtmlPath }
            $state.Job = $null
            Write-Ok "[GUI] Export completed: $dir"

            return @{
                ok        = $true
                directory = $dir
                usedExcel = $export.UsedExcel
                hasHtml   = [bool]$export.HtmlPath
                notes     = @($export.Notes)
                files     = @($export.Files | ForEach-Object {
                    $item = Get-Item -LiteralPath $_
                    @{ name = $item.Name; size = $item.Length }
                })
                records   = [int](($results | Measure-Object -Property Count -Sum).Sum)
                errors    = @($results | Where-Object { $_.Failed }).Count
            }
        }

        'open' {
            $export = $state.Export
            if ($null -eq $export) { return @{ ok = $false; error = 'There is no export yet.' } }
            $what = [string](Get-GuiProperty $Data 'what' 'folder')
            $target = if ($what -eq 'report' -and $export.HtmlPath) { $export.HtmlPath } else { $export.Directory }
            try {
                Open-GuiTarget -Target $target
                return @{ ok = $true }
            }
            catch {
                return @{ ok = $false; error = "Could not open '$target': $($_.Exception.Message)" }
            }
        }

        'quit' {
            $state.Running = $false
            return @{ ok = $true }
        }
    }
    return $null
}

function Invoke-GuiRequest {
    param([Parameter(Mandatory)][System.Net.Sockets.TcpClient]$Client)
    $stream = $Client.GetStream()
    $stream.ReadTimeout = 10000
    $stream.WriteTimeout = 10000
    $state = $script:GuiState
    try {
        $request = Read-GuiRequest -Stream $stream
        if ($null -eq $request) { return }

        # Protection against DNS rebinding: only our own Host header is allowed
        $hostHeader = if ($request.Headers.ContainsKey('host')) { $request.Headers['host'] } else { '' }
        if ($hostHeader -notin $state.AllowedHosts) {
            Send-GuiResponse -Stream $stream -Status 403 -ContentType 'text/plain; charset=utf-8' -Body ([System.Text.Encoding]::UTF8.GetBytes('Forbidden'))
            return
        }

        if ($request.Method -eq 'GET' -and ($request.Path -eq '/' -or $request.Path -eq '/report')) {
            if ($request.Query['t'] -ne $state.Token) {
                Send-GuiResponse -Stream $stream -Status 403 -ContentType 'text/plain; charset=utf-8' `
                    -Body ([System.Text.Encoding]::UTF8.GetBytes('Invalid or missing session token. Please use the address shown in the console window.'))
                return
            }
            if ($request.Path -eq '/') {
                $page = $state.Page.Replace('__TOKEN__', $state.Token)
                Send-GuiResponse -Stream $stream -ContentType 'text/html; charset=utf-8' -Body ([System.Text.Encoding]::UTF8.GetBytes($page))
            }
            elseif ($state.Export -and $state.Export.HtmlPath -and (Test-Path -LiteralPath $state.Export.HtmlPath)) {
                Send-GuiResponse -Stream $stream -ContentType 'text/html; charset=utf-8' -Body ([System.IO.File]::ReadAllBytes($state.Export.HtmlPath))
            }
            else {
                Send-GuiResponse -Stream $stream -Status 404 -ContentType 'text/plain; charset=utf-8' -Body ([System.Text.Encoding]::UTF8.GetBytes('No HTML report available.'))
            }
            return
        }

        if ($request.Path -like '/api/*') {
            if ($request.Method -ne 'POST') {
                Send-GuiResponse -Stream $stream -Status 405 -Body (ConvertTo-GuiJson @{ ok = $false; error = 'Method not allowed' })
                return
            }
            $tokenHeader = if ($request.Headers.ContainsKey('x-token')) { $request.Headers['x-token'] } else { '' }
            if ($tokenHeader -ne $state.Token) {
                Send-GuiResponse -Stream $stream -Status 403 -Body (ConvertTo-GuiJson @{ ok = $false; error = 'Forbidden' })
                return
            }
            $data = $null
            if ($request.Body) {
                try { $data = $request.Body | ConvertFrom-Json }
                catch {
                    Send-GuiResponse -Stream $stream -Status 400 -Body (ConvertTo-GuiJson @{ ok = $false; error = 'Invalid JSON' })
                    return
                }
            }
            try {
                $answer = Invoke-GuiApi -Name $request.Path.Substring(5) -Data $data
            }
            catch {
                $answer = @{ ok = $false; error = $_.Exception.Message }
            }
            if ($null -eq $answer) {
                Send-GuiResponse -Stream $stream -Status 404 -Body (ConvertTo-GuiJson @{ ok = $false; error = 'Unknown API call' })
                return
            }
            Send-GuiResponse -Stream $stream -Body (ConvertTo-GuiJson $answer)
            return
        }

        Send-GuiResponse -Stream $stream -Status 404 -ContentType 'text/plain; charset=utf-8' -Body ([System.Text.Encoding]::UTF8.GetBytes('Not Found'))
    }
    catch [System.IO.IOException] {
        # Connection closed by the browser or timeout - ignore
    }
    finally {
        $Client.Close()
    }
}

function Start-SophosGui {
    param(
        [hashtable]$Defaults,
        [int]$ListenPort = 0,
        [switch]$NoBrowser
    )

    $listener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, $ListenPort)
    try {
        $listener.Start()
    }
    catch {
        throw "GUI server could not be started (port $ListenPort): $($_.Exception.Message)"
    }
    $port  = $listener.LocalEndpoint.Port
    $token = [guid]::NewGuid().ToString('N')
    $url   = "http://127.0.0.1:$port/?t=$token"

    $script:GuiState = [pscustomobject]@{
        Token        = $token
        Page         = Get-GuiPage
        AllowedHosts = @("127.0.0.1:$port", "localhost:$port")
        Defaults     = $Defaults
        Running      = $true
        Connected    = $false
        Job          = $null
        Export       = $null
    }

    Write-Host ''
    Write-Host '=============================================================' -ForegroundColor White
    Write-Host "   Sophos Firewall - Configuration Export $($script:Version) (GUI)" -ForegroundColor White
    Write-Host '=============================================================' -ForegroundColor White
    Write-Ok   "Interface started: $url"
    Write-Host 'Please keep this window open. Stop with "Quit" in the interface or Ctrl+C.' -ForegroundColor DarkGray
    Write-Host ''

    if (-not $NoBrowser) {
        try { Open-GuiTarget -Target $url }
        catch { Write-WarnMsg "Browser could not be opened automatically - please open the address above manually." }
    }

    $pending = [System.Collections.Generic.List[object]]::new()
    try {
        while ($script:GuiState.Running) {
            while ($listener.Pending()) {
                [void]$pending.Add([pscustomobject]@{ Client = $listener.AcceptTcpClient(); Since = [datetime]::UtcNow })
            }
            # Browsers often open connections in advance; only connections
            # with incoming data are handled so that an idle connection
            # does not block the server.
            $handled = $false
            foreach ($entry in @($pending)) {
                if ($entry.Client.Available -gt 0) {
                    [void]$pending.Remove($entry)
                    Invoke-GuiRequest -Client $entry.Client
                    $handled = $true
                }
                elseif (([datetime]::UtcNow - $entry.Since).TotalSeconds -gt 30) {
                    [void]$pending.Remove($entry)
                    $entry.Client.Close()
                }
            }
            if (-not $handled) { Start-Sleep -Milliseconds 30 }
        }
    }
    finally {
        foreach ($entry in $pending) { $entry.Client.Close() }
        $listener.Stop()
        $script:ApiPassword = $null
        Write-Info 'GUI stopped.'
    }
}

function Get-GuiPage {
    # HTML/CSS/JavaScript of the interface (self-contained, no external files).
    # __TOKEN__ is replaced with the session token when the page is served.
    return @'
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Sophos Firewall Export</title>
<style>
:root{
  --bg:#eef2f7;--card:#ffffff;--ink:#16202e;--muted:#5d6b7e;--faint:#8a96a6;--line:#d9e0ea;--line2:#e8edf3;
  --brand:#1f3864;--brand2:#2e5496;--accent:#2563eb;--accent-ink:#ffffff;--accent-soft:#e7efff;--accent-line:#9dbcf7;
  --ok:#15803d;--ok-soft:#e6f6ec;--warn:#b45309;--warn-soft:#fdf3e3;--err:#c0262d;--err-soft:#fdecec;
  --field:#f7f9fc;--shadow:0 1px 2px rgba(16,32,64,.06),0 4px 16px rgba(16,32,64,.06);
  --radius:12px;--mono:ui-monospace,SFMono-Regular,Consolas,'Liberation Mono',monospace;
}
@media (prefers-color-scheme:dark){
  :root{
    --bg:#0e141d;--card:#151d29;--ink:#e6ebf2;--muted:#9aa7b8;--faint:#6f7d90;--line:#263243;--line2:#1e2836;
    --brand:#0f1c33;--brand2:#1b3460;--accent:#4c8dff;--accent-ink:#0b1220;--accent-soft:#16264a;--accent-line:#3561b3;
    --ok:#4ade80;--ok-soft:#0f2a1b;--warn:#fbbf24;--warn-soft:#2d2210;--err:#f87171;--err-soft:#301417;
    --field:#101823;--shadow:0 1px 2px rgba(0,0,0,.3),0 4px 16px rgba(0,0,0,.25);
  }
}
*{box-sizing:border-box}
html,body{margin:0}
body{font-family:'Segoe UI',system-ui,-apple-system,Roboto,Arial,sans-serif;font-size:14px;line-height:1.45;color:var(--ink);background:var(--bg);}
button,input,select{font:inherit;color:inherit}
header{background:linear-gradient(135deg,var(--brand),var(--brand2));color:#fff;}
.bar{max-width:1280px;margin:0 auto;padding:18px 24px;display:flex;align-items:center;gap:14px;}
.logo{width:40px;height:40px;border-radius:10px;background:rgba(255,255,255,.14);display:grid;place-items:center;flex:none}
.bar h1{margin:0;font-size:18px;font-weight:600;letter-spacing:.01em}
.bar .sub{font-size:12.5px;opacity:.8}
.ver{margin-left:6px;font-size:11px;font-weight:600;background:rgba(255,255,255,.18);border-radius:20px;padding:2px 8px;vertical-align:2px}
.bar .spacer{flex:1}
.btn-ghost{background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.25);color:#fff;border-radius:8px;padding:7px 14px;cursor:pointer}
.btn-ghost:hover{background:rgba(255,255,255,.2)}
.layout{max-width:1280px;margin:0 auto;padding:22px 24px 40px;display:grid;grid-template-columns:minmax(0,1fr) 380px;gap:22px;align-items:start}
.card{background:var(--card);border:1px solid var(--line);border-radius:var(--radius);box-shadow:var(--shadow);margin-bottom:22px}
.card-h{display:flex;align-items:center;gap:12px;padding:16px 20px;border-bottom:1px solid var(--line2)}
.step{width:26px;height:26px;border-radius:50%;background:var(--accent-soft);color:var(--accent);font-weight:700;font-size:13px;display:grid;place-items:center;flex:none}
.step.is-done{background:var(--ok-soft);color:var(--ok)}
.card-h h2{margin:0;font-size:15px;font-weight:600}
.card-h .hint{color:var(--muted);font-size:12.5px}
.card-h .right{margin-left:auto;display:flex;gap:8px;align-items:center}
.card-b{padding:18px 20px}
.grid{display:grid;grid-template-columns:2fr 1fr;gap:14px 16px}
.grid .full{grid-column:1/-1}
label.f{display:block}
label.f span{display:block;font-size:12.5px;font-weight:600;color:var(--muted);margin-bottom:5px}
.input{width:100%;height:38px;border:1px solid var(--line);background:var(--field);border-radius:8px;padding:0 11px;outline:none;transition:border-color .15s,box-shadow .15s}
.input:focus{border-color:var(--accent);box-shadow:0 0 0 3px var(--accent-soft)}
.pw{position:relative}
.pw .input{padding-right:44px}
.pw button{position:absolute;right:4px;top:4px;height:30px;width:36px;border:0;background:transparent;color:var(--muted);cursor:pointer;border-radius:6px}
.pw button:hover{background:var(--line2)}
.row{display:flex;align-items:center;gap:12px;flex-wrap:wrap}
.btn{text-decoration:none;height:38px;border-radius:8px;border:1px solid var(--line);background:var(--card);padding:0 16px;cursor:pointer;font-weight:600;display:inline-flex;align-items:center;gap:8px;white-space:nowrap}
.btn:hover{background:var(--field)}
.btn.primary{background:var(--accent);border-color:var(--accent);color:var(--accent-ink)}
.btn.primary:hover{filter:brightness(1.07)}
.btn.small{height:30px;padding:0 11px;font-size:12.5px;font-weight:600}
.btn:disabled{opacity:.5;cursor:not-allowed;filter:none}
.btn.big{width:100%;height:46px;justify-content:center;font-size:15px}
.status{display:inline-flex;align-items:center;gap:8px;font-size:13px;color:var(--muted)}
.dot{width:9px;height:9px;border-radius:50%;background:var(--faint);flex:none}
.status.ok{color:var(--ok)} .status.ok .dot{background:var(--ok)}
.status.err{color:var(--err)} .status.err .dot{background:var(--err)}
.status.busy .dot{background:var(--accent);animation:pulse 1s infinite ease-in-out}
@keyframes pulse{50%{opacity:.3}}
.msg{margin-top:14px;padding:10px 12px;border-radius:8px;font-size:13px;display:none}
.msg.show{display:block}
.msg.err{background:var(--err-soft);color:var(--err)}
.msg.warn{background:var(--warn-soft);color:var(--warn)}
.tiles{display:grid;grid-template-columns:repeat(auto-fill,minmax(230px,1fr));gap:10px}
.tile{position:relative;display:flex;gap:11px;align-items:flex-start;border:1px solid var(--line);border-radius:10px;padding:11px 12px;cursor:pointer;background:var(--card);transition:border-color .12s,background .12s;user-select:none}
.tile:hover{border-color:var(--accent-line)}
.tile input{position:absolute;opacity:0;pointer-events:none}
.check{width:18px;height:18px;border-radius:5px;border:1.5px solid var(--faint);flex:none;margin-top:1px;display:grid;place-items:center;transition:all .12s}
.check svg{opacity:0}
.tile.on{border-color:var(--accent);background:var(--accent-soft)}
.tile.on .check{background:var(--accent);border-color:var(--accent);color:var(--accent-ink)}
.tile.on .check svg{opacity:1}
.tile:focus-within{box-shadow:0 0 0 3px var(--accent-soft)}
.tile .t{font-weight:600;font-size:13.5px}
.tile .m{font-family:var(--mono);font-size:11px;color:var(--muted);margin-top:2px;word-break:break-word}
.seg{display:inline-flex;border:1px solid var(--line);border-radius:9px;padding:3px;background:var(--field);gap:3px;flex-wrap:wrap}
.seg button{border:0;background:transparent;border-radius:6px;padding:6px 14px;cursor:pointer;font-weight:600;font-size:13px;color:var(--muted)}
.seg button.on{background:var(--card);color:var(--ink);box-shadow:0 1px 3px rgba(0,0,0,.12)}
.desc{color:var(--muted);font-size:12.5px;margin-top:8px}
.switch{display:flex;align-items:center;gap:10px;cursor:pointer}
.switch input{position:absolute;opacity:0}
.track{width:38px;height:22px;border-radius:20px;background:var(--line);position:relative;transition:background .15s;flex:none}
.track::after{content:'';position:absolute;left:3px;top:3px;width:16px;height:16px;border-radius:50%;background:#fff;transition:left .15s;box-shadow:0 1px 2px rgba(0,0,0,.25)}
.switch input:checked+.track{background:var(--accent)}
.switch input:checked+.track::after{left:19px}
.switch input:focus-visible+.track{box-shadow:0 0 0 3px var(--accent-soft)}
.switch .lbl b{display:block;font-size:13.5px}
.switch .lbl small{color:var(--muted);font-size:12.5px}
.opts{display:grid;gap:18px}
aside{position:sticky;top:20px}
.sum{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-bottom:16px}
.kpi{background:var(--field);border:1px solid var(--line2);border-radius:10px;padding:10px 12px}
.kpi b{display:block;font-size:20px;font-variant-numeric:tabular-nums}
.kpi span{font-size:11.5px;color:var(--muted)}
.target{font-size:12.5px;color:var(--muted);margin-bottom:14px;display:flex;gap:6px;align-items:center;min-height:18px}
.target code{font-family:var(--mono);color:var(--ink);background:var(--field);border-radius:5px;padding:1px 6px}
.progress{margin-top:16px;display:none}
.progress.show{display:block}
.pbar{height:8px;background:var(--line2);border-radius:10px;overflow:hidden}
.pbar i{display:block;height:100%;width:0;background:linear-gradient(90deg,var(--accent),#60a5fa);border-radius:10px;transition:width .25s}
.pinfo{display:flex;justify-content:space-between;font-size:12.5px;color:var(--muted);margin:8px 0 10px}
.log{max-height:340px;overflow:auto;border:1px solid var(--line2);border-radius:10px}
.log div{display:flex;align-items:center;gap:8px;padding:6px 10px;border-top:1px solid var(--line2);font-size:12.5px}
.log div:first-child{border-top:0}
.log .mod{font-family:var(--mono);flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.badge{font-size:11.5px;font-weight:700;border-radius:20px;padding:1px 8px;font-variant-numeric:tabular-nums;white-space:nowrap}
.badge.n{background:var(--ok-soft);color:var(--ok)}
.badge.z{background:var(--line2);color:var(--muted);font-weight:600}
.badge.e{background:var(--err-soft);color:var(--err)}
.badge.w{background:var(--accent-soft);color:var(--accent)}
.done{display:none;margin-top:16px;border:1px solid var(--line);border-radius:10px;padding:14px}
.done.show{display:block}
.done h3{margin:0 0 4px;font-size:15px;display:flex;align-items:center;gap:8px;color:var(--ok)}
.done .path{font-family:var(--mono);font-size:11.5px;color:var(--muted);word-break:break-all;margin-bottom:10px}
.files{list-style:none;margin:0 0 12px;padding:0;font-size:12.5px}
.files li{display:flex;justify-content:space-between;gap:10px;padding:4px 0;border-top:1px dashed var(--line2)}
.files li:first-child{border-top:0}
.files li span:first-child{font-family:var(--mono);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.files li span:last-child{color:var(--faint);white-space:nowrap}
.done .actions{display:flex;gap:8px;flex-wrap:wrap}
.spin{width:15px;height:15px;border:2px solid currentColor;border-right-color:transparent;border-radius:50%;animation:rot .7s linear infinite;display:none}
.busy .spin{display:inline-block}
@keyframes rot{to{transform:rotate(360deg)}}
.foot{color:var(--faint);font-size:12px;margin-top:10px;text-align:center}
.overlay{position:fixed;inset:0;background:var(--bg);display:none;place-items:center;z-index:9;text-align:center;padding:24px}
.overlay.show{display:grid}
.overlay h2{margin:0 0 6px}
.overlay p{color:var(--muted);margin:0}
@media (max-width:980px){.layout{grid-template-columns:minmax(0,1fr)}aside{position:static}}
@media (max-width:560px){.bar{padding:14px 16px}.bar h1{font-size:16px}.layout{padding:16px}.grid{grid-template-columns:minmax(0,1fr)}.bar .sub{display:none}}
</style>
</head>
<body>
<header>
  <div class="bar">
    <div class="logo" aria-hidden="true">
      <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 2l8 3v6c0 5-3.4 9.3-8 11-4.6-1.7-8-6-8-11V5l8-3z"/><path d="M8.5 12l2.5 2.5 4.5-5"/></svg>
    </div>
    <div>
      <h1>Sophos Firewall Configuration Export <span class="ver" id="ver"></span></h1>
      <div class="sub">Export rules and settings via the XML API to Excel, JSON, XML and HTML</div>
    </div>
    <div class="spacer"></div>
    <button class="btn-ghost" id="quit" type="button">Quit</button>
  </div>
</header>

<div class="layout">
  <main>
    <section class="card" id="cardConn">
      <div class="card-h">
        <div class="step" id="step1">1</div>
        <div><h2>Connection</h2><div class="hint">Firewall address and administrator credentials</div></div>
      </div>
      <form class="card-b" id="connForm" autocomplete="off">
        <div class="grid">
          <label class="f"><span>Firewall IP address or hostname</span><input class="input" id="host" placeholder="e.g. 192.168.1.1" required></label>
          <label class="f"><span>Port</span><input class="input" id="port" type="number" min="1" max="65535" value="4444"></label>
          <label class="f"><span>API username</span><input class="input" id="user" placeholder="admin" autocomplete="username"></label>
          <label class="f"><span>API password</span>
            <div class="pw"><input class="input" id="pass" type="password" autocomplete="current-password"><button type="button" id="pwToggle" aria-label="Show password" title="Show password">
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M2 12s3.6-7 10-7 10 7 10 7-3.6 7-10 7S2 12 2 12z"/><circle cx="12" cy="12" r="3"/></svg>
            </button></div>
          </label>
        </div>
        <div class="row" style="margin-top:16px">
          <button class="btn primary" id="connBtn" type="submit"><span class="spin"></span>Test connection</button>
          <div class="status" id="connStatus"><span class="dot"></span><span id="connText">Not connected</span></div>
        </div>
        <div class="msg err" id="connMsg"></div>
      </form>
    </section>

    <section class="card">
      <div class="card-h">
        <div class="step" id="step2">2</div>
        <div><h2>Configuration sections</h2><div class="hint" id="selHint">What should be exported?</div></div>
        <div class="right">
          <button class="btn small" type="button" id="selAll">All</button>
          <button class="btn small" type="button" id="selNone">None</button>
        </div>
      </div>
      <div class="card-b"><div class="tiles" id="tiles"></div></div>
    </section>

    <section class="card">
      <div class="card-h">
        <div class="step" id="step3">3</div>
        <div><h2>Output</h2><div class="hint">File format and location</div></div>
      </div>
      <div class="card-b opts">
        <div>
          <div class="seg" id="fmt" role="radiogroup" aria-label="Output format">
            <button type="button" data-v="Auto" role="radio">Automatic</button>
            <button type="button" data-v="Excel" role="radio">Excel</button>
            <button type="button" data-v="Json" role="radio">JSON</button>
            <button type="button" data-v="Xml" role="radio">XML</button>
          </div>
          <div class="desc" id="fmtDesc"></div>
        </div>
        <label class="switch">
          <input type="checkbox" id="html" checked><span class="track"></span>
          <span class="lbl"><b>Create HTML report</b><small>Clear page with navigation and one table per module</small></span>
        </label>
        <div class="grid">
          <label class="f"><span>Target folder</span><input class="input" id="outDir" placeholder=""></label>
          <label class="f"><span>Timeout per request (s)</span><input class="input" id="timeout" type="number" min="5" max="600" value="60"></label>
        </div>
      </div>
    </section>
  </main>

  <aside>
    <section class="card">
      <div class="card-h"><div><h2>Export</h2><div class="hint">Summary and progress</div></div></div>
      <div class="card-b">
        <div class="target" id="target">No firewall connected yet</div>
        <div class="sum">
          <div class="kpi"><b id="kSec">0</b><span>Sections</span></div>
          <div class="kpi"><b id="kMod">0</b><span>Modules</span></div>
          <div class="kpi"><b id="kRec">&ndash;</b><span>Records</span></div>
        </div>
        <button class="btn primary big" id="runBtn" type="button"><span class="spin"></span><span id="runText">Start export</span></button>
        <div class="msg" id="runMsg"></div>
        <div class="progress" id="progress">
          <div class="pinfo"><span id="pText">Preparing &hellip;</span><span id="pPct">0 %</span></div>
          <div class="pbar"><i id="pBar"></i></div>
          <div class="log" id="log" style="margin-top:12px"></div>
        </div>
        <div class="done" id="done">
          <h3><svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6L9 17l-5-5"/></svg>Export completed</h3>
          <div class="path" id="donePath"></div>
          <ul class="files" id="files"></ul>
          <div class="actions">
            <a class="btn primary small" id="openReport" target="_blank" rel="noopener">View HTML report</a>
            <button class="btn small" type="button" id="openFolder">Open folder</button>
          </div>
        </div>
      </div>
    </section>
    <div class="foot">Runs locally on this computer &middot; credentials are not stored</div>
  </aside>
</div>

<div class="overlay" id="bye"><div><h2>Stopped</h2><p>The interface has been shut down. You can close this browser window.</p></div></div>

<script>
(function(){
'use strict';
var TOKEN='__TOKEN__';
var $=function(id){return document.getElementById(id);};
var state={sections:[],selected:{},format:'Auto',connected:false,running:false,cancel:false,cwd:''};
var FMT={
  Auto:'Excel file with one worksheet per module. If the ImportExcel module is not available, JSON and XML files are created automatically.',
  Excel:'Force an Excel file (requires the PowerShell module ImportExcel; Excel itself does not need to be installed).',
  Json:'One JSON file per module with the processed table data, plus _Overview.json.',
  Xml:'The unmodified XML responses of the API per module, e.g. as a backup.'
};

function api(name,body){
  return fetch('/api/'+name,{method:'POST',headers:{'Content-Type':'application/json','X-Token':TOKEN},body:JSON.stringify(body||{})})
    .then(function(r){return r.json().catch(function(){return {ok:false,error:'Invalid response ('+r.status+')'};});})
    .catch(function(){return {ok:false,error:'No connection to the export script. Is the PowerShell window still running?'};});
}
// Windows PowerShell 5.1 sometimes returns arrays as {value:[...]} or single items without an array
function arr(x){if(Array.isArray(x))return x;if(x&&Array.isArray(x.value))return x.value;return x==null?[]:[x];}
function el(tag,cls,text){var e=document.createElement(tag);if(cls)e.className=cls;if(text!=null)e.textContent=text;return e;}
function show(id,on){$(id).classList.toggle('show',!!on);}
function msg(id,kind,text){var m=$(id);m.className='msg '+(kind||'')+(text?' show':'');m.textContent=text||'';}
function fmtSize(b){if(b<1024)return b+' B';if(b<1048576)return (b/1024).toFixed(1)+' KB';return (b/1048576).toFixed(1)+' MB';}

function renderTiles(){
  var wrap=$('tiles');wrap.textContent='';
  state.sections.forEach(function(s){
    var lab=el('label','tile'+(state.selected[s.key]?' on':''));
    var cb=el('input');cb.type='checkbox';cb.checked=!!state.selected[s.key];
    cb.addEventListener('change',function(){state.selected[s.key]=cb.checked;lab.classList.toggle('on',cb.checked);update();});
    var ck=el('span','check');ck.innerHTML='<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6L9 17l-5-5"/></svg>';
    var txt=el('span');txt.appendChild(el('div','t',s.name));txt.appendChild(el('div','m',arr(s.modules).join(', ')));
    lab.appendChild(cb);lab.appendChild(ck);lab.appendChild(txt);wrap.appendChild(lab);
  });
}
function selectedKeys(){return state.sections.filter(function(s){return state.selected[s.key];}).map(function(s){return s.key;});}
function update(){
  var keys=selectedKeys(),mods=0;
  state.sections.forEach(function(s){if(state.selected[s.key])mods+=arr(s.modules).length;});
  $('kSec').textContent=keys.length;$('kMod').textContent=mods;
  $('selHint').textContent=keys.length===state.sections.length?'All sections selected':keys.length+' of '+state.sections.length+' sections selected';
  $('step2').classList.toggle('is-done',keys.length>0);
  $('runBtn').disabled=!state.running&&keys.length===0;
  var h=$('host').value.trim();
  $('outDir').placeholder=state.cwd?(state.cwd.replace(/[\\\/]$/,'')+(state.cwd.indexOf('\\')>=0?'\\':'/')+'SophosExport_'+(h||'<IP>').replace(/[^\w.\-]/g,'_')+'_<timestamp>'):'';
}
function setFormat(v){
  state.format=v;
  Array.prototype.forEach.call($('fmt').children,function(b){var on=b.getAttribute('data-v')===v;b.classList.toggle('on',on);b.setAttribute('aria-checked',on);});
  $('fmtDesc').textContent=FMT[v];
}
function setConn(kind,text){
  var s=$('connStatus');s.className='status '+(kind||'');$('connText').textContent=text;
  $('step1').classList.toggle('is-done',kind==='ok');
}
function connData(){return {host:$('host').value.trim(),port:parseInt($('port').value,10)||4444,user:$('user').value,password:$('pass').value,timeout:parseInt($('timeout').value,10)||60};}

function connect(){
  var btn=$('connBtn');btn.classList.add('busy');btn.disabled=true;msg('connMsg');setConn('busy','Checking connection \u2026');
  var d=connData();
  return api('connect',d).then(function(r){
    btn.classList.remove('busy');btn.disabled=state.running;
    state.connected=!!r.ok;
    if(r.ok){setConn('ok','Connected to '+r.host+':'+r.port);$('target').innerHTML='';$('target').appendChild(document.createTextNode('Target: '));$('target').appendChild(el('code',null,r.host+':'+r.port));}
    else{setConn('err',r.reachable===false?'Not reachable':'Failed');msg('connMsg','err',r.error);$('target').textContent='No firewall connected yet';}
    return r.ok;
  });
}
['host','port','user','pass'].forEach(function(id){$(id).addEventListener('input',function(){if(state.connected){state.connected=false;setConn('','Changed \u2013 test the connection again');}update();});});
$('timeout').addEventListener('input',function(){state.connected=false;});
$('connForm').addEventListener('submit',function(e){e.preventDefault();connect();});
$('pwToggle').addEventListener('click',function(){var p=$('pass');p.type=p.type==='password'?'text':'password';});
$('selAll').addEventListener('click',function(){state.sections.forEach(function(s){state.selected[s.key]=true;});renderTiles();update();});
$('selNone').addEventListener('click',function(){state.selected={};renderTiles();update();});
Array.prototype.forEach.call($('fmt').children,function(b){b.addEventListener('click',function(){setFormat(b.getAttribute('data-v'));});});

function setRunning(on){
  state.running=on;var b=$('runBtn');b.classList.toggle('busy',on);
  $('runText').textContent=on?'Cancel':'Start export';
  b.classList.toggle('primary',!on);
  ['host','port','user','pass','timeout','outDir','html','connBtn','selAll','selNone'].forEach(function(id){$(id).disabled=on;});
  Array.prototype.forEach.call(document.querySelectorAll('.tile input,#fmt button'),function(x){x.disabled=on;});
  update();
}
function logRow(module,badgeCls,badgeText,title){
  var r=el('div');r.appendChild(el('span','mod',module));var b=el('span','badge '+badgeCls,badgeText);if(title)b.title=title;r.appendChild(b);
  $('log').appendChild(r);r.scrollIntoView({block:'nearest'});return r;
}
function progress(i,total,label){
  var pct=total?Math.round(i/total*100):0;$('pBar').style.width=pct+'%';$('pPct').textContent=pct+' %';$('pText').textContent=label;
}

function run(){
  if(state.running){state.cancel=true;$('runText').textContent='Cancelling \u2026';return;}
  msg('runMsg');show('done',false);
  var keys=selectedKeys();if(!keys.length)return;
  setRunning(true);state.cancel=false;
  var ready=state.connected?Promise.resolve(true):connect();
  ready.then(function(ok){
    if(!ok){setRunning(false);msg('runMsg','err','Export not started: connection to the firewall failed (see step 1).');return;}
    return api('start',{sections:keys,format:state.format,html:$('html').checked,outputDirectory:$('outDir').value.trim()}).then(function(r){
      if(!r.ok){setRunning(false);msg('runMsg','err',r.error);return;}
      $('log').textContent='';show('progress',true);$('kRec').textContent='0';
      var total=r.total,i=0,records=0,errors=0,modules=arr(r.modules);
      function next(){
        if(state.cancel){setRunning(false);progress(i,total,'Cancelled after '+i+' of '+total+' modules');msg('runMsg','warn','Export cancelled. No files have been written.');return;}
        if(i>=total){return finish();}
        var m=modules[i];progress(i,total,'Reading '+m.module+' ('+(i+1)+' of '+total+')');
        var row=logRow(m.module,'w','\u2026');
        api('fetch',{index:i}).then(function(f){
          if(!f.ok){row.lastChild.className='badge e';row.lastChild.textContent='Error';setRunning(false);state.connected=!f.auth&&state.connected;if(f.auth)setConn('err','Login expired');msg('runMsg','err',f.error);return;}
          if(f.failed){errors++;row.lastChild.className='badge e';row.lastChild.textContent='Error';row.lastChild.title=f.status;}
          else if(f.count>0){records+=f.count;row.lastChild.className='badge n';row.lastChild.textContent=f.count;}
          else{row.lastChild.className='badge z';row.lastChild.textContent='empty';row.lastChild.title=f.status||'';}
          $('kRec').textContent=records;i++;next();
        });
      }
      function finish(){
        progress(total,total,'Writing files \u2026');
        return api('finish',{}).then(function(d){
          setRunning(false);
          if(!d.ok){msg('runMsg','err',d.error);progress(total,total,'Error while writing');return;}
          progress(total,total,'Done \u00b7 '+d.records+' records'+(errors?' \u00b7 '+errors+(errors===1?' error':' errors'):''));
          $('donePath').textContent=d.directory;
          var ul=$('files');ul.textContent='';
          arr(d.files).forEach(function(f){var li=el('li');li.appendChild(el('span',null,f.name));li.appendChild(el('span',null,fmtSize(f.size)));ul.appendChild(li);});
          var rep=$('openReport');rep.style.display=d.hasHtml?'':'none';rep.href='/report?t='+TOKEN+'&v='+Date.now();
          show('done',true);
          var notes=arr(d.notes).slice();if(errors)notes.push(errors+(errors===1?' module':' modules')+' could not be read \u2013 hover over \u201cError\u201d for details.');
          if(notes.length)msg('runMsg','warn',notes.join(' '));
          $('done').scrollIntoView({block:'nearest',behavior:'smooth'});
        });
      }
      next();
    });
  });
}
$('runBtn').addEventListener('click',run);
$('openFolder').addEventListener('click',function(){api('open',{what:'folder'}).then(function(r){if(!r.ok)msg('runMsg','err',r.error);});});
$('quit').addEventListener('click',function(){
  if(state.running&&!confirm('An export is still running. Quit anyway?'))return;
  api('quit').then(function(){show('bye',true);});
});

api('info').then(function(r){
  if(!r.ok){msg('connMsg','err',r.error);return;}
  $('ver').textContent='v'+r.version;document.title='Sophos Firewall Export v'+r.version;
  state.sections=arr(r.sections);state.cwd=r.cwd;
  var d=r.defaults||{};
  if(d.host)$('host').value=d.host;if(d.port)$('port').value=d.port;if(d.user)$('user').value=d.user;if(d.timeout)$('timeout').value=d.timeout;
  if(d.outputDirectory)$('outDir').value=d.outputDirectory;$('html').checked=d.html!==false;
  var pre=arr(d.sections);if(!pre.length)pre=state.sections.map(function(s){return s.key;});
  pre.forEach(function(k){state.selected[k]=true;});
  setFormat(FMT[d.format]?d.format:'Auto');renderTiles();update();
  (d.host?(d.user?$('pass'):$('user')):$('host')).focus();
});
})();
</script>
</body>
</html>
'@
}

# ---------------------------------------------------------------------------
# Main program
# ---------------------------------------------------------------------------

if ($ListSections) {
    Write-Host 'Available sections for the -Sections parameter:' -ForegroundColor White
    $script:SectionDefinitions |
        Select-Object @{n = 'Key'; e = { $_.Key } },
                      @{n = 'Description'; e = { $_.Name } },
                      @{n = 'API modules'; e = { $_.Modules -join ', ' } } |
        Format-Table -AutoSize | Out-String -Width 200 | Write-Host
    Write-Host "Additionally: 'All' = export everything"
    return
}

Enable-InsecureTls

# --- GUI mode ---------------------------------------------------------------
if ($Gui) {
    $guiSections = @()
    if ($Sections -and $Sections.Count -gt 0) {
        $guiSections = @(Resolve-SectionsFromParameter -Requested $Sections | ForEach-Object { $_.Key })
    }
    $guiDefaults = @{
        host            = $FirewallIP
        port            = $Port
        user            = $Username
        timeout         = $TimeoutSec
        format          = $OutputFormat
        html            = -not $NoHtml
        outputDirectory = $OutputDirectory
        sections        = $guiSections
    }
    Start-SophosGui -Defaults $guiDefaults -ListenPort $GuiPort -NoBrowser:$NoBrowser
    return
}

Write-Host ''
Write-Host '=============================================================' -ForegroundColor White
Write-Host "   Sophos Firewall - Configuration Export $($script:Version) (XML API)" -ForegroundColor White
Write-Host '=============================================================' -ForegroundColor White
Write-Host ''

# --- 1) Prompt for the firewall IP ------------------------------------------
if ([string]::IsNullOrWhiteSpace($FirewallIP)) {
    for ($attempt = 1; $attempt -le 5; $attempt++) {
        $FirewallIP = Read-Host 'Firewall IP address (e.g. 192.168.1.1)'
        if (-not [string]::IsNullOrWhiteSpace($FirewallIP)) { break }
        Write-WarnMsg 'Please enter an IP address.'
        $FirewallIP = $null
    }
    if ([string]::IsNullOrWhiteSpace($FirewallIP)) { throw 'No IP address specified - aborting.' }
}
$FirewallIP = $FirewallIP.Trim()

$parsedIp = [ref]([System.Net.IPAddress]$null)
if (-not [System.Net.IPAddress]::TryParse($FirewallIP, $parsedIp)) {
    Write-WarnMsg "Note: '$FirewallIP' is not an IP address - it will be used as a hostname."
}

Set-SophosTarget -TargetHost $FirewallIP -TargetPort $Port -Timeout $TimeoutSec

Write-Info "Checking reachability of ${FirewallIP}:${Port} ..."
if (Test-TcpPort -TargetHost $FirewallIP -TargetPort $Port) {
    Write-Ok "Firewall reachable at ${FirewallIP}:${Port}."
}
else {
    Write-WarnMsg "WARNING: ${FirewallIP}:${Port} is not reachable (TCP timeout). Trying anyway."
}

# --- 2) Prompt for API credentials + check login ---------------------------
$maxLoginAttempts = 3
$loginOk = $false

for ($attempt = 1; $attempt -le $maxLoginAttempts; $attempt++) {
    if ([string]::IsNullOrWhiteSpace($Username)) {
        $Username = Read-Host 'API username'
    }
    if ([string]::IsNullOrEmpty($Password)) {
        $Password = Read-PasswordInput -Prompt 'API password'
    }

    Write-Info 'Checking API credentials ...'
    try {
        Test-SophosLogin -User $Username -Pass $Password
        $loginOk = $true
        Write-Ok 'Login successful.'
        break
    }
    catch {
        if ($_.Exception.Message -eq 'AUTH_FAILURE') {
            Write-ErrMsg "Login failed: wrong username or password (attempt $attempt of $maxLoginAttempts)."
            $Username = $null
            $Password = $null
            if ($attempt -eq $maxLoginAttempts) { throw 'Login failed repeatedly - aborting.' }
        }
        else {
            throw "Connection to the firewall failed: $($_.Exception.Message)"
        }
    }
}
if (-not $loginOk) { throw 'Login failed - aborting.' }

# --- 3) Section selection ---------------------------------------------------
if ($Sections -and $Sections.Count -gt 0) {
    $selectedSections = Resolve-SectionsFromParameter -Requested $Sections
}
else {
    $selectedSections = Show-SectionMenu
}
Write-Host ''
Write-Info ("Selected sections: " + (($selectedSections | ForEach-Object { $_.Name }) -join ', '))

# --- 4) Prepare the output directory ----------------------------------------
$timestamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
$OutputDirectory = New-ExportDirectory -OutputDirectory $OutputDirectory -TargetHost $FirewallIP -Timestamp $timestamp
Write-Info "Output directory: $OutputDirectory"

# --- 5) Retrieve data from the firewall --------------------------------------
$allModules = @($selectedSections | ForEach-Object { $_.Modules } | ForEach-Object { $_ })
$results = [System.Collections.Generic.List[object]]::new()
$moduleIndex = 0

foreach ($section in $selectedSections) {
    foreach ($module in $section.Modules) {
        $moduleIndex++
        $percent = [int](($moduleIndex / $allModules.Count) * 100)
        Write-Progress -Activity 'Reading firewall configuration' -Status "$module ($moduleIndex of $($allModules.Count))" -PercentComplete $percent
        try {
            $result = Get-SophosModuleResult -Section $section -Module $module
        }
        catch {
            throw 'Login expired/failed during the export - aborting.'
        }
        [void]$results.Add($result)
        if ($result.Failed) {
            Write-ErrMsg ("  {0,-28} {1}" -f $module, $result.Status)
        }
        elseif ($result.Count -gt 0) {
            Write-Ok ("  {0,-28} {1,4} records" -f $module, $result.Count)
        }
        else {
            Write-Host ("  {0,-28} {1}" -f $module, $result.Status) -ForegroundColor DarkGray
        }
    }
}
Write-Progress -Activity 'Reading firewall configuration' -Completed

# --- 6) Export (Excel or JSON/XML, optional HTML report) ---------------------
$export = Export-SophosResults -Results $results -OutputFormat $OutputFormat -OutputDirectory $OutputDirectory `
    -Timestamp $timestamp -NoHtml:$NoHtml

# --- 7) Summary --------------------------------------------------------------
Write-Host ''
Write-Host '=============================================================' -ForegroundColor White
Write-Host '   Export completed'                                           -ForegroundColor White
Write-Host '=============================================================' -ForegroundColor White
$export.Overview | Format-Table -AutoSize | Out-String -Width 200 | Write-Host
Write-Ok "Created file(s) in: $OutputDirectory"
foreach ($file in $export.Files) {
    Write-Host "  - $file"
}
