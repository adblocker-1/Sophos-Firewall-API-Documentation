<#
.SYNOPSIS
    Exportiert die Konfiguration einer Sophos Firewall (SFOS) ueber die XML-API
    in eine Excel-Datei (ein Tabellenblatt pro Bereich) oder alternativ in
    JSON-/XML-Dateien pro Bereich.

.DESCRIPTION
    Das Skript verbindet sich per HTTPS mit der XML-API einer lokalen Sophos
    Firewall (Standard-Port 4444, Endpunkt /webconsole/APIController).

    Ablauf:
      1. Abfrage der Firewall-IP-Adresse
      2. Abfrage der API-Zugangsdaten (Benutzername + Passwort)
      3. Auswahlmenue: bestimmte Konfigurationsbereiche oder alles exportieren
      4. Export nach Excel (Modul "ImportExcel", ein Blatt pro Bereich).
         Ist Excel-Export nicht moeglich, werden automatisch JSON- und
         XML-Dateien pro Bereich erstellt.
      5. Zusaetzlich wird immer ein HTML-Report erstellt
         (Sophos_Firewall_Konfiguration.html), abschaltbar mit -NoHtml.

    Voraussetzung auf der Firewall:
      - API-Zugriff aktivieren unter: Backup & Firmware > API
      - Die IP-Adresse des Rechners, der dieses Skript ausfuehrt, muss dort
        als erlaubte API-IP eingetragen sein.

.PARAMETER FirewallIP
    IP-Adresse (oder Hostname) der Sophos Firewall. Ohne Angabe wird
    interaktiv gefragt.

.PARAMETER Port
    HTTPS-Port der Web-Admin-Konsole / API (Standard: 4444).

.PARAMETER Username
    API-Benutzername. Ohne Angabe wird interaktiv gefragt.

.PARAMETER Password
    API-Passwort im Klartext (nur fuer Automatisierung gedacht!).
    Ohne Angabe wird interaktiv und verdeckt gefragt.

.PARAMETER Sections
    Liste der zu exportierenden Bereiche (Schluesselnamen, siehe
    -ListSections) oder 'All' fuer alles. Ohne Angabe erscheint ein
    interaktives Auswahlmenue.

.PARAMETER OutputFormat
    Auto  = Excel, bei Problemen automatisch JSON+XML (Standard)
    Excel = Excel-Datei erzwingen
    Json  = eine JSON-Datei pro Bereich
    Xml   = eine XML-Rohdatei (API-Antwort) pro Bereich

.PARAMETER OutputDirectory
    Zielverzeichnis fuer den Export. Standard: Unterordner
    "SophosExport_<IP>_<Zeitstempel>" im aktuellen Verzeichnis.

.PARAMETER NoHtml
    Unterdrueckt die zusaetzliche Erstellung des HTML-Reports
    (Sophos_Firewall_Konfiguration.html).

.PARAMETER ListSections
    Zeigt nur die verfuegbaren Bereichs-Schluessel an und beendet das Skript.

.PARAMETER TimeoutSec
    Timeout pro API-Anfrage in Sekunden (Standard: 60).

.EXAMPLE
    .\Export-SophosFirewallConfig.ps1
    Vollstaendig interaktiv: fragt IP, Zugangsdaten und Bereiche ab.

.EXAMPLE
    .\Export-SophosFirewallConfig.ps1 -FirewallIP 192.168.1.1 -Username apiadmin -Password 'Geheim123!' -Sections All
    Nicht-interaktiv: exportiert alle Bereiche nach Excel.

.EXAMPLE
    .\Export-SophosFirewallConfig.ps1 -FirewallIP 192.168.1.1 -Username apiadmin -Password 'Geheim123!' -Sections FirewallRules,Hosts -OutputFormat Json

.NOTES
    Getestet mit PowerShell 5.1 und PowerShell 7.x.
    Die Verbindung akzeptiert selbstsignierte Zertifikate der Firewall.
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
    [int]$TimeoutSec = 60
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Definition der exportierbaren Konfigurationsbereiche.
# "Module" sind die XML-API-Entitaeten der Sophos Firewall (Get-Requests).
# ---------------------------------------------------------------------------
$script:SectionDefinitions = @(
    [pscustomobject]@{ Nr =  1; Key = 'FirewallRules';   Anzeige = 'Firewall-Regeln';          Module = @('FirewallRule', 'FirewallRuleGroup') }
    [pscustomobject]@{ Nr =  2; Key = 'NatRules';        Anzeige = 'NAT-Regeln';               Module = @('NATRule') }
    [pscustomobject]@{ Nr =  3; Key = 'Hosts';           Anzeige = 'Netzwerk-Objekte (Hosts)'; Module = @('IPHost', 'IPHostGroup', 'FQDNHost', 'FQDNHostGroup', 'MACHost') }
    [pscustomobject]@{ Nr =  4; Key = 'Services';        Anzeige = 'Dienste (Services)';       Module = @('Services', 'ServiceGroup') }
    [pscustomobject]@{ Nr =  5; Key = 'Zones';           Anzeige = 'Zonen';                    Module = @('Zone') }
    [pscustomobject]@{ Nr =  6; Key = 'Interfaces';      Anzeige = 'Interfaces und Netzwerk';  Module = @('Interface', 'VLAN', 'LAG', 'BridgePair') }
    [pscustomobject]@{ Nr =  7; Key = 'Routing';         Anzeige = 'Routing (statische Routen)'; Module = @('UnicastRoute') }
    [pscustomobject]@{ Nr =  8; Key = 'DnsDhcp';         Anzeige = 'DNS und DHCP';             Module = @('DNS', 'DHCPServer') }
    [pscustomobject]@{ Nr =  9; Key = 'Users';           Anzeige = 'Benutzer und Gruppen';     Module = @('User', 'Group') }
    [pscustomobject]@{ Nr = 10; Key = 'WebFilter';       Anzeige = 'Webfilter';                Module = @('WebFilterPolicy', 'WebFilterURLGroup') }
    [pscustomobject]@{ Nr = 11; Key = 'AppFilter';       Anzeige = 'Applikationsfilter';       Module = @('ApplicationFilterPolicy') }
    [pscustomobject]@{ Nr = 12; Key = 'IPS';             Anzeige = 'IPS-Richtlinien';          Module = @('IPSPolicy') }
    [pscustomobject]@{ Nr = 13; Key = 'VPN';             Anzeige = 'VPN (IPsec / SSL)';        Module = @('VPNIPSecConnection', 'SSLVPNPolicy') }
    [pscustomobject]@{ Nr = 14; Key = 'Wireless';        Anzeige = 'WLAN';                     Module = @('WirelessNetworks', 'AccessPoint') }
    [pscustomobject]@{ Nr = 15; Key = 'Certificates';    Anzeige = 'Zertifikate';              Module = @('Certificate', 'CertificateAuthority') }
    [pscustomobject]@{ Nr = 16; Key = 'Schedules';       Anzeige = 'Zeitplaene';               Module = @('Schedule') }
    [pscustomobject]@{ Nr = 17; Key = 'System';          Anzeige = 'System und Verwaltung';    Module = @('AdminSettings', 'Notification', 'SyslogServers', 'SNMPAgentConfiguration', 'BackupRestore') }
)

$script:IsCoreEdition = $PSVersionTable.PSEdition -eq 'Core'

# ---------------------------------------------------------------------------
# Hilfsfunktionen
# ---------------------------------------------------------------------------

function Write-Info    { param([string]$Text) Write-Host $Text -ForegroundColor Cyan }
function Write-Ok      { param([string]$Text) Write-Host $Text -ForegroundColor Green }
function Write-WarnMsg { param([string]$Text) Write-Host $Text -ForegroundColor Yellow }
function Write-ErrMsg  { param([string]$Text) Write-Host $Text -ForegroundColor Red }

function Enable-InsecureTls {
    # Sophos Firewalls verwenden meist selbstsignierte Zertifikate.
    # Unter Windows PowerShell 5.1 muss die Zertifikatspruefung global
    # deaktiviert werden; PowerShell 7+ nutzt stattdessen -SkipCertificateCheck.
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
    # Verdeckte Passworteingabe. Bei umgeleiteter Eingabe (z.B. Pipe/Skript)
    # unterstuetzt die Konsole keine SecureString-Eingabe - dann Klartext lesen.
    param([string]$Prompt)
    if ([System.Console]::IsInputRedirected) {
        return (Read-Host $Prompt)
    }
    $secure = Read-Host $Prompt -AsSecureString
    return (ConvertFrom-SecureStringPlain -SecureString $secure)
}

function Invoke-SophosApiRequest {
    <#
        Sendet einen Get-Request fuer eine API-Entitaet und liefert ein Objekt
        mit RawXml (Antwort als Text), Records (XmlElement-Liste) und
        StatusText (Hinweis, z.B. "keine Datensaetze") zurueck.
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
        throw "Antwort der Firewall ist kein gueltiges XML (Entitaet '$EntityName'): $($_.Exception.Message)"
    }

    # Login-Pruefung: bei falschen Zugangsdaten liefert die API
    # <Login>Authentication Failure</Login> zurueck.
    $loginNode = $xml.SelectSingleNode('/Response/Login')
    if ($loginNode -and $loginNode.InnerText -match 'Authentication Failure') {
        throw 'AUTH_FAILURE'
    }

    $records = @($xml.SelectNodes("/Response/$EntityName"))
    $statusText = $null

    # Ein einzelner Datensatz, der nur ein <Status>-Element enthaelt, ist eine
    # Statusmeldung der API (z.B. Code 526 = "Record does not exist").
    if ($records.Count -eq 1) {
        $elementChildren = @($records[0].ChildNodes | Where-Object { $_.NodeType -eq [System.Xml.XmlNodeType]::Element })
        if ($elementChildren.Count -eq 1 -and $elementChildren[0].Name -eq 'Status') {
            $code = $elementChildren[0].GetAttribute('code')
            $text = $elementChildren[0].InnerText
            $statusText = if ($code) { "Status $code - $text" } else { $text }
            $records = @()
        }
    }

    # Globale Statusmeldung der API (z.B. "API operation not permitted from ...")
    if ($records.Count -eq 0 -and -not $statusText) {
        $globalStatus = $xml.SelectSingleNode('/Response/Status')
        if ($globalStatus) {
            $text = $globalStatus.InnerText
            if ($text -match 'not permitted|No\. of') {
                throw "API-Zugriff verweigert: $text (IP dieses Rechners unter 'Backup & Firmware > API' freischalten!)"
            }
            $statusText = $text
        }
        else {
            $statusText = 'Keine Datensaetze vorhanden'
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
    # Wandelt ein (verschachteltes) XML-Element rekursiv in flache
    # Spalte=Wert-Paare um. Verschachtelte Ebenen werden mit '.' verbunden,
    # Listen (mehrfach gleiche Elemente) mit ', ' zusammengefasst.
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
    # Wandelt die XML-Datensaetze eines Moduls in Tabellenzeilen um und
    # sorgt dafuer, dass alle Zeilen dieselben Spalten besitzen.
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
    # Excel-Blattnamen: max. 31 Zeichen, keine Zeichen aus []:*?/\
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
    # Prueft, ob das Modul "ImportExcel" verfuegbar ist, und versucht sonst,
    # es aus der PowerShell Gallery zu installieren. Liefert $true/$false.
    if (Get-Module -Name ImportExcel) { return $true }
    if (Get-Module -ListAvailable -Name ImportExcel) {
        Import-Module ImportExcel -ErrorAction Stop
        return $true
    }
    Write-Info 'Modul "ImportExcel" nicht gefunden - versuche Installation aus der PowerShell Gallery ...'
    try {
        Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Import-Module ImportExcel -ErrorAction Stop
        Write-Ok 'Modul "ImportExcel" wurde installiert.'
        return $true
    }
    catch {
        Write-WarnMsg "Installation von ImportExcel fehlgeschlagen: $($_.Exception.Message)"
        return $false
    }
}

# ---------------------------------------------------------------------------
# HTML-Report (integriert aus Generate-FirewallHtml.ps1)
# ---------------------------------------------------------------------------
$script:HtmlMaxLen = 800

function Convert-HtmlSanitize {
    param([string]$s)
    if ($null -eq $s) { return '' }
    $s = [regex]::Replace($s, "[\x00-\x08\x0b\x0c\x0e-\x1f]", ' ')
    $s = [regex]::Replace($s, "[ \t]{3,}", '  ')
    $s = $s.Trim()
    if ($s.Length -gt $script:HtmlMaxLen) { $s = $s.Substring(0, $script:HtmlMaxLen) + ' ... [gekuerzt]' }
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

function Get-HtmlTable {
    # Bereitet die Tabellenzeilen eines Moduls fuer den HTML-Report auf:
    # Positions-Artefaktspalten (Endung ".Name") und komplett leere Spalten
    # werden entfernt, Werte bereinigt und gekuerzt.
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
    # Komplett leere Spalten entfernen
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
    # Erzeugt aus den bereits abgerufenen Modulergebnissen eine HTML-Seite
    # mit Navigation, Zusammenfassung und je einem Abschnitt pro Modul.
    param(
        [Parameter(Mandatory)]$Results,
        [Parameter(Mandatory)][string]$Device,
        [Parameter(Mandatory)][string]$ExportDate,
        [Parameter(Mandatory)][string]$OutFile
    )

    $sections = @()
    foreach ($r in $Results) {
        $flat = Get-HtmlTable @($r.Rows)
        $sections += [PSCustomObject]@{
            Bereich = Convert-HtmlValue $r.Bereich
            Modul   = Convert-HtmlValue $r.Modul
            Count   = $r.Anzahl
            Status  = Convert-HtmlValue $r.Status
            Cols    = @($flat.Cols)
            Rows    = @($flat.Rows)
        }
    }
    $today = (Get-Date).ToString('yyyy-MM-dd')

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append(@"
<!DOCTYPE html><html lang="de"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Sophos Firewall &#8211; Konfiguration $(Get-HtmlEscape $Device)</title>
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
<h1>Sophos Firewall &#8211; Konfigurations&#252;bersicht</h1>
<div class="meta">Ger&#228;t $(Get-HtmlEscape $Device) &nbsp;&#183;&nbsp; Export $(Get-HtmlEscape $ExportDate) &nbsp;&#183;&nbsp; erstellt $today</div>
</header>
<div class="wrap">
<nav><h2>Inhalt</h2>
"@)

    # Navigation
    for ($i = 0; $i -lt $sections.Count; $i++) {
        $s = $sections[$i]
        $cls = ''
        if ($s.Count -eq 0) { $cls = ' class=empty' }
        [void]$sb.Append(('<a href="#s{0}"{1}><span>{2}</span><span class="cnt">{3}</span></a>' -f $i, $cls, (Get-HtmlEscape $s.Modul), $s.Count))
    }
    [void]$sb.Append('</nav><main>')

    # Zusammenfassung
    [void]$sb.Append('<section id="overview"><h2>Zusammenfassung</h2><div class="cat">Alle Konfigurationsbereiche</div>')
    [void]$sb.Append('<table class="summary"><thead><tr><th>Bereich</th><th>Modul</th><th>Datens&#228;tze</th><th>Status</th></tr></thead><tbody>')
    $last = $null
    foreach ($s in $sections) {
        $zc = ''
        if ($s.Count -eq 0) { $zc = ' class=zero' }
        $bshow = ''
        if ($s.Bereich -ne $last) { $bshow = (Get-HtmlEscape $s.Bereich) }
        $last = $s.Bereich
        if ($s.Count -eq 0) { $st = 'leer' } else { $st = (Get-HtmlEscape $s.Status) }
        [void]$sb.Append(('<tr{0}><td>{1}</td><td>{2}</td><td class="n">{3}</td><td>{4}</td></tr>' -f $zc, $bshow, (Get-HtmlEscape $s.Modul), $s.Count, $st))
    }
    [void]$sb.Append('</tbody></table></section>')

    # Abschnitte
    for ($i = 0; $i -lt $sections.Count; $i++) {
        $s = $sections[$i]
        [void]$sb.Append(('<section id="s{0}"><h2>{1} <span class="pill">{2}</span></h2><div class="cat">{3}</div>' -f $i, (Get-HtmlEscape $s.Modul), $s.Count, (Get-HtmlEscape $s.Bereich)))
        if ($s.Count -eq 0 -or $s.Cols.Count -eq 0) {
            [void]$sb.Append('<p class="empty-note">Keine Eintr&#228;ge vorhanden.</p>')
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

    [void]$sb.Append('</main></div><footer>Erstellt aus dem Sophos-Export &#183; nur zur internen Dokumentation</footer></body></html>')

    # Schreiben (UTF-8 ohne BOM)
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
    # Zeigt das Auswahlmenue an und liefert die gewaehlten Bereichs-Objekte.
    Write-Host ''
    Write-Host '=============================================================' -ForegroundColor White
    Write-Host '  Welche Konfigurationsbereiche sollen exportiert werden?'    -ForegroundColor White
    Write-Host '=============================================================' -ForegroundColor White
    foreach ($section in $script:SectionDefinitions) {
        $line = '  [{0,2}] {1,-30} ({2})' -f $section.Nr, $section.Anzeige, ($section.Module -join ', ')
        Write-Host $line
    }
    Write-Host '  [ A] ALLES exportieren' -ForegroundColor Green
    Write-Host '  [ Q] Abbrechen'         -ForegroundColor Yellow
    Write-Host ''
    Write-Host 'Mehrfachauswahl mit Komma und Bereichen moeglich, z.B.: 1,3,5-8' -ForegroundColor DarkGray

    for ($attempt = 1; $attempt -le 5; $attempt++) {
        $answer = Read-Host 'Auswahl'
        if ([string]::IsNullOrWhiteSpace($answer)) { continue }
        $answer = $answer.Trim()

        if ($answer -match '^[Qq]$') {
            Write-WarnMsg 'Abgebrochen.'
            exit 0
        }
        if ($answer -match '^([Aa](lle[s]?)?|0)$') {
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
                Write-WarnMsg "Ungueltige Nummer(n): $($unknown -join ', '). Bitte erneut."
                continue
            }
            if ($selected.Count -gt 0) { return $selected }
        }
        Write-WarnMsg 'Eingabe nicht verstanden. Beispiele: "1,3", "5-8", "A" fuer alles.'
    }
    throw 'Zu viele ungueltige Eingaben - Abbruch.'
}

function Resolve-SectionsFromParameter {
    param([string[]]$Requested)
    # Auch "A,B" als ein einzelnes Argument zulassen (z.B. bei pwsh -File)
    $Requested = @($Requested | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    if ($Requested -contains 'All' -or $Requested -contains 'Alle' -or $Requested -contains 'Alles') {
        return @($script:SectionDefinitions)
    }
    $resolved = [System.Collections.Generic.List[object]]::new()
    foreach ($name in $Requested) {
        $match = $script:SectionDefinitions | Where-Object { $_.Key -eq $name }
        if (-not $match) {
            $validKeys = ($script:SectionDefinitions.Key -join ', ')
            throw "Unbekannter Bereich '$name'. Gueltige Werte: All, $validKeys"
        }
        [void]$resolved.Add($match)
    }
    return @($resolved)
}

# ---------------------------------------------------------------------------
# Hauptprogramm
# ---------------------------------------------------------------------------

if ($ListSections) {
    Write-Host 'Verfuegbare Bereiche fuer den Parameter -Sections:' -ForegroundColor White
    $script:SectionDefinitions |
        Select-Object @{n = 'Key'; e = { $_.Key } },
                      @{n = 'Beschreibung'; e = { $_.Anzeige } },
                      @{n = 'API-Module'; e = { $_.Module -join ', ' } } |
        Format-Table -AutoSize | Out-String -Width 200 | Write-Host
    Write-Host "Zusaetzlich: 'All' = alles exportieren"
    return
}

Write-Host ''
Write-Host '=============================================================' -ForegroundColor White
Write-Host '   Sophos Firewall - Konfigurations-Export (XML-API)'          -ForegroundColor White
Write-Host '=============================================================' -ForegroundColor White
Write-Host ''

Enable-InsecureTls

# --- 1) Firewall-IP abfragen -----------------------------------------------
if ([string]::IsNullOrWhiteSpace($FirewallIP)) {
    for ($attempt = 1; $attempt -le 5; $attempt++) {
        $FirewallIP = Read-Host 'Firewall IP-Adresse (z.B. 192.168.1.1)'
        if (-not [string]::IsNullOrWhiteSpace($FirewallIP)) { break }
        Write-WarnMsg 'Bitte eine IP-Adresse eingeben.'
        $FirewallIP = $null
    }
    if ([string]::IsNullOrWhiteSpace($FirewallIP)) { throw 'Keine IP-Adresse angegeben - Abbruch.' }
}
$FirewallIP = $FirewallIP.Trim()

$parsedIp = [ref]([System.Net.IPAddress]$null)
if (-not [System.Net.IPAddress]::TryParse($FirewallIP, $parsedIp)) {
    Write-WarnMsg "Hinweis: '$FirewallIP' ist keine IP-Adresse - wird als Hostname verwendet."
}

$script:TargetHost     = $FirewallIP
$script:TargetPort     = $Port
$script:RequestTimeout = $TimeoutSec

Write-Info "Pruefe Erreichbarkeit von ${FirewallIP}:${Port} ..."
if (Test-TcpPort -TargetHost $FirewallIP -TargetPort $Port) {
    Write-Ok "Firewall unter ${FirewallIP}:${Port} erreichbar."
}
else {
    Write-WarnMsg "WARNUNG: ${FirewallIP}:${Port} ist nicht erreichbar (TCP-Timeout). Es wird trotzdem versucht."
}

# --- 2) API-Zugangsdaten abfragen + Login pruefen --------------------------
$maxLoginAttempts = 3
$loginOk = $false

for ($attempt = 1; $attempt -le $maxLoginAttempts; $attempt++) {
    if ([string]::IsNullOrWhiteSpace($Username)) {
        $Username = Read-Host 'API-Benutzername'
    }
    if ([string]::IsNullOrEmpty($Password)) {
        $Password = Read-PasswordInput -Prompt 'API-Passwort'
    }

    $script:ApiUser     = $Username
    $script:ApiPassword = $Password

    Write-Info 'Pruefe API-Zugangsdaten ...'
    try {
        [void](Invoke-SophosApiRequest -EntityName 'Zone')
        $loginOk = $true
        Write-Ok 'Anmeldung erfolgreich.'
        break
    }
    catch {
        if ($_.Exception.Message -eq 'AUTH_FAILURE') {
            Write-ErrMsg "Anmeldung fehlgeschlagen: Benutzername oder Passwort falsch (Versuch $attempt von $maxLoginAttempts)."
            $Username = $null
            $Password = $null
            if ($attempt -eq $maxLoginAttempts) { throw 'Anmeldung mehrfach fehlgeschlagen - Abbruch.' }
        }
        else {
            throw "Verbindung zur Firewall fehlgeschlagen: $($_.Exception.Message)"
        }
    }
}
if (-not $loginOk) { throw 'Anmeldung fehlgeschlagen - Abbruch.' }

# --- 3) Bereichsauswahl ----------------------------------------------------
if ($Sections -and $Sections.Count -gt 0) {
    $selectedSections = Resolve-SectionsFromParameter -Requested $Sections
}
else {
    $selectedSections = Show-SectionMenu
}
Write-Host ''
Write-Info ("Ausgewaehlte Bereiche: " + (($selectedSections | ForEach-Object { $_.Anzeige }) -join ', '))

# --- 4) Ausgabeverzeichnis vorbereiten -------------------------------------
$timestamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $safeHost = $FirewallIP -replace '[^\w\.\-]', '_'
    $OutputDirectory = Join-Path -Path (Get-Location).Path -ChildPath "SophosExport_${safeHost}_$timestamp"
}
if (-not (Test-Path -LiteralPath $OutputDirectory)) {
    [void](New-Item -ItemType Directory -Path $OutputDirectory -Force)
}
Write-Info "Ausgabeverzeichnis: $OutputDirectory"

# --- 5) Daten von der Firewall abrufen -------------------------------------
$allModules = @($selectedSections | ForEach-Object { $_.Module } | ForEach-Object { $_ })
$results = [System.Collections.Generic.List[object]]::new()
$moduleIndex = 0

foreach ($section in $selectedSections) {
    foreach ($module in $section.Module) {
        $moduleIndex++
        $percent = [int](($moduleIndex / $allModules.Count) * 100)
        Write-Progress -Activity 'Lese Firewall-Konfiguration' -Status "$module ($moduleIndex von $($allModules.Count))" -PercentComplete $percent
        try {
            $apiResult = Invoke-SophosApiRequest -EntityName $module
            $rows = ConvertTo-FlatRows -Records $apiResult.Records
            $status = if ($apiResult.Records.Count -gt 0) { 'OK' } else { $apiResult.StatusText }
            [void]$results.Add([pscustomobject]@{
                Bereich    = $section.Anzeige
                SectionKey = $section.Key
                Modul      = $module
                Anzahl     = $apiResult.Records.Count
                Status     = $status
                Rows       = $rows
                RawXml     = $apiResult.RawXml
            })
            if ($apiResult.Records.Count -gt 0) {
                Write-Ok ("  {0,-28} {1,4} Datensaetze" -f $module, $apiResult.Records.Count)
            }
            else {
                Write-Host ("  {0,-28} {1}" -f $module, $status) -ForegroundColor DarkGray
            }
        }
        catch {
            if ($_.Exception.Message -eq 'AUTH_FAILURE') { throw 'Anmeldung waehrend des Exports abgelaufen/fehlgeschlagen - Abbruch.' }
            Write-ErrMsg ("  {0,-28} FEHLER: {1}" -f $module, $_.Exception.Message)
            [void]$results.Add([pscustomobject]@{
                Bereich    = $section.Anzeige
                SectionKey = $section.Key
                Modul      = $module
                Anzahl     = 0
                Status     = "FEHLER: $($_.Exception.Message)"
                Rows       = @()
                RawXml     = $null
            })
        }
    }
}
Write-Progress -Activity 'Lese Firewall-Konfiguration' -Completed

# --- 6) Export -------------------------------------------------------------
$overview = $results | Select-Object `
    @{n = 'Bereich'; e = { $_.Bereich } },
    @{n = 'Modul'; e = { $_.Modul } },
    @{n = 'Datensaetze'; e = { $_.Anzahl } },
    @{n = 'Status'; e = { $_.Status } }

$useExcel = $false
if ($OutputFormat -in @('Auto', 'Excel')) {
    $useExcel = Initialize-ImportExcel
    if (-not $useExcel) {
        if ($OutputFormat -eq 'Excel') {
            throw 'Excel-Export nicht moeglich (Modul "ImportExcel" fehlt). Bitte manuell installieren: Install-Module ImportExcel'
        }
        Write-WarnMsg 'Excel-Export nicht moeglich - Fallback auf JSON- und XML-Dateien pro Bereich.'
    }
}

$createdFiles = [System.Collections.Generic.List[string]]::new()

if ($useExcel) {
    $excelPath = Join-Path $OutputDirectory "Sophos_Firewall_Export_$timestamp.xlsx"
    $usedSheetNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    $exportExcel = {
        param($Data, $SheetName)
        try {
            $Data | Export-Excel -Path $excelPath -WorksheetName $SheetName -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter
        }
        catch {
            # -AutoSize kann auf Systemen ohne Grafikbibliothek fehlschlagen
            $Data | Export-Excel -Path $excelPath -WorksheetName $SheetName -FreezeTopRow -BoldTopRow -AutoFilter
        }
    }

    $overviewSheet = Get-SafeSheetName -Name 'Uebersicht' -UsedNames $usedSheetNames
    & $exportExcel $overview $overviewSheet

    foreach ($result in $results) {
        if ($result.Anzahl -eq 0) { continue }
        $sheetName = Get-SafeSheetName -Name $result.Modul -UsedNames $usedSheetNames
        & $exportExcel $result.Rows $sheetName
    }
    [void]$createdFiles.Add($excelPath)
}
else {
    $wantJson = $OutputFormat -in @('Auto', 'Json')
    $wantXml  = $OutputFormat -in @('Auto', 'Xml')

    foreach ($result in $results) {
        $baseName = "$($result.SectionKey)_$($result.Modul)"
        if ($wantJson -and $result.Anzahl -gt 0) {
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

    $summaryPath = Join-Path $OutputDirectory '_Uebersicht.json'
    $overview | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $summaryPath -Encoding UTF8
    [void]$createdFiles.Add($summaryPath)
}

# --- 6b) HTML-Report erzeugen ----------------------------------------------
if (-not $NoHtml) {
    try {
        $htmlPath = Join-Path $OutputDirectory 'Sophos_Firewall_Konfiguration.html'
        Write-Info 'Erstelle HTML-Report ...'
        New-HtmlReport -Results $results -OutFile $htmlPath `
            -Device $script:TargetHost -ExportDate (Get-Date -Format 'yyyy-MM-dd HH:mm')
        Write-Ok "HTML erstellt: $htmlPath"
        [void]$createdFiles.Add($htmlPath)
    }
    catch {
        Write-WarnMsg "HTML-Report konnte nicht erstellt werden: $($_.Exception.Message)"
    }
}

# --- 7) Zusammenfassung ----------------------------------------------------
Write-Host ''
Write-Host '=============================================================' -ForegroundColor White
Write-Host '   Export abgeschlossen'                                       -ForegroundColor White
Write-Host '=============================================================' -ForegroundColor White
$overview | Format-Table -AutoSize | Out-String -Width 200 | Write-Host
Write-Ok "Erstellte Datei(en) in: $OutputDirectory"
foreach ($file in $createdFiles) {
    Write-Host "  - $file"
}
