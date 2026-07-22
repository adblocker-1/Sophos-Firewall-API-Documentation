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
      5. Zusaetzlich wird immer ein HTML-Report erstellt (ueber das Skript
         Generate-FirewallHtml.ps1 im selben Ordner), abschaltbar mit -NoHtml.

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
$jsonSourceDir = $null   # Verzeichnis mit JSON-Dateien fuer den HTML-Report

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

    if ($wantJson) { $jsonSourceDir = $OutputDirectory }
}

# --- 6b) HTML-Report erzeugen (Generate-FirewallHtml.ps1) ------------------
if (-not $NoHtml) {
    $htmlScript = Join-Path $PSScriptRoot 'Generate-FirewallHtml.ps1'
    if (-not (Test-Path -LiteralPath $htmlScript)) {
        Write-WarnMsg "HTML-Report uebersprungen: '$htmlScript' nicht gefunden."
    }
    else {
        $tempJsonDir = $null
        try {
            $htmlSourceDir = $jsonSourceDir
            if (-not $htmlSourceDir) {
                # Es liegen keine JSON-Dateien im Ausgabeverzeichnis (z.B.
                # reiner Excel- oder XML-Export) - fuer den HTML-Report werden
                # sie temporaer erzeugt und danach wieder geloescht.
                $tempJsonDir = Join-Path ([System.IO.Path]::GetTempPath()) ('SophosHtml_' + [guid]::NewGuid().ToString('N'))
                [void](New-Item -ItemType Directory -Path $tempJsonDir -Force)
                foreach ($result in $results) {
                    if ($result.Anzahl -gt 0) {
                        $tempJsonPath = Join-Path $tempJsonDir "$($result.SectionKey)_$($result.Modul).json"
                        $result.Rows | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $tempJsonPath -Encoding UTF8
                    }
                }
                $overview | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $tempJsonDir '_Uebersicht.json') -Encoding UTF8
                $htmlSourceDir = $tempJsonDir
            }

            $htmlPath = Join-Path $OutputDirectory 'Sophos_Firewall_Konfiguration.html'
            Write-Info 'Erstelle HTML-Report ...'
            & $htmlScript -SourceDir $htmlSourceDir -OutFile $htmlPath `
                -Device $script:TargetHost -ExportDate (Get-Date -Format 'yyyy-MM-dd HH:mm')
            [void]$createdFiles.Add($htmlPath)
        }
        catch {
            Write-WarnMsg "HTML-Report konnte nicht erstellt werden: $($_.Exception.Message)"
        }
        finally {
            if ($tempJsonDir -and (Test-Path -LiteralPath $tempJsonDir)) {
                Remove-Item -LiteralPath $tempJsonDir -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
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
