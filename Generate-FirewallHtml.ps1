<#
.SYNOPSIS
    Erzeugt aus einem Sophos-Firewall-Export (JSON-Dateien) eine
    HTML-Konfigurationsuebersicht (Sophos_Firewall_Konfiguration.html).

.DESCRIPTION
    Liest _Uebersicht.json als Inhaltsverzeichnis und laedt zu jedem Modul
    die passende Datei "*_<Modul>.json". Die Datensaetze werden zu Tabellen
    aufbereitet (Positions-Artefaktspalten ".Name" und komplett leere Spalten
    werden entfernt, Steuerzeichen bereinigt, sehr lange Werte gekuerzt) und
    als einzelne HTML-Seite mit Navigation, Zusammenfassung und je einem
    Abschnitt pro Modul ausgegeben.

.PARAMETER SourceDir
    Ordner mit den Export-JSON-Dateien. Standard: Ordner des Skripts.

.PARAMETER OutFile
    Zielpfad der HTML-Datei. Standard: <SourceDir>\Sophos_Firewall_Konfiguration.html

.PARAMETER Device
    Anzeigename/IP der Firewall fuer die Kopfzeile. Standard: wird aus dem
    Ordnernamen (SophosExport_<IP>_<Zeitstempel>) abgeleitet.

.PARAMETER ExportDate
    Anzeigedatum des Exports fuer die Kopfzeile. Standard: wird aus dem
    Ordnernamen abgeleitet.

.EXAMPLE
    .\Generate-FirewallHtml.ps1
    .\Generate-FirewallHtml.ps1 -SourceDir "C:\Export" -OutFile "C:\Export\report.html"
    .\Generate-FirewallHtml.ps1 -SourceDir "C:\Export" -Device 192.168.1.1 -ExportDate "2026-07-21 22:03"
#>
[CmdletBinding()]
param(
    [string]$SourceDir,
    [string]$OutFile,
    [string]$Device,
    [string]$ExportDate
)

# ---------- Pfade ----------
if (-not $SourceDir) {
    if ($PSScriptRoot) { $SourceDir = $PSScriptRoot } else { $SourceDir = (Get-Location).Path }
}
if (-not (Test-Path -LiteralPath $SourceDir)) { throw "Ordner nicht gefunden: $SourceDir" }
if (-not $OutFile)  { $OutFile = Join-Path $SourceDir 'Sophos_Firewall_Konfiguration.html' }

$script:MaxLen = 800

# ---------- Hilfsfunktionen ----------
function Convert-Sanitize {
    param([string]$s)
    if ($null -eq $s) { return '' }
    $s = [regex]::Replace($s, "[\x00-\x08\x0b\x0c\x0e-\x1f]", ' ')
    $s = [regex]::Replace($s, "[ \t]{3,}", '  ')
    $s = $s.Trim()
    if ($s.Length -gt $script:MaxLen) { $s = $s.Substring(0, $script:MaxLen) + ' … [gekürzt]' }
    return $s
}

function Convert-Value {
    param($v)
    if ($null -eq $v) { return '' }
    if ($v -is [System.Collections.IEnumerable] -and -not ($v -is [string])) {
        $items = @()
        foreach ($x in $v) { if ($null -ne $x) { $items += (Convert-Value $x) } }
        return (Convert-Sanitize ($items -join ', '))
    }
    if ($v -is [System.Management.Automation.PSCustomObject]) {
        $parts = @()
        foreach ($p in $v.PSObject.Properties) { $parts += ('{0}: {1}' -f $p.Name, (Convert-Value $p.Value)) }
        return (Convert-Sanitize ($parts -join '; '))
    }
    return (Convert-Sanitize ([string]$v))
}

function Get-HtmlEscape {
    param($s)
    $t = [string]$s
    $t = $t.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;').Replace("'", '&#x27;')
    return $t
}

function Get-Flattened {
    param($records)
    # Spalten in Reihenfolge des ersten Auftretens sammeln
    $cols = New-Object System.Collections.Generic.List[string]
    foreach ($r in $records) {
        if ($r -is [System.Management.Automation.PSCustomObject]) {
            foreach ($p in $r.PSObject.Properties) {
                if (-not $cols.Contains($p.Name)) { [void]$cols.Add($p.Name) }
            }
        }
    }
    # Positions-Artefaktspalten entfernen (Endung ".Name", ausser genau "Name")
    $cols = @($cols | Where-Object { -not ($_.EndsWith('.Name') -and $_ -ne 'Name') })

    $rows = @()
    foreach ($r in $records) {
        if ($r -is [System.Management.Automation.PSCustomObject]) {
            $row = @()
            foreach ($c in $cols) {
                $prop = $r.PSObject.Properties[$c]
                if ($prop) { $row += (Convert-Value $prop.Value) } else { $row += '' }
            }
            $rows += ,$row
        } else {
            $rows += ,@((Convert-Value $r))
        }
    }
    # Komplett leere Spalten entfernen
    $keep = @()
    for ($i = 0; $i -lt $cols.Count; $i++) {
        $has = $false
        foreach ($row in $rows) { if ($row[$i].Trim().Length -gt 0) { $has = $true; break } }
        if ($has) { $keep += $i }
    }
    $ncols = @()
    foreach ($i in $keep) { $ncols += $cols[$i] }
    $nrows = @()
    foreach ($row in $rows) {
        $nr = @()
        foreach ($i in $keep) { $nr += $row[$i] }
        $nrows += ,$nr
    }
    return [PSCustomObject]@{ Cols = $ncols; Rows = $nrows }
}

function Read-Json {
    param([string]$path)
    $raw = Get-Content -LiteralPath $path -Raw -Encoding UTF8
    if ($raw.Length -gt 0 -and $raw[0] -eq [char]0xFEFF) { $raw = $raw.Substring(1) }  # evtl. BOM entfernen
    return ($raw | ConvertFrom-Json)
}

# ---------- Datenquellen ----------
$overviewPath = Join-Path $SourceDir '_Uebersicht.json'
if (-not (Test-Path -LiteralPath $overviewPath)) { throw "_Uebersicht.json nicht gefunden in $SourceDir" }
$data = Read-Json $overviewPath

$jsonFiles = Get-ChildItem -Path $SourceDir -Filter '*.json' -File

function Get-Records {
    param([string]$modul)
    $suffix = ('_{0}.json' -f $modul).ToLower()
    $file = $jsonFiles | Where-Object { $_.Name.ToLower().EndsWith($suffix) } | Select-Object -First 1
    if (-not $file) { return @() }
    try { $d = Read-Json $file.FullName } catch { return @() }
    if ($null -eq $d) { return @() }
    if ($d -is [System.Array]) { return $d }
    return @($d)
}

# ---------- Geraet / Datum: Parameter oder aus Ordnernamen ableiten ----------
# Hinweis: PowerShell-Variablen sind case-insensitiv, daher direkt mit den
# Parametern $Device/$ExportDate arbeiten (keine separaten $device-Variablen).
$folderName = Split-Path $SourceDir -Leaf
if ($folderName -match 'SophosExport_(.+)_(\d{4}-\d{2}-\d{2})_(\d{2})-(\d{2})-(\d{2})') {
    if (-not $Device)     { $Device = $Matches[1] }
    if (-not $ExportDate) { $ExportDate = ('{0} {1}:{2}' -f $Matches[2], $Matches[3], $Matches[4]) }
}
if (-not $Device)     { $Device = 'Unbekannt' }
if (-not $ExportDate) { $ExportDate = (Get-Date).ToString('yyyy-MM-dd HH:mm') }
$today = (Get-Date).ToString('yyyy-MM-dd')

# ---------- Abschnitte aufbauen ----------
$sections = @()
foreach ($e in $data) {
    $bereich = Convert-Value $e.Bereich
    $modul   = Convert-Value $e.Modul
    $status  = Convert-Value $e.Status
    $records = @(Get-Records $modul)
    $flat    = Get-Flattened $records
    $sections += [PSCustomObject]@{
        Bereich = $bereich
        Modul   = $modul
        Count   = $records.Count
        Status  = $status
        Cols    = @($flat.Cols)
        Rows    = @($flat.Rows)
    }
}

# ---------- HTML erzeugen ----------
$sb = New-Object System.Text.StringBuilder
[void]$sb.Append(@"
<!DOCTYPE html><html lang="de"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Sophos Firewall – Konfiguration $(Get-HtmlEscape $device)</title>
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
<h1>Sophos Firewall – Konfigurationsübersicht</h1>
<div class="meta">Gerät $(Get-HtmlEscape $device) &nbsp;·&nbsp; Export $(Get-HtmlEscape $exportDate) &nbsp;·&nbsp; erstellt $today</div>
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
[void]$sb.Append('<table class="summary"><thead><tr><th>Bereich</th><th>Modul</th><th>Datensätze</th><th>Status</th></tr></thead><tbody>')
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
        [void]$sb.Append('<p class="empty-note">Keine Einträge vorhanden.</p>')
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

[void]$sb.Append('</main></div><footer>Erstellt aus dem Sophos-Export · nur zur internen Dokumentation</footer></body></html>')

# ---------- Schreiben (UTF-8 ohne BOM) ----------
$enc = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($OutFile, $sb.ToString(), $enc)

$nonEmpty = ($sections | Where-Object { $_.Count -gt 0 }).Count
Write-Host ("HTML erstellt: {0}" -f $OutFile) -ForegroundColor Green
Write-Host ("Abschnitte gesamt: {0}  |  mit Daten: {1}" -f $sections.Count, $nonEmpty)
