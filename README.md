# Sophos Firewall – Konfigurations-Export (PowerShell + XML-API)

`Export-SophosFirewallConfig.ps1` verbindet sich per HTTPS mit der **XML-API einer
lokalen Sophos Firewall (SFOS)** und exportiert Firewall-Regeln und Einstellungen
wahlweise in eine **Excel-Datei** (ein Tabellenblatt pro Bereich, mit allen Feldern
als Spalten) oder – falls Excel-Export nicht möglich ist – automatisch in
**JSON- und XML-Dateien pro Bereich**.

Zusätzlich wird am Ende jedes Exports automatisch ein **HTML-Report**
(`Sophos_Firewall_Konfiguration.html`) mit Navigation, Zusammenfassung und
einer Tabelle pro Modul erstellt – so erhält man standardmäßig
**Excel und HTML** in einem Durchlauf. Alles ist in **einem einzigen Skript**
enthalten.

## Funktionen

- Interaktive Abfrage von **Firewall-IP-Adresse** und **API-Zugangsdaten**
  (Passwort wird verdeckt eingegeben)
- Erreichbarkeits- und **Login-Prüfung** vor dem Export (bis zu 3 Versuche bei
  falschen Zugangsdaten)
- **Auswahlmenü**: nur bestimmte Konfigurationsbereiche exportieren oder alles
  (Mehrfachauswahl mit Komma und Bereichen, z. B. `1,3,5-8`, `A` = alles)
- **Excel-Export** mit dem PowerShell-Modul [ImportExcel](https://www.powershellgallery.com/packages/ImportExcel)
  (wird bei Bedarf automatisch installiert):
  - Tabellenblatt **„Uebersicht"** mit Bereich, Modul, Datensatzanzahl und Status
  - **ein Tabellenblatt pro API-Modul** (z. B. `FirewallRule`, `IPHost`, …)
  - verschachtelte XML-Felder werden zu Spalten abgeflacht
    (`NetworkPolicy.SourceZones.Zone`), Listen zu kommagetrennten Werten
    (`HTTP, HTTPS, DNS`)
  - Autofilter, fixierte und fett formatierte Kopfzeile
- **Automatischer Fallback**: steht ImportExcel nicht zur Verfügung, wird pro
  Bereich eine JSON-Datei (aufbereitete Tabellendaten) **und** eine XML-Datei
  (API-Rohantwort) erstellt, plus `_Uebersicht.json`
- **HTML-Report**: nach jedem Export entsteht zusätzlich eine übersichtliche
  HTML-Seite (`Sophos_Firewall_Konfiguration.html`) mit Inhaltsverzeichnis,
  Zusammenfassung und einer Tabelle pro Modul (abschaltbar mit `-NoHtml`)
- Vollständig **skriptbar** über Parameter (für Automatisierung/Aufgabenplanung)
- Funktioniert mit **Windows PowerShell 5.1** und **PowerShell 7+**
  (Windows/Linux/macOS); selbstsignierte Firewall-Zertifikate werden akzeptiert

## Voraussetzungen

1. **PowerShell 5.1 oder neuer** (empfohlen: PowerShell 7)
2. **API-Zugriff auf der Sophos Firewall aktivieren:**
   - Anmelden an der Web-Admin-Konsole (Standard: `https://<Firewall-IP>:4444`)
   - **Backup & Firmware → API**
   - **API-Konfiguration aktivieren**
   - Die **IP-Adresse des Rechners, der das Skript ausführt**, unter
     *„Zulässige IP-Adresse"* eintragen
3. Ein **Administrator-Konto** der Firewall (Benutzername + Passwort) – die
   XML-API nutzt die normalen Admin-Zugangsdaten
4. Für den Excel-Export: Internetzugang zur PowerShell Gallery **oder** das
   Modul ImportExcel bereits installiert
   (`Install-Module ImportExcel -Scope CurrentUser`)

> **Hinweis:** Excel selbst muss **nicht** installiert sein – die `.xlsx`-Datei
> wird direkt vom ImportExcel-Modul erzeugt.

## Schnellstart (interaktiv)

```powershell
# ggf. Ausführung von Skripten erlauben (einmalig, nur aktuelle Sitzung):
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

.\Export-SophosFirewallConfig.ps1
```

Ablauf:

1. **Firewall IP-Adresse** eingeben (z. B. `192.168.1.1`)
2. **API-Benutzername** und **API-Passwort** eingeben (Passwort verdeckt)
3. Im **Auswahlmenü** die gewünschten Bereiche wählen:

```
=============================================================
  Welche Konfigurationsbereiche sollen exportiert werden?
=============================================================
  [ 1] Firewall-Regeln                (FirewallRule, FirewallRuleGroup)
  [ 2] NAT-Regeln                     (NATRule)
  [ 3] Netzwerk-Objekte (Hosts)       (IPHost, IPHostGroup, FQDNHost, ...)
  [ 4] Dienste (Services)             (Services, ServiceGroup)
  ...
  [ A] ALLES exportieren
  [ Q] Abbrechen

Mehrfachauswahl mit Komma und Bereichen moeglich, z.B.: 1,3,5-8
Auswahl: A
```

4. Der Export landet im Unterordner `SophosExport_<IP>_<Zeitstempel>`.

## Nutzung mit Parametern (Automatisierung)

```powershell
# Alles nach Excel exportieren:
.\Export-SophosFirewallConfig.ps1 -FirewallIP 192.168.1.1 `
    -Username admin -Password 'Geheim123!' -Sections All

# Nur Firewall-Regeln und Hosts, als JSON:
.\Export-SophosFirewallConfig.ps1 -FirewallIP 192.168.1.1 `
    -Username admin -Password 'Geheim123!' `
    -Sections FirewallRules,Hosts -OutputFormat Json

# Rohe XML-API-Antworten sichern (z. B. als Backup/Doku):
.\Export-SophosFirewallConfig.ps1 -FirewallIP 192.168.1.1 `
    -Username admin -Password 'Geheim123!' -Sections All -OutputFormat Xml

# Verfügbare Bereichs-Schlüssel anzeigen:
.\Export-SophosFirewallConfig.ps1 -ListSections
```

### Parameterübersicht

| Parameter          | Beschreibung                                                                 | Standard |
|--------------------|------------------------------------------------------------------------------|----------|
| `-FirewallIP`      | IP-Adresse oder Hostname der Firewall (sonst interaktive Abfrage)            | –        |
| `-Port`            | HTTPS-Port der Admin-Konsole/API                                             | `4444`   |
| `-Username`        | API-Benutzername (sonst interaktive Abfrage)                                 | –        |
| `-Password`        | API-Passwort im Klartext – **nur für Automatisierung** (sonst verdeckte Abfrage) | –    |
| `-Sections`        | Bereiche (`All` oder Schlüssel, siehe `-ListSections`)                       | Menü     |
| `-OutputFormat`    | `Auto` (Excel, sonst Fallback), `Excel`, `Json`, `Xml`                       | `Auto`   |
| `-OutputDirectory` | Zielverzeichnis                                                              | `SophosExport_<IP>_<Zeit>` |
| `-NoHtml`          | Unterdrückt die zusätzliche Erstellung des HTML-Reports                      | –        |
| `-ListSections`    | Zeigt alle Bereichs-Schlüssel an und beendet das Skript                      | –        |
| `-TimeoutSec`      | Timeout pro API-Anfrage (Sekunden)                                           | `60`     |

### Exportierbare Bereiche

| Nr | Schlüssel (`-Sections`) | Bereich                    | API-Module (Sophos XML-API)                                  |
|----|--------------------------|----------------------------|--------------------------------------------------------------|
| 1  | `FirewallRules`          | Firewall-Regeln            | `FirewallRule`, `FirewallRuleGroup`                          |
| 2  | `NatRules`               | NAT-Regeln                 | `NATRule`                                                    |
| 3  | `Hosts`                  | Netzwerk-Objekte (Hosts)   | `IPHost`, `IPHostGroup`, `FQDNHost`, `FQDNHostGroup`, `MACHost` |
| 4  | `Services`               | Dienste                    | `Services`, `ServiceGroup`                                   |
| 5  | `Zones`                  | Zonen                      | `Zone`                                                       |
| 6  | `Interfaces`             | Interfaces und Netzwerk    | `Interface`, `VLAN`, `LAG`, `BridgePair`                     |
| 7  | `Routing`                | Statische Routen           | `UnicastRoute`                                               |
| 8  | `DnsDhcp`                | DNS und DHCP               | `DNS`, `DHCPServer`                                          |
| 9  | `Users`                  | Benutzer und Gruppen       | `User`, `Group`                                              |
| 10 | `WebFilter`              | Webfilter                  | `WebFilterPolicy`, `WebFilterURLGroup`                       |
| 11 | `AppFilter`              | Applikationsfilter         | `ApplicationFilterPolicy`                                    |
| 12 | `IPS`                    | IPS-Richtlinien            | `IPSPolicy`                                                  |
| 13 | `VPN`                    | VPN (IPsec / SSL)          | `VPNIPSecConnection`, `SSLVPNPolicy`                         |
| 14 | `Wireless`               | WLAN                       | `WirelessNetworks`, `AccessPoint`                            |
| 15 | `Certificates`           | Zertifikate                | `Certificate`, `CertificateAuthority`                        |
| 16 | `Schedules`              | Zeitpläne                  | `Schedule`                                                   |
| 17 | `System`                 | System und Verwaltung      | `AdminSettings`, `Notification`, `SyslogServers`, `SNMPAgentConfiguration`, `BackupRestore` |

Module ohne konfigurierte Objekte melden in der Übersicht
`Status 526 - Record does not exist` – das ist normal und kein Fehler.

## Ausgabe

**Excel-Modus** (`-OutputFormat Auto`/`Excel`):

```
SophosExport_192.168.1.1_2026-07-21_18-00-00/
├── Sophos_Firewall_Export_2026-07-21_18-00-00.xlsx
│   ├── Blatt "Uebersicht"   (alle Module mit Anzahl + Status)
│   ├── Blatt "FirewallRule" (eine Zeile pro Regel, alle Felder als Spalten)
│   ├── Blatt "IPHost"
│   └── ... (ein Blatt pro Modul mit Daten)
└── Sophos_Firewall_Konfiguration.html   (HTML-Report)
```

**Datei-Modus** (`-OutputFormat Json`/`Xml` oder automatischer Fallback):

```
SophosExport_192.168.1.1_2026-07-21_18-00-00/
├── FirewallRules_FirewallRule.json   (aufbereitete Tabellendaten)
├── FirewallRules_FirewallRule.xml    (API-Rohantwort)
├── Hosts_IPHost.json
├── ...
├── _Uebersicht.json
└── Sophos_Firewall_Konfiguration.html   (HTML-Report)
```

## HTML-Report

Der HTML-Report ist direkt in `Export-SophosFirewallConfig.ps1` integriert
und wird bei jedem Export automatisch mit erzeugt (abschaltbar mit `-NoHtml`).
Er enthält:

- Kopfzeile mit Gerät (Firewall-IP), Export- und Erstellungsdatum
- seitliche Navigation mit Datensatzanzahl pro Modul
- Zusammenfassungstabelle aller Bereiche
- einen Abschnitt mit Tabelle pro Modul; leere Artefakt-Spalten werden
  entfernt, sehr lange Werte gekürzt

Die Datei ist eigenständig (kein Internet/JavaScript nötig) und kann direkt
im Browser geöffnet oder weitergegeben werden.

## Technischer Hintergrund

Das Skript sendet pro Modul einen `Get`-Request an den XML-API-Endpunkt der
Firewall:

```
POST https://<Firewall-IP>:4444/webconsole/APIController
     reqxml=<Request>
              <Login><Username>...</Username><Password>...</Password></Login>
              <Get><FirewallRule/></Get>
            </Request>
```

Die XML-Antwort wird geparst, jeder Datensatz rekursiv zu einer Tabellenzeile
abgeflacht (verschachtelte Ebenen mit `.` verbunden, Listen mit `, `
zusammengefasst) und anschließend exportiert.

## Fehlerbehebung

| Meldung / Problem                                            | Ursache und Lösung |
|--------------------------------------------------------------|--------------------|
| `Anmeldung fehlgeschlagen: Benutzername oder Passwort falsch` | Zugangsdaten prüfen (Admin-Konto der Firewall). |
| `API-Zugriff verweigert: API operation not permitted from …`  | Die IP des Skript-Rechners ist auf der Firewall nicht als API-IP freigegeben → **Backup & Firmware → API** eintragen. |
| `… ist nicht erreichbar (TCP-Timeout)`                        | IP/Port prüfen (Standard 4444), Firewall-Regeln/Zonen-Zugriff (Admin-Dienste HTTPS) für das Quellnetz prüfen. |
| `Excel-Export nicht moeglich (Modul "ImportExcel" fehlt)`     | Ohne Internetzugang: `Install-Module ImportExcel` auf einem Rechner mit Internet, Modul kopieren – oder `-OutputFormat Json`/`Xml` verwenden. |
| Blatt fehlt in der Excel-Datei                                | Module ohne Datensätze werden nur in der Übersicht aufgeführt (Status 526). |
| Skript startet nicht (`... kann nicht geladen werden`)        | `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass` |

## Sicherheitshinweise

- Der Parameter `-Password` (Klartext) ist nur für Automatisierung gedacht –
  interaktiv wird das Passwort **verdeckt** abgefragt und nicht gespeichert.
- Die Zertifikatsprüfung wird für die Firewall-Verbindung deaktiviert, da
  Sophos Firewalls standardmäßig selbstsignierte Zertifikate verwenden.
- Die exportierten Dateien enthalten die vollständige Konfiguration
  (u. a. Regelwerk, Objekte) – entsprechend vertraulich behandeln.

## Getestet

- PowerShell 7.6 (Linux) gegen eine simulierte Sophos-XML-API
  (Endpunkt `/webconsole/APIController`, SFOS-Antwortformat), inkl.:
  - interaktiver Ablauf (IP → Zugangsdaten → Menü, Auswahl `1,3-5` und `A`)
  - falsche Zugangsdaten (3 Versuche, sauberer Abbruch)
  - ungültige Menüeingaben, nicht erreichbare Firewall, Statuscode 526
  - Excel-, JSON- und XML-Ausgabe inkl. automatischem Fallback
