<#
.SYNOPSIS
    LIGHT VERSION (6 of 24 checks, free). File server report: checks shares, file system permissions, SMB settings, and data inventory of one or more
    Windows file servers and generates an HTML report with anomalies.

    Deutsch:
    LIGHT-VERSION (6 von 24 Prüfpunkten, kostenlos). Dateiserver-Bericht: prüft Freigaben, Dateisystemrechte, SMB-Einstellungen und Datenbestand eines oder mehrerer
    Windows-Dateiserver und erzeugt einen HTML-Bericht mit Auffälligkeiten.

.DESCRIPTION
    The script operates exclusively in READ-ONLY mode. It changes neither shares nor permissions, files, or settings and
    sends no data to third parties or the internet. It writes exactly one HTML file (and CSV files if requested) to the specified output folder.

    Prerequisites: Windows PowerShell 5.1 and an account with read permissions on shares and file system. The evaluation is most accurate when the script runs directly on the file server with administrator rights. With -Server, additional servers can be checked over the network (remote administration via CIM/WinRM, files via administrative shares such as C$).
    Anything the account cannot read appears as "not checked" and never as "OK".

    Checks performed: shares (Everyone access, system folders, hidden shares, orphaned shares), file system permissions of
    the top folder levels of each share (write permissions for Everyone, orphaned accounts, permissions directly assigned to users, disabled
    inheritance), SMB settings (SMB1, signing, anonymous access), volumes (free space, shadow copies, quotas),
    the operating system, and data inventory (old and large files, executable files, suspicious file names).

    Deutsch:
    Das Skript arbeitet ausschließlich LESEND. Es ändert weder Freigaben noch Berechtigungen, Dateien oder Einstellungen und
    sendet keine Daten an Dritte oder ins Internet. Es schreibt genau eine HTML-Datei (und auf Wunsch CSV-Dateien) in den angegebenen Ausgabeordner.

    Voraussetzungen: Windows PowerShell 5.1 und ein Konto mit Leserechten auf Freigaben und Dateisystem. Am genauesten wird die
    Auswertung, wenn das Skript direkt auf dem Dateiserver mit Administratorrechten läuft. Mit -Server können weitere Server
    über das Netzwerk geprüft werden (Remote-Verwaltung per CIM/WinRM, Dateien über die administrativen Freigaben wie C$).
    Was das Konto nicht lesen darf, erscheint als "nicht geprüft" und nie als "in Ordnung".

    Geprüft werden: Freigaben (Zugriff für Jeder, Systemordner, versteckte Freigaben, verwaiste Freigaben), Dateisystemrechte der
    obersten Ordnerebenen jeder Freigabe (Schreibrechte für Jeder, verwaiste Konten, Rechte direkt an Benutzern, abgeschaltete
    Vererbung), SMB-Einstellungen (SMB1, Signierung, anonymer Zugriff), Volumes (freier Platz, Schattenkopien, Kontingente),
    das Betriebssystem und den Datenbestand (alte und große Dateien, ausführbare Dateien, auffällige Dateinamen).

.PARAMETER Server
    One or more file servers (name). Default: the local computer.

    Deutsch:
    Ein oder mehrere Dateiserver (Name). Standard: der lokale Computer.

.PARAMETER Freigaben
    Evaluate only these shares (names, wildcards allowed). Default: all shares except default shares (C$, ADMIN$, IPC$).

    Deutsch:
    Nur diese Freigaben auswerten (Namen, Platzhalter erlaubt). Standard: alle Freigaben außer den Standardfreigaben (C$, ADMIN$, IPC$).

.PARAMETER Tiefe
    How many folder levels below each share to check file system permissions (default 3).

    Deutsch:
    Wie viele Ordnerebenen unterhalb jeder Freigabe die Dateisystemrechte geprüft werden (Standard 3).

.PARAMETER MaxOrdner
    Check at most this many folders per share for permissions (default 3000). When reached, a note appears in the report.

    Deutsch:
    Höchstens so viele Ordner je Freigabe auf Rechte prüfen (Standard 3000). Bei Erreichen steht ein Hinweis im Bericht.

.PARAMETER MaxDateien
    Read at most this many files per share for data analysis (default 50000). When reached, a note appears in the report.

    Deutsch:
    Höchstens so viele Dateien je Freigabe für die Datenauswertung lesen (Standard 50000). Bei Erreichen steht ein Hinweis im Bericht.

.PARAMETER OhneDateien
    Do not count files (only shares, permissions, and settings). Much faster on large servers.

    Deutsch:
    Keine Dateien zählen (nur Freigaben, Rechte und Einstellungen). Deutlich schneller auf großen Servern.

.PARAMETER AlterJahre
    Files that have not been modified for this number of years are considered old (default 5).

    Deutsch:
    Dateien, die so viele Jahre nicht geändert wurden, gelten als alt (Standard 5).

.PARAMETER GrosseDateiGB
    Files at or above this size in GB are listed (default 1).

    Deutsch:
    Dateien ab dieser Größe in GB werden aufgelistet (Standard 1).

.PARAMETER FreiProzent
    Warns when a volume with shares has less than this percentage free (default 15).

    Deutsch:
    Warnt, wenn auf einem Volume mit Freigaben weniger Prozent frei sind (Standard 15).

.PARAMETER Ausgabe
    Path to the HTML file. Default: Dateiserver-Bericht_<Server>_<Datum>.html in the current folder.

    Deutsch:
    Pfad der HTML-Datei. Standard: Dateiserver-Bericht_<Server>_<Datum>.html im aktuellen Ordner.

.PARAMETER Anmeldung
    Credentials (Get-Credential) for querying remote servers.

    Deutsch:
    Anmeldedaten (Get-Credential) für die Abfrage entfernter Server.

.PARAMETER Csv
    Writes a CSV file alongside the report for each check point with hits.

    Deutsch:
    Schreibt je Prüfpunkt mit Treffern eine CSV-Datei neben den Bericht.

.PARAMETER NichtOeffnen
    Does not automatically open the completed report in the browser.

    Deutsch:
    Öffnet den fertigen Bericht nicht automatisch im Browser.

.PARAMETER PassThru
    Additionally returns the findings as objects (for custom analysis).

    Deutsch:
    Gibt die Befunde zusätzlich als Objekte zurück (für eigene Auswertungen).

.PARAMETER Language
    Language of the report and the console output: 'de' (German) or 'en' (English). Default: German if the Windows display language
    is German, otherwise English.

    Deutsch:
    Sprache von Bericht und Konsolenausgabe: 'de' (Deutsch) oder 'en' (Englisch). Standard: Deutsch, wenn die Anzeigesprache von
    Windows Deutsch ist, sonst Englisch.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Dateiserver-Bericht.ps1

.EXAMPLE
    .\Dateiserver-Bericht.ps1 -Server FS01, FS02 -Freigaben Daten*, Projekte -Tiefe 4 -Csv

.NOTES
    IT Admin Tools by MM - Version 1.0 (11.2026)
    Kein Ersatz für eine Sicherheitsprüfung oder Antivirenlösung. Nutzung auf eigene Verantwortung.
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [ValidateSet('de', 'en')][string]$Language = $(if ((Get-UICulture).TwoLetterISOLanguageName -eq 'de') { 'de' } else { 'en' }),
    [string[]]$Server,
    [string[]]$Freigaben,
    [ValidateRange(1, 10)][int]$Tiefe = 3,
    [ValidateRange(10, 100000)][int]$MaxOrdner = 3000,
    [ValidateRange(100, 2000000)][int]$MaxDateien = 50000,
    [switch]$OhneDateien,
    [ValidateRange(1, 50)][int]$AlterJahre = 5,
    [ValidateRange(1, 1000)][int]$GrosseDateiGB = 1,
    [ValidateRange(1, 90)][int]$FreiProzent = 15,
    [Alias('Output')][string]$Ausgabe,
    [Alias('Credential')][System.Management.Automation.PSCredential]$Anmeldung,
    [switch]$Csv,
    [Alias('NoOpen')][switch]$NichtOeffnen,
    [switch]$PassThru
)
# Sprache des Berichts: de oder en (Parameter -Language). T 'deutsch' 'english' liefert den Text der gewählten Sprache.
$En = ($Language -eq 'en')
function T([string]$de, [string]$en) { if ($script:En) { $en } else { $de } }

$ErrorActionPreference = 'Stop'
$Version = '1.1'

#@SPRACHE@
# ------------------------------------------------------------------ Hilfsfunktionen
function Datum($d, $leer = $null) { if ($null -eq $leer) { $leer = T 'nie' 'never' }; if ($d) { ([datetime]$d).ToString((T 'dd.MM.yyyy' 'yyyy-MM-dd')) } else { $leer } }
function JaNein($b) { if ($b) { T 'Ja' 'Yes' } else { T 'Nein' 'No' } }
function Ort($dn) { ($dn -replace '^(CN|OU)=.+?(?<!\\),', '') -replace ',DC=.*$', '' }
function Html($t) { [System.Net.WebUtility]::HtmlEncode([string]$t) }
function Kurz($t, $n = 220) { $t = ([string]$t -replace '\s+', ' ').Trim(); if ($t.Length -gt $n) { $t.Substring(0, $n) + ' ...' } else { $t } }

# CSV-Dateien öffnen sich in Excel: Werte, die mit = + - @ beginnen, würden dort als Formel ausgeführt (Namen von Dateien, Ordnern, Konten oder Computern
# stammen von Benutzern). Diese Funktion ersetzt für dieses Skript den Befehl Export-Csv und setzt vor solche Werte ein Hochkomma.
function Export-Csv {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline = $true)]$InputObject, [Parameter(Mandatory)][string]$Path, [switch]$NoTypeInformation, [string]$Delimiter = ';', [string]$Encoding = 'UTF8')
    begin { $alle = New-Object System.Collections.Generic.List[object] }
    process { if ($null -ne $InputObject) { $alle.Add($InputObject) } }
    end {
        $sicher = foreach ($o in $alle) {
            $n = [ordered]@{}
            foreach ($p in $o.PSObject.Properties) { $v = $p.Value; if (($v -is [string]) -and ($v -match '^[=+@\t\r]|^-.')) { $v = "'" + $v }; $n[$p.Name] = $v }
            [pscustomobject]$n
        }
        if ($alle.Count) { $sicher | Microsoft.PowerShell.Utility\Export-Csv -Path $Path -NoTypeInformation:$NoTypeInformation -Delimiter $Delimiter -Encoding $Encoding }   # leere Liste: keine Datei (sie hätte keine Kopfzeile)
    }
}

# Tabellen: ConvertTo-Html und Export-Csv nehmen die Spalten des ersten Objekts. Haben die Zeilen einer Liste verschiedene
# Eigenschaften (z. B. Fehlzeile neben normalen Treffern), gingen Spalten verloren: alle Namen der Reihe nach vereinigen.
function Einheitlich($liste) {
    $liste = @($liste | Where-Object { $null -ne $_ })
    if ($liste.Count -ge 2) {
        $namen = New-Object System.Collections.Generic.List[string]
        foreach ($o in $liste) { foreach ($p in $o.PSObject.Properties.Name) { if (-not $namen.Contains($p)) { $namen.Add($p) } } }
        if ($namen.Count -ne @($liste[0].PSObject.Properties.Name).Count) { $liste = @($liste | Select-Object -Property $namen) }
    }
    SpaltenAnzeige $liste
}
# Englische Berichte: Spaltenüberschriften aus der Übersetzungstabelle (intern bleiben die deutschen Namen, der Code liest sie weiter)
$SpaltenEn = @{
    'Abfrage' = 'Query'
    'Administrator' = 'Administrator'
    'Adresse' = 'Address'
    'Aktiv' = 'Active'
    'Aktiv verknüpft mit' = 'Active linked to'
    'Aktive Verknüpfungen' = 'Active links'
    'Alter in Tagen' = 'Age in days'
    'Angelegt' = 'Created'
    'Anmeldung' = 'Last logon'
    'Anmeldungen' = 'Logons'
    'Anteil (%)' = 'Share (%)'
    'Anzahl' = 'Count'
    'Art' = 'Account type'
    'Auf Domänencontroller' = 'On domain controller'
    'Aufräumen' = 'Cleanup'
    'Auslastung' = 'Utilization'
    'Auslösender Rechner' = 'Triggering computer'
    'Auslösender Rechner (Anzahl)' = 'Triggering computers (count)'
    'Aussteller' = 'Issuer'
    'Beispiele' = 'Examples'
    'Belegt' = 'Used'
    'Belegt (%)' = 'Used (%)'
    'Benutzer und Computer (auch verschachtelt)' = 'Users and computers (also nested)'
    'Benutzereinstellungen' = 'User settings'
    'Benutzername' = 'Username'
    'Berechtigter' = 'Principal'
    'Bereich' = 'Scope'
    'Bereiche' = 'Scopes'
    'Beschreibung' = 'Description'
    'Besitzer' = 'Owner'
    'Betriebssystem' = 'OS'
    'Betroffene Konten' = 'Affected accounts'
    'Bewertung' = 'Rating'
    'Bis' = 'To'
    'BitLocker' = 'BitLocker'
    'Computer' = 'Computer'
    'Computereinstellungen' = 'Computer settings'
    'Darf anwenden' = 'Can apply'
    'Darf delegieren an' = 'Can delegate to'
    'Darf sich hier als Benutzer ausgeben' = 'Can log on as user'
    'Datei' = 'File'
    'Dateien' = 'Files'
    'Dateisystem' = 'File system'
    'Dateityp' = 'File type'
    'Datenbank' = 'Database'
    'Detail' = 'Detail'
    'Details' = 'Details'
    'Direkte Mitglieder' = 'Direct members'
    'Domäne' = 'Domain'
    'Domänencontroller' = 'Domain controller'
    'Domänennamenmaster' = 'Domain naming master'
    'Dynamische Updates' = 'Dynamic updates'
    'E-Mail' = 'Email'
    'Eigenschaft' = 'Property'
    'Einstellung' = 'Setting'
    'Empfohlen' = 'Recommended'
    'Enthaltene Gruppen' = 'Included groups'
    'Ereignisse' = 'Events'
    'Erreichbar' = 'Reachable'
    'Erste' = 'First'
    'Erster' = 'First'
    'Erzwungen' = 'Forced'
    'Failover' = 'Failover'
    'Falsches Kennwort' = 'Wrong password'
    'Fehlanmeldungen am DC' = 'Failed logons at DC'
    'Fehlversuche' = 'Failed attempts'
    'Feststellung' = 'Finding'
    'Fingerabdruck' = 'Fingerprint'
    'Firewall' = 'Firewall'
    'Frei' = 'Free'
    'Freigabe' = 'Share'
    'Freigaberecht' = 'Share permission'
    'Funktion' = 'Function'
    'Gelesen von' = 'Read by'
    'Geprüft' = 'Checked'
    'Gesamtgröße' = 'Total size'
    'Geschützte Benutzer' = 'Protected users'
    'Gesperrt seit' = 'Locked since'
    'Geändert' = 'Changed'
    'Gilt' = 'Applies to'
    'Gilt für' = 'Applies to'
    'Globaler Katalog' = 'Global catalog'
    'Grenze' = 'Limit'
    'Groesse' = 'Size'
    'Grund' = 'Reason'
    'Gruppe' = 'Group'
    'Gruppen' = 'Groups'
    'Gruppen (auch verschachtelt)' = 'Groups (including nested)'
    'Gruppenrichtlinie' = 'Group policy'
    'Größe' = 'Size'
    'Größe (GB)' = 'Size (GB)'
    'Gültig ab' = 'Valid from'
    'Gültig bis' = 'Valid to'
    'Herkunft' = 'Origin'
    'Hersteller' = 'Vendor'
    'Hinweis' = 'Note'
    'Häufigster Grund' = 'Most common reason'
    'IP-Adresse' = 'IP address'
    'Im AD gespeichert' = 'Stored in AD'
    'In der Domäne vorhanden' = 'Present in domain'
    'Infrastrukturmaster' = 'Infrastructure master'
    'Inhalt' = 'Content'
    'Ist' = 'Is'
    'Kennung' = 'ID'
    'Kennwort gesetzt' = 'Password set'
    'Kennwort läuft nie ab' = 'Password never expires'
    'Kerberos' = 'Kerberos'
    'Konten' = 'Accounts'
    'Konten (auch verschachtelt)' = 'Accounts (also nested)'
    'Konto' = 'Account'
    'Konto für DNS-Einträge' = 'Account for DNS entries'
    'Kopie bleibt im Postfach' = 'Copy remains in mailbox'
    'LAPS' = 'LAPS'
    'LAPS-Ablauf' = 'LAPS expiry'
    'Laufwerk' = 'Drive'
    'Lease-Dauer' = 'Lease duration'
    'Leitet weiter an' = 'Forwards to'
    'Leseweg' = 'Read path'
    'Letzte Aktualisierung' = 'Last update'
    'Letzte Anmeldung' = 'Last logon'
    'Letzte Vollsicherung' = 'Last full backup'
    'Letzte inkrementelle Sicherung' = 'Last incremental backup'
    'Letzter' = 'Last'
    'Letzter Aufräumlauf' = 'Last cleanup run'
    'Letzter Start' = 'Last start'
    'Lokale Admins' = 'Local admins'
    'Mindestlänge' = 'Minimum length'
    'Mitglied' = 'Member'
    'Mitglieder' = 'Members'
    'NLA' = 'NLA'
    'NTLM' = 'NTLM'
    'Name' = 'Name'
    'Nr.' = 'No.'
    'Objekt' = 'Object'
    'Ordner' = 'Folder'
    'Organisationseinheit' = 'OU'
    'Ort' = 'Location'
    'PDC-Emulator' = 'PDC emulator'
    'Pfad' = 'Path'
    'Postfach' = 'Mailbox'
    'Postfach-Typ' = 'Mailbox type'
    'Postfachtyp' = 'Mailbox type'
    'Postfächer' = 'Mailboxes'
    'Privilegierte Gruppen' = 'Privileged groups'
    'Programm' = 'Program'
    'Prüfpunkt' = 'Checkpoint'
    'Quelle' = 'Source'
    'Quelle (Anzahl)' = 'Sources (count)'
    'RID-Master' = 'RID master'
    'Recht' = 'Permission'
    'Rechte' = 'Permissions'
    'Regel' = 'Rule'
    'Reihenfolge' = 'Order'
    'SPN' = 'SPN'
    'Schemamaster' = 'Schema master'
    'Schlüssel' = 'Key'
    'Schreibgeschützt' = 'Read-only'
    'Sendesperre ab' = 'Send blocked from'
    'Server' = 'Server'
    'Sid' = 'SID'
    'Signatur' = 'Signature'
    'Speicher' = 'Storage'
    'Sperrungen' = 'Lockouts'
    'Stand' = 'Status'
    'Standort' = 'Site'
    'Status' = 'Status'
    'Support-Ende' = 'End of support'
    'Support-Ende (spätestes)' = 'End of support (latest)'
    'Tage bis Ablauf' = 'Days until expiry'
    'Tage her' = 'Days ago'
    'Tage ohne Anmeldung' = 'Days without logon'
    'Tage seit Ablauf' = 'Days since expiry'
    'Treffer' = 'Hits'
    'Typ' = 'Type'
    'Vererbung' = 'Inheritance'
    'Verschiedene Konten' = 'Distinct accounts'
    'Verschlüsselt' = 'Encrypted'
    'Version' = 'Version'
    'Version C/B' = 'Version C/B'
    'Verweigert für' = 'Denied for'
    'Verwendung' = 'Use'
    'Virenschutz' = 'Antivirus'
    'Von' = 'From'
    'WMI-Filter' = 'WMI filter'
    'Weg' = 'Path'
    'Weiterleitungen' = 'Forwarding'
    'Wert' = 'Value'
    'Wirksame Richtlinie' = 'Effective policy'
    'Wirksames Recht' = 'Effective permission'
    'Wirkung' = 'Effect'
    'Zeit' = 'Time'
    'Zeitpunkt' = 'Time'
    'Zeitraum' = 'Period'
    'Zertifikat' = 'Certificate'
    'Zone' = 'Zone'
    'Zonen' = 'Zones'
    'Zonenübertragung' = 'Zone transfer'
    'Zugriff für große Gruppen' = 'Access for large groups'
    'Zugriffsbasierte Aufzählung' = 'Access-based enumeration'
    'Zuletzt' = 'Last'
    'Zustand' = 'State'
    'davon Größe (GB)' = 'Of which size (GB)'
    'davon aktiv' = 'Of which active'
    'davon inaktiv' = 'Of which inactive'
    'davon unbekannte Konten' = 'Of which unknown accounts'
    'ist Mitglied von' = 'Member of'
    'Älter als 5 Jahre' = 'Older than 5 years'
    'Älter als 5 Jahren' = 'Older than 5 years'
    'Ältester Eintrag' = 'Oldest entry'
    'Über' = 'Via'
    'Überwachungsprotokoll' = 'Audit log'
}
# Werte, die der Code intern auf Deutsch vergleicht (z. B. Art = 'Benutzer'), erscheinen im englischen Bericht übersetzt (nur Anzeige und CSV)
$WerteEn = New-Object System.Collections.Specialized.OrderedDictionary([System.StringComparer]::Ordinal)   # Groß-/Kleinschreibung zählt (Verwaist / verwaist)
$WerteEn.Add('Alle erweiterten Rechte (u. a. Kennwort zurücksetzen)', 'All extended rights (including password reset)')
$WerteEn.Add('Kein Serverbetriebssystem (Support nicht bewertet)', 'No server operating system (support not assessed)')
$WerteEn.Add('Rechte je Ordner (nur dort, wo sie gesetzt werden)', 'Permissions per folder (only where set)')
$WerteEn.Add('Mehrfach freigegebene Ordner (nur einmal gezählt)', 'Folders shared multiple times (counted only once)')
$WerteEn.Add('Übersprungene Verknüpfungen (Junctions/Symlinks)', 'Skipped links (Junctions/Symlinks)')
$WerteEn.Add('Der Ordner der Freigabe existiert nicht mehr.', 'The share folder no longer exists.')
$WerteEn.Add('DHCP-Überwachungsprotokoll ist ausgeschaltet', 'DHCP audit log is disabled')
$WerteEn.Add('Vollzugriff und "Senden als" im Überblick', 'Full access and "Send as" overview')
$WerteEn.Add('Ordner der Freigabe existiert nicht mehr', 'Share folder no longer exists')
$WerteEn.Add('Enthält außer Sicherheitseinstellungen:', 'Contains except security settings:')
$WerteEn.Add('Nicht geprüft - Postfach nicht gefunden', 'Not checked - mailbox not found')
$WerteEn.Add('WMI-Filter nicht lesbar - nicht geprüft', 'WMI filter not readable - not checked')
$WerteEn.Add('Abfrage fehlgeschlagen - nicht geprüft', 'Query failed - not checked')
$WerteEn.Add('Nicht im Active Directory gespeichert', 'Not stored in Active Directory')
$WerteEn.Add('keine im Active Directory autorisiert', 'None authorized in Active Directory')
$WerteEn.Add('Zugriff auf den Ordner nicht möglich', 'Access to folder not possible')
$WerteEn.Add('Zertifikate (öffentliche Schlüssel)', 'Certificates (public keys)')
$WerteEn.Add('Domänencontroller der Organisation', 'Organization domain controllers')
$WerteEn.Add('Remotedesktop erlaubt Verbindungen', 'Remote Desktop allows connections')
$WerteEn.Add('autorisierte DHCP-Server ermitteln', 'Determine authorized DHCP servers')
$WerteEn.Add('Windows Server 2008 R2 und älter', 'Windows Server 2008 R2 and older')
$WerteEn.Add('Konto deaktiviert oder gesperrt', 'Account disabled or locked')
$WerteEn.Add('Vollzugriff (Zusatz: ObjectType', 'Full access (additional: ObjectType')
$WerteEn.Add('Kennwort oder Konto abgelaufen', 'Password or account expired')
$WerteEn.Add('Lokale Administratorengruppen', 'Local administrator groups')
$WerteEn.Add('Berechtigungen/Besitz ändern', 'Change permissions/ownership')
$WerteEn.Add('Freigabepfade nicht sichtbar', 'Share paths not visible')
$WerteEn.Add('Aktive privilegierte Konten', 'Active privileged accounts')
$WerteEn.Add('Gruppe hat keine Mitglieder', 'Group has no members')
$WerteEn.Add('Keine Verbindung zum Server', 'No connection to server')
$WerteEn.Add('unerlaubte Zeichen im Namen', 'Illegal characters in name')
$WerteEn.Add('Geprüfte Domänencontroller', 'Checked domain controllers')
$WerteEn.Add('Authentifizierte Benutzer', 'Authenticated Users')
$WerteEn.Add('Schwellen dieses Berichts', 'Thresholds of this report')
$WerteEn.Add('Benutzer (inetOrgPerson)', 'User (inetOrgPerson)')
$WerteEn.Add('Datenbestand je Freigabe', 'Data inventory per share')
$WerteEn.Add('Kontosperrungsrichtlinie', 'Account lockout policy')
$WerteEn.Add('Sicherheitseinstellungen', 'Security settings')
$WerteEn.Add('Alle erweiterten Rechte', 'All extended rights')
$WerteEn.Add('Hinweise zur Auswertung', 'Notes for evaluation')
$WerteEn.Add('Nicht abgefragte Server', 'Servers not queried')
$WerteEn.Add('Unterordner und Dateien', 'Subfolders and files')
$WerteEn.Add('Bestätigtes Schreiben:', 'Validated write:')
$WerteEn.Add('Installierte Programme', 'Installed programs')
$WerteEn.Add('Konto für DNS-Einträge', 'Account for DNS records')
$WerteEn.Add('Kontosperrungsschwelle', 'Account lockout threshold')
$WerteEn.Add('Nicht lesbare Bereiche', 'Unreadable areas')
$WerteEn.Add('Berechtigungen ändern', 'Change permissions')
$WerteEn.Add('Hinweise zum Auslesen', 'Notes for reading')
$WerteEn.Add('Windows 8.1 und älter', 'Windows 8.1 and older')
$WerteEn.Add('Überwachungsprotokoll', 'Audit log')
$WerteEn.Add('Nicht berücksichtigt', 'Not considered')
$WerteEn.Add('Anonymous-Anmeldung', 'Anonymous logon')
$WerteEn.Add('Infrastrukturmaster', 'Infrastructure master')
$WerteEn.Add('Secure Boot ist aus', 'Secure Boot is off')
$WerteEn.Add('Domänennamenmaster', 'Domain naming master')
$WerteEn.Add('Kennwortrichtlinie', 'Password policy')
$WerteEn.Add('Komponentenwartung', 'Component servicing')
$WerteEn.Add('alle Eigenschaften', 'all properties')
$WerteEn.Add('Besitz übernehmen', 'Take ownership')
$WerteEn.Add('Domänencontroller', 'Domain Controllers')
$WerteEn.Add('Falsches Kennwort', 'Incorrect password')
$WerteEn.Add('Konten-Operatoren', 'Account Operators')
$WerteEn.Add('Nur dieser Ordner', 'Only this folder')
$WerteEn.Add('Domänen-Benutzer', 'Domain Users')
$WerteEn.Add('Format unbekannt', 'Format unknown')
$WerteEn.Add('Sonstige Dateien', 'Other files')
$WerteEn.Add('Ältester Eintrag', 'Oldest entry')
$WerteEn.Add('Domänencomputer', 'Domain Computers')
$WerteEn.Add('Konto unbekannt', 'Account unknown')
$WerteEn.Add('Standard (alle)', 'Default (all)')
$WerteEn.Add('Computerkonten', 'Computer accounts')
$WerteEn.Add('Konto gesperrt', 'Account locked')
$WerteEn.Add('Nicht erreicht', 'Not reached')
$WerteEn.Add('DHCP-Bereiche', 'DHCP scopes')
$WerteEn.Add('Dieser Ordner', 'This folder')
$WerteEn.Add('Gruppe/System', 'Group/system')
$WerteEn.Add('Remotedesktop', 'Remote Desktop')
$WerteEn.Add('Arbeitsplatz', 'Workstation')
$WerteEn.Add('Schemamaster', 'Schema master')
$WerteEn.Add('DHCP-Server', 'DHCP server')
$WerteEn.Add('Datenbanken', 'Databases')
$WerteEn.Add('Sonderkonto', 'Special account')
$WerteEn.Add('Unterordner', 'Subfolders')
$WerteEn.Add('Vollzugriff', 'Full control')
$WerteEn.Add('Zugriff von', 'Access from')
$WerteEn.Add('DNS-Server', 'DNS server')
$WerteEn.Add('Inaktiv ab', 'Inactive from')
$WerteEn.Add('Schreiben:', 'Write:')
$WerteEn.Add('Verweigern', 'Deny')
$WerteEn.Add('DNS-Zonen', 'DNS zones')
$WerteEn.Add('Schreiben', 'Write')
$WerteEn.Add('Sonstiges', 'Other')
$WerteEn.Add('Standorte', 'Sites')
$WerteEn.Add('Unbekannt', 'Unknown')
$WerteEn.Add('unbekannt', 'unknown')
$WerteEn.Add('(Wurzel)', '(root)')
$WerteEn.Add('Anlegen:', 'Create:')
$WerteEn.Add('Benutzer', 'User')
$WerteEn.Add('Erlauben', 'Allow')
$WerteEn.Add('Netzwerk', 'Network')
$WerteEn.Add('Verwaist', 'Orphaned')
$WerteEn.Add('verwaist', 'orphaned')
$WerteEn.Add('Dateien', 'Files')
$WerteEn.Add('Gruppen', 'Groups')
$WerteEn.Add('Anonym', 'Anonymous')
$WerteEn.Add('Domäne', 'Domain')
$WerteEn.Add('Gruppe', 'Group')
$WerteEn.Add('Umfang', 'Scope')
$WerteEn.Add('direkt', 'direct')
$WerteEn.Add('Ändern', 'Change')
$WerteEn.Add('Lesen', 'Read')
$WerteEn.Add('über', 'via')
$WerteMuster = @()
foreach ($k in ($WerteEn.Keys | Sort-Object Length -Descending)) { $WerteMuster += , @([regex]::new('(?<![\p{L}\d])' + [regex]::Escape($k) + '(?![\p{L}\d])'), $WerteEn[$k]) }
function WertAnzeige($v) {
    if (-not $script:En -or $v -isnot [string] -or -not $v) { return $v }
    if ($v.Contains('\')) { return $v }   # Kontonamen (DOMÄNE\Name) und Pfade bleiben, wie sie sind
    foreach ($m in $WerteMuster) { $v = $m[0].Replace($v, $m[1]) }
    $v
}
function SpaltenAnzeige($liste) {
    if (-not $script:En -or -not @($liste).Count) { return $liste }
    foreach ($o in @($liste)) {
        $n = [ordered]@{}
        foreach ($p in $o.PSObject.Properties) { $k = $p.Name; if ($SpaltenEn.ContainsKey($k)) { $k = $SpaltenEn[$k] }; $i = 2; $k0 = $k; while ($n.Contains($k)) { $k = "$k0 $i"; $i++ }; $n[$k] = WertAnzeige $p.Value }
        [pscustomobject]$n
    }
}
# Sortierschluessel fuer Nummern wie K2, K10: erst Buchstaben, dann die Zahl als Zahl
function IdZahl($id) { $m = [regex]::Match([string]$id, '\d+'); if ($m.Success) { [int]$m.Value } else { 0 } }
function IdText($id) { [regex]::Replace([string]$id, '\d+', '') }

$Befunde = New-Object System.Collections.Generic.List[object]

# Platzhalterzeilen ("nicht geprüft") erkennt das Skript nur an den vollständigen, festen Texten, die es selbst erzeugt - keine Teilstring-Suche
# in Konto-, Objekt- oder Beschreibungstexten, die ein Angreifer beeinflussen könnte
# Deutsch und Englisch: Skripte erzeugen den Text je nach -Language, erkannt werden beide Fassungen
$PlatzhalterTexte = @('Abfrage fehlgeschlagen - nicht geprüft', 'Berechtigungen nicht lesbar - nicht geprüft', 'SYSVOL nicht lesbar - nicht geprüft', 'WMI-Filter nicht lesbar - nicht geprüft', 'Kein Domänencontroller ermittelbar - SYSVOL nicht geprüft', 'Inhalt des Ordners nicht lesbar - Richtlinie nicht geprüft',
    'Query failed - not checked', 'Permissions not readable - not checked', 'SYSVOL not readable - not checked', 'WMI filter not readable - not checked', 'No domain controller found - SYSVOL not checked', 'Folder content not readable - policy not checked')
function IstPlatzhalter($o) {
    # Ein Skript kann $script:PlatzhalterSpalten setzen: dann zählen nur diese Spalten (Namensspalten sind von Benutzern beeinflussbar)
    foreach ($pr in $o.PSObject.Properties) {
        if ($script:PlatzhalterSpalten -and ($pr.Name -notin $script:PlatzhalterSpalten)) { continue }
        if (($pr.Value -is [string]) -and ($PlatzhalterTexte -ccontains $pr.Value)) { return $true }
    }
    ("$($o.Feststellung)" -clike 'Nicht geprüft - *') -or ("$($o.Feststellung)" -clike 'Not checked - *')
}

# Besteht ein Prüfpunkt nur aus Platzhalterzeilen (Abfrage fehlgeschlagen), ist das kein Mangel dieses Prüfpunkts: als Info führen statt
# mit seiner Stufe, damit es in den Kacheln nicht als hoher Befund zählt. Nur für Prüfpunkte, deren Ergebnis von einer Abfrage abhängt.
function InfoWennNurNichtGeprueft([string]$IdMuster) {
    foreach ($x in ($Befunde | Where-Object { $_.Id -match $IdMuster -and $_.Anzahl -gt 0 -and $_.Stufe -ne 'Info' })) {
        if (-not @($x.Objekte | Where-Object { -not (IstPlatzhalter $_) }).Count) { $x.Stufe = 'Info' }
    }
}

# Prüfpunkte, die für einzelne Server nicht ausgewertet werden konnten (Abfrage fehlgeschlagen), dürfen nicht als "ohne Befund" erscheinen:
# je Server eine Zeile "nicht geprüft" anhängen; hatte der Prüfpunkt sonst keinen Treffer, steht er als Info statt als OK.
function NichtGeprueft([string]$IdMuster, [string[]]$Namen, [string]$Verweis, [string]$Spalte = 'Server', [string]$Grund = $(T 'Abfrage nicht möglich' 'query not possible')) {
    foreach ($x in ($Befunde | Where-Object { $_.Id -match $IdMuster })) {
        $war = $x.Anzahl
        $zeilen = @($Namen | ForEach-Object { [pscustomobject][ordered]@{ $Spalte = $_; Feststellung = (T "Nicht geprüft - $Grund (siehe $Verweis)" "Not checked - $Grund (see $Verweis)"); Details = '' } })
        $x.Objekte = @($x.Objekte) + $zeilen
        $x.Anzahl = $x.Objekte.Count
        if (-not $war) { $x.Stufe = 'Info' }
    }
}
function Befund {
    param([string]$Id, [string]$Bereich, [string]$Titel, [ValidateSet('Hoch', 'Mittel', 'Niedrig', 'Info')][string]$Stufe,
        $Objekte, [string]$Erklaerung, [string]$Empfehlung)
    $liste = @($Objekte | Where-Object { $null -ne $_ })
    $status = $Stufe
    if ($liste.Count -eq 0) { $status = 'OK' }
    $Befunde.Add([pscustomobject]@{
            Id = $Id; Bereich = $Bereich; Titel = $Titel; Stufe = $status; Anzahl = $liste.Count
            Objekte = $liste; Erklaerung = $Erklaerung; Empfehlung = $Empfehlung
        })
}

# Schreibt den HTML-Bericht (und auf Wunsch CSV-Dateien). $Abschnitte: Überschrift -> Objekte oder fertiges HTML,
# erscheinen vor den Befunden. $Anhang: Überschrift -> Text, erscheint am Ende unverändert.
function BerichtSchreiben {
    param([string]$Name, [string]$Datei, [string]$Bezug, [string]$BezugName = $(T 'Domäne' 'Domain'), [System.Collections.IDictionary]$Umgebung,
        [System.Collections.IDictionary]$Abschnitte = @{}, [System.Collections.IDictionary]$Anhang = @{}, [string]$Fuss)
    Write-Host (T '  Bericht schreiben ...' '  Writing report ...')
    $jetzt = Get-Date
    $ziel = $script:Ausgabe
    if (-not $ziel) { $ziel = Join-Path (Get-Location).Path ("{0}_{1}_{2:yyyy-MM-dd_HHmm}.html" -f $Datei, ($Bezug -replace '[^\w.\-]+', '-'), $jetzt) }
    $ziel = [System.IO.Path]::GetFullPath($ziel)
    $ordner = Split-Path $ziel -Parent
    if (-not (Test-Path $ordner)) { New-Item -ItemType Directory -Path $ordner -Force | Out-Null }

    $stufeText = @{ Hoch = (T 'Hoch' 'High'); Mittel = (T 'Mittel' 'Medium'); Niedrig = (T 'Niedrig' 'Low'); Info = 'Info'; OK = 'OK' }
    $rang = @{ Hoch = 0; Mittel = 1; Niedrig = 2; Info = 3; OK = 4 }
    $zahl = @{}
    foreach ($s in 'Hoch', 'Mittel', 'Niedrig', 'Info', 'OK') { $zahl[$s] = @($Befunde | Where-Object Stufe -eq $s).Count }
    $css = @'
body{font-family:Segoe UI,Arial,sans-serif;color:#2b2f36;margin:0;background:#f4f6fa}
.seite{max-width:1100px;margin:0 auto;padding:28px 24px 60px}
h1{color:#1f5fae;margin:0 0 4px;font-size:28px} h2{color:#1f5fae;font-size:19px;margin:34px 0 10px;border-bottom:2px solid #d0d5dd;padding-bottom:6px}
.meta{color:#6b7280;font-size:13px;margin-bottom:22px}
.kacheln{display:flex;gap:14px;flex-wrap:wrap;margin:18px 0}
.kachel{flex:1;min-width:150px;background:#fff;border-radius:10px;padding:14px 18px;border-left:8px solid #999;box-shadow:0 1px 3px rgba(0,0,0,.08)}
.kachel b{display:block;font-size:30px} .kachel span{font-size:13px;color:#6b7280}
table{border-collapse:collapse;width:100%;background:#fff;font-size:13px;margin:8px 0 4px}
th{background:#1f5fae;color:#fff;text-align:left;padding:7px 9px;font-weight:600} td{padding:6px 9px;border-bottom:1px solid #e5e7eb;vertical-align:top;overflow-wrap:break-word}
tr:nth-child(even) td{background:#f9fafb}
.marke{display:inline-block;min-width:58px;text-align:center;padding:2px 9px;border-radius:11px;color:#fff;font-size:12px;font-weight:600}
.m-Hoch{background:#d92d20}.m-Mittel{background:#f79009}.m-Niedrig{background:#ca8a04}.m-OK{background:#12b76a}.m-Info{background:#1f5fae}
.befund{background:#fff;border-radius:10px;padding:16px 20px;margin:14px 0;border-left:8px solid #999;box-shadow:0 1px 3px rgba(0,0,0,.08)}
.befund h3{margin:0 0 6px;font-size:16px} .befund p{margin:5px 0;font-size:14px;line-height:1.45} .befund .emp{color:#14532d}
.Hoch{border-color:#d92d20}.Mittel{border-color:#f79009}.Niedrig{border-color:#eab308}.OK{border-color:#12b76a}.Info{border-color:#1f5fae}
.klein{color:#6b7280;font-size:12px} a{color:#1f5fae;text-decoration:none}
pre{background:#fff;border:1px solid #d0d5dd;border-radius:8px;padding:12px 14px;font-size:12px;white-space:pre-wrap;word-break:break-word}
.fuss{margin-top:40px;color:#6b7280;font-size:12px;border-top:1px solid #d0d5dd;padding-top:12px}
@media print{body{background:#fff}.befund,.kachel{box-shadow:none;border:1px solid #d0d5dd;border-left-width:8px;break-inside:avoid}}
'@
    $h = New-Object System.Text.StringBuilder
    [void]$h.Append("<!DOCTYPE html><html lang='$Language'><head><meta charset='utf-8'><title>$(Html $Name) $(Html $Bezug)</title><style>$css</style></head><body><div class='seite'>")
    [void]$h.Append("<h1>$(Html $Name)</h1><div class='meta'>$(Html $BezugName) <b>$(Html $Bezug)</b> · $(T 'erstellt am' 'created') $($jetzt.ToString((T 'dd.MM.yyyy HH:mm' 'yyyy-MM-dd HH:mm'))) $(T 'von' 'by') $(Html "$env:USERDOMAIN\$env:USERNAME") $(T 'auf' 'on') $(Html $env:COMPUTERNAME)</div>")
    [void]$h.Append("<div class='kacheln'>")
    foreach ($s in 'Hoch', 'Mittel', 'Niedrig', 'Info', 'OK') {
        $text = @{ Hoch = (T 'Befunde mit hoher Priorität' 'findings, high priority'); Mittel = (T 'mit mittlerer Priorität' 'medium priority'); Niedrig = (T 'mit niedriger Priorität' 'low priority'); Info = (T 'Hinweise und nicht geprüft' 'notes and not checked'); OK = (T 'Prüfpunkte ohne Befund' 'checks without findings') }[$s]
        [void]$h.Append("<div class='kachel $s'><b>$($zahl[$s])</b><span>$text</span></div>")
    }
    [void]$h.Append("</div><h2>$(T 'Umgebung' 'Environment')</h2><table><tr><th>$(T 'Eigenschaft' 'Property')</th><th>$(T 'Wert' 'Value')</th></tr>")
    foreach ($k in $Umgebung.Keys) { [void]$h.Append("<tr><td>$(Html (WertAnzeige $k))</td><td>$(Html (WertAnzeige $Umgebung[$k]))</td></tr>") }
    [void]$h.Append("</table>")

    [void]$h.Append("<h2>$(T 'Übersicht der Prüfpunkte' 'Overview of checks')</h2><table><tr><th>$(T 'Nr.' 'No.')</th><th>$(T 'Bereich' 'Area')</th><th>$(T 'Prüfpunkt' 'Check')</th><th>$(T 'Treffer' 'Hits')</th><th>$(T 'Bewertung' 'Rating')</th></tr>")
    foreach ($x in $Befunde) {
        [void]$h.Append("<tr><td>$($x.Id)</td><td>$(Html $x.Bereich)</td><td><a href='#$($x.Id)'>$(Html $x.Titel)</a></td><td>$($x.Anzahl)</td><td><span class='marke m-$($x.Stufe)'>$($stufeText[$x.Stufe])</span></td></tr>")
    }
    [void]$h.Append("</table>")

    foreach ($k in $Abschnitte.Keys) {
        [void]$h.Append("<h2>$(Html (WertAnzeige $k))</h2>")
        $inhalt = $Abschnitte[$k]
        if ($inhalt -is [string]) { [void]$h.Append($inhalt) }
        elseif (@($inhalt).Count) { [void]$h.Append(((Einheitlich $inhalt | ConvertTo-Html -Fragment) -join "`n")) }
        else { [void]$h.Append("<p class='klein'>$(T 'Keine Einträge.' 'No entries.')</p>") }
    }

    [void]$h.Append("<h2>$(T 'Befunde' 'Findings')</h2>")
    $offen = @($Befunde | Where-Object { $_.Stufe -ne 'OK' } | Sort-Object { $rang[$_.Stufe] }, { IdText $_.Id }, { IdZahl $_.Id })
    foreach ($x in $offen) {
        [void]$h.Append("<div class='befund $($x.Stufe)' id='$($x.Id)'><h3><span class='marke m-$($x.Stufe)'>$($stufeText[$x.Stufe])</span> $($x.Id) · $(Html $x.Titel) <span class='klein'>($($x.Anzahl) $(T 'Treffer' 'hits'))</span></h3>")
        [void]$h.Append("<p>$(Html $x.Erklaerung)</p><p class='emp'><b>$(T 'Empfehlung' 'Recommendation'):</b> $(Html $x.Empfehlung)</p>")
        [void]$h.Append(((Einheitlich $x.Objekte | ConvertTo-Html -Fragment) -join "`n"))
        [void]$h.Append("</div>")
    }
    if (-not $offen.Count) { [void]$h.Append("<p>$(T 'Keine Auffälligkeiten gefunden.' 'No anomalies found.')</p>") }
    $ok = @($Befunde | Where-Object Stufe -eq 'OK')
    if ($ok.Count) {
        [void]$h.Append("<h2>$(T 'Ohne Befund' 'No findings')</h2><table><tr><th>$(T 'Nr.' 'No.')</th><th>$(T 'Prüfpunkt' 'Check')</th></tr>")
        foreach ($x in $ok) { [void]$h.Append("<tr id='$($x.Id)'><td>$($x.Id)</td><td>$(Html $x.Titel)</td></tr>") }
        [void]$h.Append("</table>")
    }
    foreach ($k in $Anhang.Keys) { [void]$h.Append("<h2>$(Html $k)</h2><pre>$(Html $Anhang[$k])</pre>") }
    [void]$h.Append("<div class='fuss'>$(Html $Name) Version $script:Version · IT Admin Tools by MM · $(T 'Das Skript hat ausschließlich gelesen und nichts verändert.' 'The script only read and changed nothing.') $(Html $Fuss)</div>")
    [void]$h.Append("</div></body></html>")
    [System.IO.File]::WriteAllText($ziel, $h.ToString(), (New-Object System.Text.UTF8Encoding($true)))

    $basis = [System.IO.Path]::ChangeExtension($ziel, $null).TrimEnd('.')
    if ($script:Csv) {
        foreach ($x in ($Befunde | Where-Object { $_.Anzahl -gt 0 })) {
            Einheitlich $x.Objekte | Export-Csv -Path "${basis}_$($x.Id).csv" -NoTypeInformation -Delimiter ';' -Encoding UTF8
        }
    }
    Write-Host ''
    Write-Host ((T "Fertig: {0} hoch, {1} mittel, {2} niedrig, {3} ohne Befund." "Done: {0} high, {1} medium, {2} low, {3} without findings.") -f $zahl.Hoch, $zahl.Mittel, $zahl.Niedrig, $zahl.OK) -ForegroundColor Green
    Write-Host ((T 'Bericht' 'Report') + ": $ziel")
    if (-not $script:NichtOeffnen -and [Environment]::UserInteractive) { try { Start-Process $ziel } catch { } }
    $basis
}

# Platzhalterzeilen ("nicht geprüft") stehen in diesem Skript immer in der Spalte Feststellung. Namen von Freigaben, Ordnern, Dateien und Konten
# (von Benutzern beeinflussbar) werden nie als Platzhalter gewertet.
$script:PlatzhalterSpalten = @('Feststellung')

# ------------------------------------------------------------------ Support-Enddaten Windows Server (Stand 10.2026, Quelle: Microsoft)
# Spalten: Build ab | Name | Support-Ende
$ServerTabelle = @(
    @(26100, 'Windows Server 2025', '2034-11-14'),
    @(20348, 'Windows Server 2022', '2031-10-14'),
    @(17763, 'Windows Server 2019', '2029-01-09'),
    @(14393, 'Windows Server 2016', '2027-01-12'),
    @(9200,  'Windows Server 2012 / 2012 R2', '2023-10-10'),
    @(0,     'Windows Server 2008 R2 und älter', '2020-01-14')
)

# Bekannte Konten (nach SID, unabhängig von der Sprache des Systems)
$sidJeder = 'S-1-1-0'; $sidAuth = 'S-1-5-11'; $sidBenutzer = 'S-1-5-32-545'; $sidAnonym = 'S-1-5-7'
$weit = @($sidJeder, $sidAuth, $sidBenutzer, $sidAnonym)                                   # sehr große Personenkreise
$verwalter = @('S-1-5-18', 'S-1-5-19', 'S-1-5-20', 'S-1-5-32-544', 'S-1-3-0', 'S-1-5-32-549', 'S-1-5-32-551')   # System, Lokaler Dienst, Netzwerkdienst, Administratoren, Ersteller-Besitzer, Server-/Sicherungs-Operatoren
$dienstPraefix = @('S-1-5-80-', 'S-1-5-82-', 'S-1-5-83-', 'S-1-15-2-', 'S-1-5-5-', 'S-1-5-21-0-0-0')    # Dienst-, Anwendungs- und Sitzungskonten

function IstDomaenenBenutzer($sid) { $sid -match '^S-1-5-21-[\d-]+-513$' }
function IstVerwalter($sid) { ($sid -in $verwalter) -or ($sid -match '^S-1-5-21-[\d-]+-(512|519)$') }

# Dateisystemrechte als vorzeichenlose Zahl; generische Rechte (z. B. GENERIC_ALL in "nur Unterordner und Dateien"-Einträgen) in konkrete Rechte umsetzen
function RechteInt($r) {
    $u = [int64]([int]$r -band 0xFFFFFFFF)
    if ($u -band 0x10000000) { $u = $u -bor 0x1F01FF }
    if ($u -band 0x40000000) { $u = $u -bor 0x120116 }
    if ($u -band 0x80000000) { $u = $u -bor 0x120089 }
    if ($u -band 0x20000000) { $u = $u -bor 0x1200A0 }
    $u
}
function RechteKurz([int64]$u) {
    if (($u -band 0x1F01FF) -eq 0x1F01FF) { return 'Vollzugriff' }
    if (($u -band 0x1301BF) -eq 0x1301BF) { return 'Ändern' }
    if ($u -band 0x6) { return 'Schreiben' }
    if ($u -band 0xC0000) { return 'Berechtigungen/Besitz ändern' }
    if ($u -band 0x10040) { return (T 'Löschen' 'Delete') }
    'Lesen'
}
$rang = @{ 'Lesen' = 0; 'Schreiben' = 1; 'Ändern' = 2; 'Vollzugriff' = 3 }

# ------------------------------------------------------------------ Server bestimmen
Write-Host (T 'Dateiserver-Bericht - es wird nur gelesen, nichts verändert.' 'File server report - read-only, nothing is changed.') -ForegroundColor Cyan
$lokaleNamen = @($env:COMPUTERNAME, 'localhost', '.', "$env:COMPUTERNAME.$env:USERDNSDOMAIN")
try { $lokaleNamen += [System.Net.Dns]::GetHostEntry('').HostName } catch { }
$lokaleNamen = @($lokaleNamen | Where-Object { $_ -and $_ -notlike '*.' } | ForEach-Object { $_.ToLower() } | Select-Object -Unique)
if (-not $Server) { $Server = @($env:COMPUTERNAME) }
$jahrText = if ($AlterJahre -eq 1) { (T 'Jahr' 'year') } else { (T 'Jahren' 'years') }      # "seit über 5 Jahren"
$jahreText = if ($AlterJahre -eq 1) { (T 'Jahr' 'year') } else { (T 'Jahre' 'years') }      # "älter als 5 Jahre"
$sitzungen = @{}
function IstLokal($s) { "$s".ToLower() -in $lokaleNamen }
function Sitzung($s) {
    if (IstLokal $s) { return $null }
    if (-not $sitzungen.ContainsKey($s)) {
        $p = @{ ComputerName = $s; ErrorAction = 'Stop' }
        if ($Anmeldung) { $p.Credential = $Anmeldung }
        $sitzungen[$s] = New-CimSession @p
    }
    $sitzungen[$s]
}
function Cim($s, $klasse, $filter) {
    $p = @{ ClassName = $klasse; ErrorAction = 'Stop' }
    if ($filter) { $p.Filter = $filter }
    $cs = Sitzung $s
    if ($cs) { $p.CimSession = $cs }
    Get-CimInstance @p
}

# Namensauflösung der Konten in den Berechtigungen (je SID nur einmal). Lokale Konten eines entfernten Servers kennt nur der Server selbst:
# sie werden aus seiner eigenen Kontenliste ($z.Lokal) aufgelöst und nie als "verwaist" gemeldet, solange die Liste fehlt.
$sidCache = @{}
function Konto($sid, $z) {
    $lokalPrefix = $null
    if ($z -and $z.MaschinenSid) { $lokalPrefix = $z.MaschinenSid + '-' }
    if ($lokalPrefix -and $sid.StartsWith($lokalPrefix)) {
        $schl = "$($z.Name)|$sid"
        if ($sidCache.ContainsKey($schl)) { return $sidCache[$schl] }
        $e = [pscustomobject]@{ Name = $sid; Art = 'unbekannt'; Verwaist = $false }
        if ($z.Lokal -and $z.Lokal.ContainsKey($sid)) { $e = $z.Lokal[$sid] }
        elseif ($z.Lokal) { $e.Verwaist = $true; $e.Art = 'verwaist' }
        else { $e.Art = (T 'lokales Konto (nicht aufgelöst)' 'local account (not resolved)') }
        $sidCache[$schl] = $e
        return $e
    }
    if ($sidCache.ContainsKey($sid)) { return $sidCache[$sid] }
    $e = [pscustomobject]@{ Name = $sid; Art = 'unbekannt'; Verwaist = $false }
    try { $e.Name = ([System.Security.Principal.SecurityIdentifier]$sid).Translate([System.Security.Principal.NTAccount]).Value } catch { }
    $dienst = $false
    foreach ($pre in $dienstPraefix) { if ($sid.StartsWith($pre)) { $dienst = $true } }
    if ($e.Name -eq $sid -and -not $dienst -and ($sid -match '^S-1-5-21-')) { $e.Verwaist = $true; $e.Art = 'verwaist' }
    elseif ($dienst -or $sid -in $weit -or $sid -in $verwalter -or $sid -match '^S-1-5-32-') { $e.Art = 'Gruppe/System' }
    else {
        # Benutzer oder Gruppe? Domänenkonten über das Verzeichnis (Bindung per SID), lokale Konten über die lokale Benutzerverwaltung
        try {
            $de = [ADSI]"LDAP://<SID=$sid>"
            $oc = @($de.Properties['objectClass'] | ForEach-Object { "$_" })
            if ($oc -contains 'computer') { $e.Art = 'Computer' }
            elseif ($oc -contains 'group') { $e.Art = (T 'Gruppe' 'Group') }
            elseif ($oc -contains 'user') { $e.Art = 'Benutzer' }
        } catch { }
        if ($e.Art -eq 'unbekannt') {
            try { if (Get-LocalUser -SID $sid -ErrorAction Stop) { $e.Art = 'Benutzer' } } catch { }
        }
        if ($e.Art -eq 'unbekannt') {
            try { if (Get-LocalGroup -SID $sid -ErrorAction Stop) { $e.Art = (T 'Gruppe' 'Group') } } catch { }
        }
    }
    $sidCache[$sid] = $e
    $e
}

# ------------------------------------------------------------------ Daten lesen
$srvListe = @()          # Eigenschaften je Server
$nichtErreicht = @()     # Server, die nicht abgefragt werden konnten
$laufwerkeVerbunden = @()
$laufNr = 0
foreach ($s in $Server) {
    Write-Host (T "  Server $s ..." "Server $s ...")
    $z = [ordered]@{ Name = $s; Os = ''; Build = 0; Client = $false; OsFehler = $null; Smb = $null; Freigaben = @(); Zugriff = @{}; Pfad = @{}; Platten = @(); Schatten = @(); Volumes = @(); Fsrm = $null; Gast = $null; GastFehler = $false; Fehler = @(); Lokal = $null; MaschinenSid = $null; Verbindung = $null }
    try {
        $os = Cim $s Win32_OperatingSystem
        $z.Os = $os.Caption; $z.Build = [int]$os.BuildNumber; $z.Client = ([int]$os.ProductType -eq 1)
    } catch { $z.OsFehler = Kurz $_.Exception.Message; $z.Os = 'unbekannt' }
    $cs = $null
    try { $cs = Sitzung $s } catch { if (-not $z.OsFehler) { $z.OsFehler = Kurz $_.Exception.Message } }
    # Ohne Verbindung zu einem entfernten Server nie lokal weiterfragen (das würde die Daten des falschen Rechners unter dem Servernamen zeigen)
    if (-not (IstLokal $s) -and -not $cs) {
        $nichtErreicht += [pscustomobject][ordered]@{ Server = $s; Feststellung = (T 'Abfrage fehlgeschlagen - nicht geprüft' 'Query failed - not checked'); Details = $(if ($z.OsFehler) { $z.OsFehler } else { (T 'Keine Verbindung zum Server' 'No connection to the server') }) }
        continue
    }
    $cp = @{}; if ($cs) { $cp.CimSession = $cs }
    try { $z.Smb = Get-SmbServerConfiguration @cp -ErrorAction Stop } catch { $z.Fehler += (T "SMB-Einstellungen: $(Kurz $_.Exception.Message)" "SMB settings: $(Kurz $_.Exception.Message)") }
    $freigabenGelesen = $true
    try {
        $alle = @(Get-SmbShare @cp -ErrorAction Stop)
        $z.Freigaben = @($alle | Where-Object { -not $_.Special -and $_.Name -ne 'IPC$' })
        if ($Freigaben) { $z.Freigaben = @($z.Freigaben | Where-Object { $n = $_.Name; @($Freigaben | Where-Object { $n -like $_ }).Count }) }
    } catch { $freigabenGelesen = $false; $z.Fehler += "Freigaben: $(Kurz $_.Exception.Message)" }
    if ($z.OsFehler -and -not $freigabenGelesen -and -not $z.Smb) {
        $nichtErreicht += [pscustomobject][ordered]@{ Server = $s; Feststellung = (T 'Abfrage fehlgeschlagen - nicht geprüft' 'Query failed - not checked'); Details = $z.OsFehler }
        continue
    }
    # Ohne Administratorrechte liefert Get-SmbShare keine Pfade: Win32_Share nennt sie auch normalen Benutzern
    $ws = @{}
    try { foreach ($w in (Cim $s Win32_Share $null)) { $ws["$($w.Name)"] = "$($w.Path)" } } catch { }
    foreach ($f in $z.Freigaben) {
        $pf = "$($f.Path)"
        if (-not $pf -and $ws.ContainsKey($f.Name)) { $pf = $ws[$f.Name] }
        $z.Pfad[$f.Name] = $pf
        try { $z.Zugriff[$f.Name] = @(Get-SmbShareAccess -Name $f.Name @cp -ErrorAction Stop) } catch { $z.Zugriff[$f.Name] = $null }
    }
    try { $z.Platten = @(Cim $s Win32_LogicalDisk 'DriveType=3') } catch { $z.Platten = $null; $z.Fehler += "Volumes: $(Kurz $_.Exception.Message)" }
    try { $z.Schatten = @(Cim $s Win32_ShadowCopy $null) } catch { $z.Schatten = $null }
    try { $z.Volumes = @(Cim $s Win32_Volume $null) } catch { $z.Volumes = @() }
    try { $z.Gast = @(Cim $s Win32_UserAccount "LocalAccount=True AND SID LIKE '%-501'") | Select-Object -First 1 } catch { $z.GastFehler = $true }
    try { $z.Fsrm = [bool](Cim $s Win32_Service "Name='SrmSvc'") } catch { $z.Fsrm = $null }
    # Lokale Konten und Gruppen des Servers (für die Auflösung von Berechtigungen bei entfernten Servern)
    if (-not (IstLokal $s)) {
        try {
            $lk = @{}
            foreach ($u in (Cim $s Win32_UserAccount 'LocalAccount=True')) { $lk["$($u.SID)"] = [pscustomobject]@{ Name = "$($u.Domain)\$($u.Name)"; Art = 'Benutzer'; Verwaist = $false } }
            foreach ($g in (Cim $s Win32_Group 'LocalAccount=True')) { $lk["$($g.SID)"] = [pscustomobject]@{ Name = "$($g.Domain)\$($g.Name)"; Art = (T 'Gruppe' 'Group'); Verwaist = $false } }
            $erste = @($lk.Keys | Where-Object { $_ -match '^S-1-5-21-' } | Select-Object -First 1)
            if ($erste.Count) { $z.MaschinenSid = $erste[0] -replace '-\d+$', ''; $z.Lokal = $lk }
        } catch { }
        # Mit -Anmeldung: Dateizugriff über die administrativen Freigaben mit denselben Anmeldedaten verbinden (nur innerhalb dieses Skripts)
        if ($Anmeldung) {
            $lws = @($z.Pfad.Values | ForEach-Object { if ("$_" -match '^([A-Za-z]):') { $Matches[1].ToUpper() } } | Select-Object -Unique)
            foreach ($l in $lws) {
                try {
                    $laufNr++
                    New-PSDrive -Name "DSB$laufNr" -PSProvider FileSystem -Root "\\$s\$l`$" -Credential $Anmeldung -Scope Script -ErrorAction Stop | Out-Null
                    $laufwerkeVerbunden += "DSB$laufNr"
                } catch { $z.Fehler += (T "Dateizugriff mit -Anmeldung auf $l`$: $(Kurz $_.Exception.Message)" "File access with -Anmeldung on $l`$: $(Kurz $_.Exception.Message)") }
            }
        }
    }
    $srvListe += [pscustomobject]$z
}

# Liest Ordner oder Dateien unterhalb einer Wurzel selbst (statt Get-ChildItem -Recurse): bricht bei der Obergrenze sofort ab, folgt keinen Junctions
# und Symlinks (keine Schleifen, kein Ausbruch aus der Freigabe) und sammelt nicht lesbare Ordner samt Fehlertext.
function Durchlaufe([string]$Wurzel, [bool]$NurOrdner, [int]$MaxTiefe, [int]$Max) {
    $erg = New-Object System.Collections.Generic.List[object]
    $fehl = New-Object System.Collections.Generic.List[object]
    $links = 0; $gekappt = $false
    $stapel = New-Object System.Collections.Generic.Stack[object]
    $stapel.Push(@($Wurzel, 0))
    while ($stapel.Count -and -not $gekappt) {
        $e = $stapel.Pop(); $pfad = $e[0]; $t = $e[1]
        try {
            $di = New-Object System.IO.DirectoryInfo $pfad
            if (-not $NurOrdner) {
                foreach ($fi in $di.EnumerateFiles()) { if ($erg.Count -ge $Max) { $gekappt = $true; break }; $erg.Add($fi) }
            }
            if ($gekappt) { break }
            if ($t -lt $MaxTiefe) {
                foreach ($sub in $di.EnumerateDirectories()) {
                    if ($sub.Attributes -band [System.IO.FileAttributes]::ReparsePoint) { $links++; continue }
                    if ($NurOrdner) { if ($erg.Count -ge $Max) { $gekappt = $true; break }; $erg.Add($sub) }
                    $stapel.Push(@($sub.FullName, ($t + 1)))
                }
            }
        } catch { $fehl.Add([pscustomobject]@{ Pfad = $pfad; Meldung = Kurz $_.Exception.Message 160 }) }
    }
    [pscustomobject]@{ Eintraege = $erg; Fehler = $fehl; Verknuepfungen = $links; Gekappt = $gekappt }
}

# Pfad der Freigabe auf dem Server (aus Get-SmbShare, sonst Win32_Share); leer = ohne Administratorrechte nicht sichtbar
function SharePfad($z, $f) { "$($z.Pfad[$f.Name])" }
# Pfad, über den die Dateien gelesen werden: lokal der echte Pfad. Fernzugriff: über die administrative Freigabe des Laufwerks (z. B. \\Server\D$\Daten),
# damit die Freigaberechte den Zugriff nicht einschränken; ohne Laufwerksbuchstaben über die Freigabe selbst
function LeseWurzel($z, $f) {
    $pf = SharePfad $z $f
    if (-not $pf) { return '' }
    if (IstLokal $z.Name) { return $pf }
    if ($pf -match '^([A-Za-z]):\\?(.*)$') { return "\\$($z.Name)\$($Matches[1])`$\$($Matches[2])".TrimEnd('\') }
    "\\$($z.Name)\$($f.Name)"
}
function Relativ($pfad, $wurzel) { if ($pfad.Length -gt $wurzel.Length) { $pfad.Substring($wurzel.Length).TrimStart('\') } else { '(Wurzel)' } }

# ------------------------------------------------------------------ Freigaben auswerten
$f1 = @(); $f2 = @(); $f3 = @(); $f4 = @(); $f5 = @()
$uebersicht = @()
$ohnePfad = @()
$mehrfach = @()
$gesehenPfade = @{}
$ueberspringen = @{}     # Freigaben, deren Ordner nicht geprüft werden (Pfad unbekannt, fehlt oder schon über eine andere Freigabe erfasst)
foreach ($z in $srvListe) {
    foreach ($f in $z.Freigaben) {
        $zug = $z.Zugriff[$f.Name]
        $weitText = ''
        $standard = ($f.Name -in 'SYSVOL', 'NETLOGON')
        # Ohne Administratorrechte liefert Windows statt der echten Freigabe-ACL einen Platzhalter ("Everyone" Vollzugriff) und keinen
        # Sicherheitsdeskriptor: lesbar ist die ACL nur, wenn Get-SmbShare den Deskriptor mitliefert
        $zugLesbar = ($null -ne $zug) -and (@($zug).Count -gt 0) -and ("$($f.SecurityDescriptor)" -ne '')
        if (-not $zugLesbar) {
            $f1 += [pscustomobject][ordered]@{ Server = $z.Name; Freigabe = $f.Name; Konto = ''; Recht = ''; Feststellung = (T 'Berechtigungen nicht lesbar - nicht geprüft' 'Permissions not readable - not checked'); Details = (T 'Freigaberechte sind ohne Administratorrechte nicht (vollständig) lesbar' 'Share permissions are not (fully) readable without administrator rights') }
        } else {
            # Verweigern gegenrechnen: Ein Verweigern für dasselbe Konto mit gleich großem oder größerem Recht hebt den Eintrag auf
            $verw = @{}
            foreach ($a in $zug) {
                if ("$($a.AccessControlType)" -eq 'Deny') {
                    $rr = switch ("$($a.AccessRight)") { 'Full' { 3 } 'Change' { 2 } default { 0 } }
                    $verw["$($a.AccountName)"] = [math]::Max($rr, $(if ($verw.ContainsKey("$($a.AccountName)")) { $verw["$($a.AccountName)"] } else { 0 }))
                }
            }
            foreach ($a in $zug) {
                if ("$($a.AccessControlType)" -ne 'Allow') { continue }
                $k = $a.AccountName
                $sid = $null
                try { $sid = ([System.Security.Principal.NTAccount]$k).Translate([System.Security.Principal.SecurityIdentifier]).Value } catch { }
                if (-not $sid -and "$k" -match '^S-1-[\d-]+$') { $sid = "$k" }
                $recht = switch ("$($a.AccessRight)") { 'Full' { 'Vollzugriff' } 'Change' { 'Ändern' } default { 'Lesen' } }
                if (-not $sid) {
                    if ($recht -ne 'Lesen' -and -not $standard) { $f1 += [pscustomobject][ordered]@{ Server = $z.Name; Freigabe = $f.Name; Konto = $k; Recht = $recht; Feststellung = (T 'Abfrage fehlgeschlagen - nicht geprüft' 'Query failed - not checked'); Details = (T 'Konto nicht auflösbar' 'Account cannot be resolved') } }
                    continue
                }
                $aufgehoben = $verw.ContainsKey("$k") -and ($verw["$k"] -ge $rang[$recht])
                if ($sid -and ($sid -in $weit -or (IstDomaenenBenutzer $sid)) -and $recht -ne 'Lesen' -and -not $aufgehoben -and -not $standard) {
                    $f1 += [pscustomobject][ordered]@{ Server = $z.Name; Freigabe = $f.Name; Konto = $k; Recht = $recht; Feststellung = (T 'Freigabe erlaubt Schreiben für einen sehr großen Personenkreis' 'Share allows write access for a very large group of users') }
                }
                if ($sid -in $weit -or (IstDomaenenBenutzer $sid)) { $weitText += "$k ($recht); " }
            }
        }
        $pfad = SharePfad $z $f
        if (-not $pfad) {
            $ohnePfad += [pscustomobject][ordered]@{ Server = $z.Name; Freigabe = $f.Name }
            $ueberspringen["$($z.Name)|$($f.Name)"] = 'Pfad'
            $f5 += [pscustomobject][ordered]@{ Server = $z.Name; Freigabe = $f.Name; Pfad = ''; Feststellung = (T 'Berechtigungen nicht lesbar - nicht geprüft' 'Permissions not readable - not checked'); Details = (T 'Der Ordner der Freigabe ist nicht ermittelbar (Administratorrechte oder WMI-Zugriff fehlen)' 'The share''s folder cannot be determined (missing administrator rights or WMI access)') }
        } else {
            if ($pfad -match '^[A-Za-z]:\\?$' -or $pfad -match '^[A-Za-z]:\\(Windows|Windows\\System32|Program Files|Program Files \(x86\)|ProgramData|Users)\\?$') {
                $f2 += [pscustomobject][ordered]@{ Server = $z.Name; Freigabe = $f.Name; Pfad = $pfad; Feststellung = (T 'Freigabe zeigt auf ein ganzes Laufwerk oder einen Systemordner' 'Share points to a whole drive or a system folder') }
            }
            $wz = LeseWurzel $z $f
            $da = $true
            try { $da = [bool](Test-Path -LiteralPath $wz -ErrorAction Stop) } catch { $da = $null }
            if ($null -eq $da) {
                $f5 += [pscustomobject][ordered]@{ Server = $z.Name; Freigabe = $f.Name; Pfad = $pfad; Feststellung = (T 'Abfrage fehlgeschlagen - nicht geprüft' 'Query failed - not checked'); Details = 'Zugriff auf den Ordner nicht möglich' }
                $ueberspringen["$($z.Name)|$($f.Name)"] = 'Zugriff'
            } elseif (-not $da) {
                $f5 += [pscustomobject][ordered]@{ Server = $z.Name; Freigabe = $f.Name; Pfad = $pfad; Feststellung = 'Ordner der Freigabe existiert nicht mehr'; Details = '' }
                $ueberspringen["$($z.Name)|$($f.Name)"] = 'Fehlt'
            } else {
                $schl = "$($z.Name)|$($pfad.ToLower().TrimEnd('\'))"
                if ($gesehenPfade.ContainsKey($schl)) {
                    $mehrfach += "$($z.Name): $($f.Name) = $($gesehenPfade[$schl])"
                    $ueberspringen["$($z.Name)|$($f.Name)"] = 'Doppelt'
                } else { $gesehenPfade[$schl] = $f.Name }
            }
        }
        if ($f.Name.EndsWith('$')) {
            $f3 += [pscustomobject][ordered]@{ Server = $z.Name; Freigabe = $f.Name; Pfad = $pfad; Beschreibung = $f.Description }
        }
        if ("$($f.FolderEnumerationMode)" -eq 'Unrestricted') {
            $f4 += [pscustomobject][ordered]@{ Server = $z.Name; Freigabe = $f.Name; Pfad = $pfad; Beschreibung = $f.Description }
        }
        $uebersicht += [pscustomobject][ordered]@{
            Server = $z.Name; Freigabe = $f.Name; Pfad = $pfad; Beschreibung = $f.Description
            'Zugriff für große Gruppen' = $(if (-not $zugLesbar) { (T 'nicht lesbar' 'not readable') } elseif ($weitText) { $weitText.TrimEnd(' ', ';') } else { '-' })
            Verschlüsselt = JaNein $f.EncryptData; 'Zugriffsbasierte Aufzählung' = JaNein ("$($f.FolderEnumerationMode)" -eq 'AccessBased')
        }
    }
}
Befund 'F1' (T 'Freigaben' 'Shares') (T 'Freigaben mit Schreibrecht für Jeder, Authentifizierte Benutzer oder Benutzer' 'Shares with write permission for Everyone, Authenticated Users, or Users') 'Hoch' $f1 `
    (T 'Auf Freigabe-Ebene dürfen sehr viele Personen schreiben. Wirksam ist immer das strengere von Freigabe- und Dateisystemrechten; der Freigabeeintrag "Jeder" ist üblich, solange die Dateisystemrechte streng sind (siehe N1). Bleibt beides offen, kann jeder Mitarbeiter und im schlimmsten Fall Schadsoftware alle Dateien verändern oder verschlüsseln. Die Standardfreigaben SYSVOL und NETLOGON von Domänencontrollern werden nicht bewertet. Ein gleichzeitiges Verweigern für dasselbe Konto wird gegengerechnet.' 'At the share level, very many people are allowed to write. The stricter of share and file system permissions always takes effect; the share entry "Everyone" is common as long as file system permissions are strict (see N1). If both remain open, any employee and, worst case, malware can change or encrypt all files. The default shares SYSVOL and NETLOGON on domain controllers are not evaluated. A simultaneous deny for the same account is offset.') `
    (T 'Freigaberechte auf die zuständigen Gruppen beschränken, die eigentliche Steuerung über Dateisystemrechte mit Gruppen. "Jeder" und "Authentifizierte Benutzer" nicht mit Schreibrecht eintragen.' 'Restrict share permissions to the responsible groups; actual control is via file-system permissions using groups. Do not add "Everyone" and "Authenticated Users" with write permission.')
Befund 'F3' (T 'Freigaben' 'Shares') (T 'Versteckte Freigaben (Name endet auf $) außer den Standardfreigaben' 'Hidden shares (names ending in $) other than default shares') 'Info' $f3 `
    (T 'Versteckte Freigaben erscheinen nicht in der Netzwerkumgebung, sind aber per Name erreichbar. Sie dienen oft Sicherung oder Software und geraten leicht in Vergessenheit.' 'Hidden shares do not appear in Network Neighborhood, but are accessible by name. They are often used for backups or software and are easily forgotten.') `
    (T 'Zweck jeder Freigabe klären und dokumentieren; nicht mehr benötigte Freigaben entfernen.' 'Determine and document the purpose of each share; remove shares no longer needed.')

# ------------------------------------------------------------------ Dateisystemrechte
$n1 = @(); $n2 = @(); $n3 = @(); $n4 = @(); $n5 = @(); $n6 = @()
$ordnerGeprueft = 0
$gekuerztOrdner = @()
$verknuepfungen = @{}
foreach ($z in $srvListe) {
    foreach ($f in $z.Freigaben) {
        $stat = $ueberspringen["$($z.Name)|$($f.Name)"]
        if ($stat -eq 'Pfad' -or $stat -eq 'Zugriff') {
            $n1 += [pscustomobject][ordered]@{ Server = $z.Name; Freigabe = $f.Name; Ordner = ''; Konto = ''; Recht = ''; Feststellung = (T 'Berechtigungen nicht lesbar - nicht geprüft' 'Permissions not readable - not checked'); Details = $(if ($stat -eq 'Pfad') { (T 'Freigabepfad nicht ermittelbar (Administratorrechte oder WMI-Zugriff fehlen)' 'Share path cannot be determined (missing administrator rights or WMI access)') } else { 'Zugriff auf den Ordner nicht möglich' }) }
            continue
        }
        if ($stat) { continue }     # fehlender oder schon über eine andere Freigabe erfasster Ordner
        $wurzel = LeseWurzel $z $f
        Write-Host (T "  Rechte $($z.Name)\$($f.Name) ..." "  Permissions $($z.Name)\$($f.Name) ...")
        $ordner = New-Object System.Collections.Generic.List[string]
        $ordner.Add($wurzel)
        $lauf = Durchlaufe $wurzel $true $Tiefe ($MaxOrdner - 1)
        foreach ($d in $lauf.Eintraege) { $ordner.Add($d.FullName) }
        if ($lauf.Gekappt) { $gekuerztOrdner += "$($z.Name)\$($f.Name)" }
        if ($lauf.Verknuepfungen) { $verknuepfungen["$($z.Name)|$($f.Name)"] = [math]::Max($lauf.Verknuepfungen, [int]$verknuepfungen["$($z.Name)|$($f.Name)"]) }
        $platz = [ordered]@{}      # nicht lesbare Ordner (je Ordner nur eine Zeile)
        foreach ($fe in $lauf.Fehler) { $rl = Relativ $fe.Pfad $wurzel; if (-not $platz.Contains($rl)) { $platz[$rl] = $fe.Meldung } }
        $erste = $true
        $n4Zahl = 0
        foreach ($o in $ordner) {
            $ordnerGeprueft++
            $rel = Relativ $o $wurzel
            try { $acl = Get-Acl -LiteralPath $o -ErrorAction Stop } catch {
                if (-not $platz.Contains($rel)) { $platz[$rel] = Kurz $_.Exception.Message 160 }
                continue
            }
            if ($acl.AreAccessRulesProtected -and -not $erste) {
                $n4Zahl++
                if ($n4Zahl -le 50) { $n4 += [pscustomobject][ordered]@{ Server = $z.Name; Freigabe = $f.Name; Ordner = $rel; Beschreibung = (T 'Vererbung von oben abgeschaltet' 'Inheritance from above disabled') } }
                elseif ($n4Zahl -eq 51) { $n4 += [pscustomobject][ordered]@{ Server = $z.Name; Freigabe = $f.Name; Ordner = '...'; Beschreibung = (T 'weitere Ordner mit abgeschalteter Vererbung (Liste auf 50 je Freigabe gekürzt)' 'additional folders with inheritance disabled (list truncated to 50 per share)') } }
            }
            $erste = $false
            # Rechte je Konto vereinigen (mehrere Einträge für dasselbe Konto) und Verweigern gegenrechnen
            $erlaubt = [ordered]@{}; $verweigert = @{}
            foreach ($r in $acl.GetAccessRules($true, $false, [System.Security.Principal.SecurityIdentifier])) {
                $sid = $r.IdentityReference.Value
                $ri = RechteInt $r.FileSystemRights
                if ("$($r.AccessControlType)" -eq 'Deny') {
                    # Nur Verweigern-Einträge, die für den Ordner selbst gelten (nicht "nur Unterordner und Dateien"), werden gegengerechnet
                    if (-not ("$($r.PropagationFlags)" -match 'InheritOnly')) { $verweigert[$sid] = [int64]$verweigert[$sid] -bor $ri }
                } else { $erlaubt[$sid] = [int64]$erlaubt[$sid] -bor $ri }
            }
            foreach ($sid in $erlaubt.Keys) {
                $k = Konto $sid $z
                # Wirksame Rechte = erlaubte Bits ohne die verweigerten Bits desselben Kontos (am selben Ordner)
                $wirksam = [int64]$erlaubt[$sid]
                if ($verweigert.ContainsKey($sid)) { $wirksam = $wirksam -band (-bnot [int64]$verweigert[$sid]) }
                $recht = RechteKurz $wirksam
                if ($k.Verwaist) {
                    $n2 += [pscustomobject][ordered]@{ Server = $z.Name; Freigabe = $f.Name; Ordner = $rel; Konto = $sid; Recht = $recht; Hinweis = (T 'Konto existiert nicht mehr' 'Account no longer exists') }
                    continue
                }
                if (($sid -in $weit -or (IstDomaenenBenutzer $sid)) -and $recht -ne 'Lesen') {
                    $n1 += [pscustomobject][ordered]@{ Server = $z.Name; Freigabe = $f.Name; Ordner = $rel; Konto = $k.Name; Recht = $recht; Feststellung = (T 'Schreibrecht für einen sehr großen Personenkreis' 'Write permission for a very large group of users'); Details = '' }
                }
                if ($k.Art -eq 'Benutzer' -and $recht -ne 'Lesen' -and -not (IstVerwalter $sid)) {
                    $n3 += [pscustomobject][ordered]@{ Server = $z.Name; Freigabe = $f.Name; Ordner = $rel; Konto = $k.Name; Recht = $recht }
                }
                if ($recht -eq 'Vollzugriff' -and -not (IstVerwalter $sid) -and $sid -notin $weit -and $k.Art -ne 'Gruppe/System') {
                    $n5 += [pscustomobject][ordered]@{ Server = $z.Name; Freigabe = $f.Name; Ordner = $rel; Konto = $k.Name; Art = $k.Art }
                }
            }
            $osid = $null
            try { $osid = $acl.GetOwner([System.Security.Principal.SecurityIdentifier]).Value } catch { }
            if ($osid -and -not (IstVerwalter $osid) -and $osid -match '^S-1-5-21-') {
                $ok = Konto $osid $z
                if ($ok.Verwaist) { $n6 += [pscustomobject][ordered]@{ Server = $z.Name; Freigabe = $f.Name; Ordner = $rel; Besitzer = (T "$osid (Konto existiert nicht mehr)" "$osid (account no longer exists)") } }
                elseif ($ok.Art -eq 'Benutzer') { $n6 += [pscustomobject][ordered]@{ Server = $z.Name; Freigabe = $f.Name; Ordner = $rel; Besitzer = $ok.Name } }
            }
        }
        foreach ($pk in $platz.Keys) {
            $n1 += [pscustomobject][ordered]@{ Server = $z.Name; Freigabe = $f.Name; Ordner = $pk; Konto = ''; Recht = ''; Feststellung = (T 'Berechtigungen nicht lesbar - nicht geprüft' 'Permissions not readable - not checked'); Details = $platz[$pk] }
        }
    }
}
Befund 'N2' (T 'Dateisystemrechte' 'File system permissions') (T 'Berechtigungen für nicht mehr vorhandene Konten (verwaiste SIDs)' 'Permissions for accounts no longer present (orphaned SIDs)') 'Mittel' $n2 `
    (T 'Die Einträge gehören zu gelöschten Konten. Sie bewirken nichts, sind aber Unordnung und ein Risiko, falls eine SID später wiederverwendet wird (z. B. nach einer Domänenmigration).' 'The entries belong to deleted accounts. They have no effect, but are clutter and a risk if a SID is later reused (e.g., after a domain migration).') `
    (T 'Verwaiste Einträge aufräumen (Sicherheitseinstellungen des Ordners) und Konten künftig erst nach dem Entfernen der Rechte löschen.' 'Clean up orphaned entries (folder security settings) and delete accounts only after removing permissions going forward.')

# ------------------------------------------------------------------ Server-Einstellungen und Volumes
$s1 = @(); $s2 = @(); $s3 = @(); $s4 = @(); $s5 = @(); $s6 = @(); $s7 = @(); $s8 = @()
$nichtBewertet = @()
foreach ($z in $srvListe) {
    if ($z.Smb) {
        if ($z.Smb.EnableSMB1Protocol) { $s1 += [pscustomobject][ordered]@{ Server = $z.Name; Feststellung = (T 'SMB1 ist aktiv' 'SMB1 is enabled'); Details = '' } }
        if (-not $z.Smb.RequireSecuritySignature) { $s2 += [pscustomobject][ordered]@{ Server = $z.Name; Feststellung = (T 'SMB-Signierung wird nicht verlangt' 'SMB signing is not required'); Details = (T "Signierung angeboten: $(JaNein $z.Smb.EnableSecuritySignature)" "Signing offered: $(JaNein $z.Smb.EnableSecuritySignature)") } }
        $ns = @($z.Smb.NullSessionShares | Where-Object { $_ })
        if ($ns.Count) { $s3 += [pscustomobject][ordered]@{ Server = $z.Name; Feststellung = (T 'Freigaben ohne Anmeldung erreichbar (Null-Sitzung)' 'Shares accessible without logon (null session)'); Details = ($ns -join ', ') } }
    } else {
        $grund = (($z.Fehler | Where-Object { $_ -like 'SMB*' }) -join ' ')
        $pl = [pscustomobject][ordered]@{ Server = $z.Name; Feststellung = (T 'Abfrage fehlgeschlagen - nicht geprüft' 'Query failed - not checked'); Details = $grund }
        $s1 += $pl; $s2 += $pl; $s3 += $pl
    }
    if ($z.GastFehler) { $s4 += [pscustomobject][ordered]@{ Server = $z.Name; Feststellung = (T 'Abfrage fehlgeschlagen - nicht geprüft' 'Query failed - not checked'); Details = (T 'Lokale Benutzer des Servers nicht abfragbar (Administratorrechte oder WMI-Zugriff fehlen)' 'Local users on the server cannot be queried (missing administrator rights or WMI access)') } }
    elseif ($z.Gast -and -not $z.Gast.Disabled) { $s4 += [pscustomobject][ordered]@{ Server = $z.Name; Feststellung = (T 'Lokales Gastkonto ist aktiviert' 'Local guest account is enabled'); Details = $z.Gast.Name } }
    # Volumes mit Freigaben
    $laufwerke = @($z.Pfad.Values | ForEach-Object { if ("$_" -match '^([A-Za-z]):') { $Matches[1].ToUpper() + ':' } } | Select-Object -Unique)
    if ($null -eq $z.Platten -and ($laufwerke.Count -or $z.Freigaben.Count)) { $s5 += [pscustomobject][ordered]@{ Server = $z.Name; Volume = ''; 'Frei (GB)' = ''; 'Gesamt (GB)' = ''; 'Frei (%)' = ''; Feststellung = (T 'Abfrage fehlgeschlagen - nicht geprüft' 'Query failed - not checked') } }
    if ($null -eq $z.Schatten -and -not $laufwerke.Count -and $z.Freigaben.Count) { $s6 += [pscustomobject][ordered]@{ Server = $z.Name; Volume = ''; Feststellung = (T 'Abfrage fehlgeschlagen - nicht geprüft' 'Query failed - not checked') } }
    foreach ($lw in $laufwerke) {
        $pl = @($z.Platten) | Where-Object { $_.DeviceID -eq $lw } | Select-Object -First 1
        if ($pl -and $pl.Size) {
            $proz = [math]::Round(100 * $pl.FreeSpace / $pl.Size, 1)
            if ($proz -lt $FreiProzent) { $s5 += [pscustomobject][ordered]@{ Server = $z.Name; Volume = $lw; 'Frei (GB)' = [math]::Round($pl.FreeSpace / 1GB, 1); 'Gesamt (GB)' = [math]::Round($pl.Size / 1GB, 1); 'Frei (%)' = $proz } }
        }
        if ($null -eq $z.Schatten) {
            $s6 += [pscustomobject][ordered]@{ Server = $z.Name; Volume = $lw; Feststellung = (T 'Abfrage fehlgeschlagen - nicht geprüft' 'Query failed - not checked') }
        } else {
            $vol = @($z.Volumes | Where-Object { $_.DriveLetter -eq $lw } | Select-Object -First 1)
            $anz = 0
            if ($vol.Count -and $vol[0].DeviceID) { $anz = @($z.Schatten | Where-Object { $_.VolumeName -eq $vol[0].DeviceID }).Count }
            if (-not $anz -and $lw -ne 'C:') { $s6 += [pscustomobject][ordered]@{ Server = $z.Name; Volume = $lw; Feststellung = (T 'Keine Schattenkopien (Vorgängerversionen)' 'No shadow copies (previous versions)') } }
        }
    }
    if ($null -eq $z.Fsrm) { $s7 += [pscustomobject][ordered]@{ Server = $z.Name; Feststellung = (T 'Abfrage fehlgeschlagen - nicht geprüft' 'Query failed - not checked'); Details = '' } }
    elseif ($z.Fsrm -eq $false -and $z.Freigaben.Count) { $s7 += [pscustomobject][ordered]@{ Server = $z.Name; Feststellung = (T 'Kein Ressourcen-Manager für Dateiserver (Kontingente und Dateiprüfung) installiert' 'No File Server Resource Manager installed (quotas and file screening)'); Details = '' } }
    # Betriebssystem
    if ($z.OsFehler) {
        $s8 += [pscustomobject][ordered]@{ Server = $z.Name; Betriebssystem = ''; 'Support-Ende' = ''; Feststellung = (T 'Abfrage fehlgeschlagen - nicht geprüft' 'Query failed - not checked'); Details = $z.OsFehler }
    } elseif ($z.Client) {
        $nichtBewertet += "$($z.Name) ($($z.Os))"
    } else {
        $tab = $ServerTabelle | Where-Object { $z.Build -ge $_[0] } | Select-Object -First 1
        if ($tab) {
            $ende = [datetime]$tab[2]
            if ($ende -lt (Get-Date)) { $s8 += [pscustomobject][ordered]@{ Server = $z.Name; Betriebssystem = $tab[1]; 'Support-Ende' = Datum $ende; Feststellung = (T 'Support abgelaufen' 'Support expired'); Details = '' } }
            elseif ($ende -lt (Get-Date).AddDays(365)) { $s8 += [pscustomobject][ordered]@{ Server = $z.Name; Betriebssystem = $tab[1]; 'Support-Ende' = Datum $ende; Feststellung = (T 'Support endet innerhalb eines Jahres' 'Support ends within a year'); Details = '' } }
        }
    }
}
Befund 'S1' 'Server' (T 'SMB1 aktiv' 'SMB1 enabled') 'Hoch' $s1 `
    (T 'SMB1 ist ein veraltetes Protokoll mit schweren Sicherheitslücken (u. a. von der Ransomware WannaCry ausgenutzt). Heutige Windows-Versionen brauchen es nicht mehr.' 'SMB1 is an outdated protocol with severe security vulnerabilities (among other things, exploited by the WannaCry ransomware). Current Windows versions no longer require it.') `
    (T 'SMB1 abschalten, vorher mit der Überwachung (Audit) prüfen, ob noch alte Geräte (Scanner, NAS) es verwenden.' 'Disable SMB1; first check with monitoring (audit) whether older devices (scanners, NAS) still use it.')
Befund 'S2' 'Server' (T 'SMB-Signierung nicht erforderlich' 'SMB signing not required') 'Mittel' $s2 `
    (T 'Ohne verlangte Signierung können Angreifer im Netzwerk Datenverkehr manipulieren oder umleiten (Relay-Angriffe).' 'Without required signing, attackers can manipulate or redirect network traffic (relay attacks).') `
    (T 'SMB-Signierung verlangen (Gruppenrichtlinie "Microsoft-Netzwerk (Server): Kommunikation digital signieren (immer)"), vorher Kompatibilität alter Geräte prüfen.' 'Require SMB signing (Group Policy "Microsoft network server: Digitally sign communications (always)"), check compatibility of older devices first.')

# ------------------------------------------------------------------ Datenbestand
$d1 = @(); $d2 = @(); $d3 = @(); $d4 = @(); $d5 = @(); $dHinweis = @()
$endungen = @('.exe', '.bat', '.cmd', '.ps1', '.vbs', '.msi', '.scr', '.com', '.js', '.jar')
$grenzeAlt = (Get-Date).AddYears(-$AlterJahre)
$grenzeGross = [int64]$GrosseDateiGB * 1GB
$gesamt = @()
if (-not $OhneDateien) {
    foreach ($z in $srvListe) {
        foreach ($f in $z.Freigaben) {
            $stat = $ueberspringen["$($z.Name)|$($f.Name)"]
            if ($stat -eq 'Pfad' -or $stat -eq 'Zugriff') {
                $dHinweis += [pscustomobject][ordered]@{ Server = $z.Name; Freigabe = $f.Name; Feststellung = (T 'Abfrage fehlgeschlagen - nicht geprüft' 'Query failed - not checked'); Details = $(if ($stat -eq 'Pfad') { (T 'Freigabepfad nicht ermittelbar (Administratorrechte oder WMI-Zugriff fehlen)' 'Share path cannot be determined (missing administrator rights or WMI access)') } else { 'Zugriff auf den Ordner nicht möglich' }) }
                continue
            }
            if ($stat) { continue }
            $wurzel = LeseWurzel $z $f
            $pfadSrv = SharePfad $z $f
            Write-Host (T "  Dateien $($z.Name)\$($f.Name) ..." "  Files $($z.Name)\$($f.Name) ...")
            $n = 0; $groesse = 0L; $altN = 0; $altG = 0L
            $exe = @{}; $gross = @(); $pw = @(); $lang = @()
            $lauf = Durchlaufe $wurzel $false ([int]::MaxValue) $MaxDateien
            $erreicht = $lauf.Gekappt
            foreach ($d in $lauf.Eintraege) {
                $n++; $groesse += $d.Length
                if ($d.LastWriteTime -lt $grenzeAlt) { $altN++; $altG += $d.Length }
                if ($d.Length -ge $grenzeGross) { $gross += $d }
                $ext = $d.Extension.ToLower()
                if ($ext -in $endungen) { $exe[$ext] = 1 + $exe[$ext] }
                if ($d.Name -match '(?i)passw|kennwort|zugangsdaten|credential') { $pw += $d }
                # Pfadlänge auf dem Server (bei Fernzugriff ohne den Vorsatz \\Server\Laufwerk$)
                if (($pfadSrv.TrimEnd('\').Length + ($d.FullName.Length - $wurzel.TrimEnd('\').Length)) -ge 240) { $lang += $d }
            }
            if ($lauf.Fehler.Count) {
                $beispiel = ($lauf.Fehler | Select-Object -First 1).Meldung
                $dHinweis += [pscustomobject][ordered]@{ Server = $z.Name; Freigabe = $f.Name; Feststellung = (T 'Abfrage fehlgeschlagen - nicht geprüft' 'Query failed - not checked'); Details = (T "$($lauf.Fehler.Count) Ordner nicht lesbar, z. B.: $beispiel" "$($lauf.Fehler.Count) folders cannot be read, e.g.: $beispiel") }
            }
            if ($lauf.Verknuepfungen) { $verknuepfungen["$($z.Name)|$($f.Name)"] = [math]::Max($lauf.Verknuepfungen, [int]$verknuepfungen["$($z.Name)|$($f.Name)"]) }
            if ($erreicht) { $dHinweis += [pscustomobject][ordered]@{ Server = $z.Name; Freigabe = $f.Name; Feststellung = (T 'Nur die ersten Dateien ausgewertet' 'Only the first files evaluated'); Details = (T "Grenze $MaxDateien Dateien erreicht (-MaxDateien erhöhen)" "$MaxDateien file limit reached (-MaxDateien increase)") } }
            $gesamt += [pscustomobject][ordered]@{ Server = $z.Name; Freigabe = $f.Name; Dateien = $n; 'Größe (GB)' = [math]::Round($groesse / 1GB, 2); (T "Älter als $AlterJahre $jahreText" "Older than $AlterJahre $jahreText") = $altN; 'davon Größe (GB)' = [math]::Round($altG / 1GB, 2) }
            if ($altN) { $d1 += [pscustomobject][ordered]@{ Server = $z.Name; Freigabe = $f.Name; Dateien = $altN; 'Größe (GB)' = [math]::Round($altG / 1GB, 2); 'Anteil (%)' = [math]::Round(100 * $altN / [math]::Max($n, 1), 1) } }
            foreach ($g in ($gross | Sort-Object Length -Descending | Select-Object -First 20)) {
                $d2 += [pscustomobject][ordered]@{ Server = $z.Name; Freigabe = $f.Name; Datei = $g.FullName.Substring($wurzel.Length).TrimStart('\'); 'Größe (GB)' = [math]::Round($g.Length / 1GB, 2); Geändert = Datum $g.LastWriteTime }
            }
            foreach ($k in $exe.Keys) { $d3 += [pscustomobject][ordered]@{ Server = $z.Name; Freigabe = $f.Name; Dateityp = $k; Anzahl = $exe[$k] } }
            foreach ($p in ($pw | Select-Object -First 25)) { $d4 += [pscustomobject][ordered]@{ Server = $z.Name; Freigabe = $f.Name; Datei = $p.FullName.Substring($wurzel.Length).TrimStart('\'); Geändert = Datum $p.LastWriteTime } }
            foreach ($l in ($lang | Select-Object -First 25)) { $d5 += [pscustomobject][ordered]@{ Server = $z.Name; Freigabe = $f.Name; Datei = Kurz ($l.FullName.Substring($wurzel.Length).TrimStart('\')) 120; 'Pfadlänge' = $pfadSrv.TrimEnd('\').Length + ($l.FullName.Length - $wurzel.TrimEnd('\').Length) } }
        }
    }
}
Befund 'D3' (T 'Datenbestand' 'Data volume') (T 'Ausführbare Dateien und Skripte in Datenfreigaben' 'Executable files and scripts in data shares') 'Niedrig' $d3 `
    (T 'Programme und Skripte auf Datenfreigaben (.exe, .bat, .ps1, .vbs, .msi ...) sind oft Installationsreste. Schreibbare Freigaben mit ausführbaren Dateien sind ein typischer Weg, Schadsoftware zu verteilen.' 'Programs and scripts on data shares (.exe, .bat, .ps1, .vbs, .msi ...) are often installation leftovers. Writable shares with executable files are a typical way to distribute malware.') `
    (T 'Zweck prüfen. Installationsmedien auf eine schreibgeschützte Software-Freigabe legen, Reste löschen.' 'Check purpose. Place installation media on a read-only software share, delete leftovers.')

# Server, die nicht abgefragt werden konnten, und gekürzte Listen dürfen nicht als "ohne Befund" erscheinen
if ($nichtErreicht.Count) {
    $namen = @($nichtErreicht | ForEach-Object { $_.Server })
    NichtGeprueft '^[FNSD]\d+$' $namen (T 'Abschnitt "Nicht abgefragte Server"' 'Section "Servers not queried"') 'Server' (T 'Server nicht abfragbar' 'Server not queryable')
}
foreach ($x in ($Befunde | Where-Object { $_.Id -match '^N\d$' })) {
    foreach ($g in $gekuerztOrdner) {
        $x.Objekte = @($x.Objekte) + [pscustomobject][ordered]@{ Server = ''; Freigabe = $g; Ordner = ''; Konto = ''; Recht = ''; Feststellung = (T 'Nur die ersten Ordner geprüft - Grenze erreicht (-MaxOrdner erhöhen)' 'Only the first folders checked - limit reached (increase -MaxOrdner)'); Details = '' }
        $x.Anzahl = @($x.Objekte).Count; if ($x.Stufe -eq 'OK') { $x.Stufe = 'Info' }
    }
}
if ($dHinweis.Count) {
    foreach ($x in ($Befunde | Where-Object { $_.Id -match '^D[1-5]$' -and -not $OhneDateien })) {
        $x.Objekte = @($x.Objekte) + @($dHinweis | ForEach-Object { [pscustomobject][ordered]@{ Server = $_.Server; Freigabe = $_.Freigabe; Feststellung = $_.Feststellung; Details = $_.Details } })
        $x.Anzahl = @($x.Objekte).Count; if ($x.Stufe -eq 'OK') { $x.Stufe = 'Info' }
    }
}
InfoWennNurNichtGeprueft '^(F1|F5|N[1-6]|S[1-8])$'

# ------------------------------------------------------------------ Bericht
$umgebung = [ordered]@{
    'Server'                    = ($srvListe | ForEach-Object { "$($_.Name) ($($_.Os))" }) -join '; '
    (T 'Freigaben ausgewertet' 'Shares evaluated')     = @($uebersicht).Count
    (T 'Ordner auf Rechte geprüft' 'Folders checked for permissions') = (T "$ordnerGeprueft (Tiefe $Tiefe, höchstens $MaxOrdner je Freigabe)" "$ordnerGeprueft (depth $Tiefe, at most $MaxOrdner per share)")
    'Dateien'                   = $(if ($OhneDateien) { (T 'nicht ausgewertet (-OhneDateien)' 'not evaluated (-OhneDateien)') } else { (T "höchstens $MaxDateien je Freigabe, alt = älter als $AlterJahre $jahreText" "at most $MaxDateien per share, old = older than $AlterJahre $jahreText") })
    (T 'Ausgeführt als' 'Run as')            = "$env:USERDOMAIN\$env:USERNAME"
}
if ($Anmeldung -and $laufwerkeVerbunden.Count) { $umgebung[(T 'Dateizugriff' 'File access')] = (T "bei entfernten Servern als $($Anmeldung.UserName)" "on remote servers as $($Anmeldung.UserName)") }
if ($verknuepfungen.Count) { $umgebung[(T 'Übersprungene Verknüpfungen (Junctions/Symlinks)' 'Skipped links (junctions/symlinks)')] = (($verknuepfungen.Keys | Sort-Object | ForEach-Object { "$($_ -replace '^.*\|', ''): $($verknuepfungen[$_])" }) -join '; ') }
if ($mehrfach.Count) { $umgebung[(T 'Mehrfach freigegebene Ordner (nur einmal gezählt)' 'Folders shared multiple times (counted once)')] = $mehrfach -join '; ' }
if ($ohnePfad.Count) { $umgebung[(T 'Freigabepfade nicht sichtbar' 'Share paths not visible')] = (T "$($ohnePfad.Count) Freigaben (ohne Administratorrechte oder WMI-Zugriff liefert Windows keine Pfade)" "$($ohnePfad.Count) shares (Windows does not return paths without administrator rights or WMI access)") }
if ($nichtBewertet.Count) { $umgebung[(T 'Kein Serverbetriebssystem (Support nicht bewertet)' 'No server operating system (support not assessed)')] = $nichtBewertet -join '; ' }
if ($nichtErreicht.Count) { $umgebung[(T 'Nicht erreicht' 'Not reached')] = ($nichtErreicht | ForEach-Object { $_.Server }) -join ', ' }
$abschn = [ordered]@{ (T 'Freigaben im Überblick' 'Shares overview') = @($uebersicht) }
if ($nichtErreicht.Count) { $abschn[(T 'Nicht abgefragte Server' 'Servers not queried')] = @($nichtErreicht) }
if ($gesamt.Count) { $abschn[(T 'Datenbestand je Freigabe' 'Data per share')] = $gesamt }
$bezug = ($srvListe | ForEach-Object { $_.Name }) -join '_'
if (-not $bezug) { $bezug = ($Server -join '_') }
$basis = BerichtSchreiben -Name (T 'Dateiserver-Bericht Light' 'File server report Light') -Datei 'Dateiserver-Bericht' -Bezug $bezug -BezugName 'Server' -Umgebung $umgebung -Abschnitte $abschn `
    -Fuss (T 'Light-Version mit 6 von 24 Prüfpunkten der Vollversion (https://mueller86.gumroad.com/l/wbbwj). Grundlage sind die Freigaben, Berechtigungen und Dateieigenschaften zum Zeitpunkt der Abfrage; Dateiinhalte wurden nicht gelesen. Der Bericht ersetzt keine Sicherheitsprüfung und keine Datensicherung.' 'Light version with 6 of 24 checks of the full version (https://mueller86.gumroad.com/l/wbbwj). The basis is the shares, permissions, and file properties at the time of query; file contents were not read. The report does not replace a security check or a data backup.')
if ($Csv) {
    if ($uebersicht.Count) { Einheitlich $uebersicht | Export-Csv -Path "${basis}_Freigaben.csv" -NoTypeInformation -Delimiter ';' -Encoding UTF8 }
    if ($gesamt.Count) { Einheitlich $gesamt | Export-Csv -Path "${basis}_Datenbestand.csv" -NoTypeInformation -Delimiter ';' -Encoding UTF8 }
}
foreach ($cs in $sitzungen.Values) { try { Remove-CimSession -CimSession $cs } catch { } }
foreach ($l in $laufwerkeVerbunden) { try { Remove-PSDrive -Name $l -Force -ErrorAction SilentlyContinue } catch { } }
if ($PassThru) { $Befunde }
