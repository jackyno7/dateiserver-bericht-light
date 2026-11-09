#Requires -Version 5.1
<#
.SYNOPSIS
    LIGHT-VERSION (6 von 24 Prüfpunkten, kostenlos). Dateiserver-Bericht: prüft Freigaben, Dateisystemrechte, SMB-Einstellungen und Datenbestand eines oder mehrerer
    Windows-Dateiserver und erzeugt einen HTML-Bericht mit Auffälligkeiten.

.DESCRIPTION
    Das Skript arbeitet ausschließlich LESEND. Es ändert weder Freigaben noch Berechtigungen, Dateien oder Einstellungen und
    überträgt keine Daten. Es schreibt genau eine HTML-Datei (und auf Wunsch CSV-Dateien) in den angegebenen Ausgabeordner.

    Voraussetzungen: Windows PowerShell 5.1 und ein Konto mit Leserechten auf Freigaben und Dateisystem. Am genauesten wird die
    Auswertung, wenn das Skript direkt auf dem Dateiserver mit Administratorrechten läuft. Mit -Server können weitere Server
    über das Netzwerk geprüft werden (Remote-Verwaltung per CIM/WinRM, Dateien über die administrativen Freigaben wie C$).
    Was das Konto nicht lesen darf, erscheint als "nicht geprüft" und nie als "in Ordnung".

    Geprüft werden: Freigaben (Zugriff für Jeder, Systemordner, versteckte Freigaben, verwaiste Freigaben), Dateisystemrechte der
    obersten Ordnerebenen jeder Freigabe (Schreibrechte für Jeder, verwaiste Konten, Rechte direkt an Benutzern, abgeschaltete
    Vererbung), SMB-Einstellungen (SMB1, Signierung, anonymer Zugriff), Volumes (freier Platz, Schattenkopien, Kontingente),
    das Betriebssystem und den Datenbestand (alte und große Dateien, ausführbare Dateien, auffällige Dateinamen).

.PARAMETER Server
    Ein oder mehrere Dateiserver (Name). Standard: der lokale Computer.

.PARAMETER Freigaben
    Nur diese Freigaben auswerten (Namen, Platzhalter erlaubt). Standard: alle Freigaben außer den Standardfreigaben (C$, ADMIN$, IPC$).

.PARAMETER Tiefe
    Wie viele Ordnerebenen unterhalb jeder Freigabe die Dateisystemrechte geprüft werden (Standard 3).

.PARAMETER MaxOrdner
    Höchstens so viele Ordner je Freigabe auf Rechte prüfen (Standard 3000). Bei Erreichen steht ein Hinweis im Bericht.

.PARAMETER MaxDateien
    Höchstens so viele Dateien je Freigabe für die Datenauswertung lesen (Standard 50000). Bei Erreichen steht ein Hinweis im Bericht.

.PARAMETER OhneDateien
    Keine Dateien zählen (nur Freigaben, Rechte und Einstellungen). Deutlich schneller auf großen Servern.

.PARAMETER AlterJahre
    Dateien, die so viele Jahre nicht geändert wurden, gelten als alt (Standard 5).

.PARAMETER GrosseDateiGB
    Dateien ab dieser Größe in GB werden aufgelistet (Standard 1).

.PARAMETER FreiProzent
    Warnt, wenn auf einem Volume mit Freigaben weniger Prozent frei sind (Standard 15).

.PARAMETER Ausgabe
    Pfad der HTML-Datei. Standard: Dateiserver-Bericht_<Server>_<Datum>.html im aktuellen Ordner.

.PARAMETER Anmeldung
    Anmeldedaten (Get-Credential) für die Abfrage entfernter Server.

.PARAMETER Csv
    Schreibt je Prüfpunkt mit Treffern eine CSV-Datei neben den Bericht.

.PARAMETER NichtOeffnen
    Öffnet den fertigen Bericht nicht automatisch im Browser.

.PARAMETER PassThru
    Gibt die Befunde zusätzlich als Objekte zurück (für eigene Auswertungen).

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Dateiserver-Bericht.ps1

.EXAMPLE
    .\Dateiserver-Bericht.ps1 -Server FS01, FS02 -Freigaben Daten*, Projekte -Tiefe 4 -Csv

.NOTES
    IT Admin Tools by MM - Version 1.0 (11.2026)
    Kein Ersatz für eine Sicherheitsprüfung oder Antivirenlösung. Nutzung auf eigene Verantwortung.
#>
[CmdletBinding()]
param(
    [string[]]$Server,
    [string[]]$Freigaben,
    [ValidateRange(1, 10)][int]$Tiefe = 3,
    [ValidateRange(10, 100000)][int]$MaxOrdner = 3000,
    [ValidateRange(100, 2000000)][int]$MaxDateien = 50000,
    [switch]$OhneDateien,
    [ValidateRange(1, 50)][int]$AlterJahre = 5,
    [ValidateRange(1, 1000)][int]$GrosseDateiGB = 1,
    [ValidateRange(1, 90)][int]$FreiProzent = 15,
    [string]$Ausgabe,
    [System.Management.Automation.PSCredential]$Anmeldung,
    [switch]$Csv,
    [switch]$NichtOeffnen,
    [switch]$PassThru
)

$ErrorActionPreference = 'Stop'
$Version = '1.0'

# ------------------------------------------------------------------ Hilfsfunktionen
function Datum($d, $leer = 'nie') { if ($d) { ([datetime]$d).ToString('dd.MM.yyyy') } else { $leer } }
function JaNein($b) { if ($b) { 'Ja' } else { 'Nein' } }
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
        $sicher | Microsoft.PowerShell.Utility\Export-Csv -Path $Path -NoTypeInformation:$NoTypeInformation -Delimiter $Delimiter -Encoding $Encoding
    }
}

# Tabellen: ConvertTo-Html und Export-Csv nehmen die Spalten des ersten Objekts. Haben die Zeilen einer Liste verschiedene
# Eigenschaften (z. B. Fehlzeile neben normalen Treffern), gingen Spalten verloren: alle Namen der Reihe nach vereinigen.
function Einheitlich($liste) {
    $liste = @($liste | Where-Object { $null -ne $_ })
    if ($liste.Count -lt 2) { return $liste }
    $namen = New-Object System.Collections.Generic.List[string]
    foreach ($o in $liste) { foreach ($p in $o.PSObject.Properties.Name) { if (-not $namen.Contains($p)) { $namen.Add($p) } } }
    if ($namen.Count -eq @($liste[0].PSObject.Properties.Name).Count) { return $liste }
    @($liste | Select-Object -Property $namen)
}
# Sortierschluessel fuer Nummern wie K2, K10: erst Buchstaben, dann die Zahl als Zahl
function IdZahl($id) { $m = [regex]::Match([string]$id, '\d+'); if ($m.Success) { [int]$m.Value } else { 0 } }
function IdText($id) { [regex]::Replace([string]$id, '\d+', '') }

$Befunde = New-Object System.Collections.Generic.List[object]

# Platzhalterzeilen ("nicht geprüft") erkennt das Skript nur an den vollständigen, festen Texten, die es selbst erzeugt - keine Teilstring-Suche
# in Konto-, Objekt- oder Beschreibungstexten, die ein Angreifer beeinflussen könnte
$PlatzhalterTexte = @('Abfrage fehlgeschlagen - nicht geprüft', 'Berechtigungen nicht lesbar - nicht geprüft', 'SYSVOL nicht lesbar - nicht geprüft', 'WMI-Filter nicht lesbar - nicht geprüft', 'Kein Domänencontroller ermittelbar - SYSVOL nicht geprüft', 'Inhalt des Ordners nicht lesbar - Richtlinie nicht geprüft')
function IstPlatzhalter($o) {
    # Ein Skript kann $script:PlatzhalterSpalten setzen: dann zählen nur diese Spalten (Namensspalten sind von Benutzern beeinflussbar)
    foreach ($pr in $o.PSObject.Properties) {
        if ($script:PlatzhalterSpalten -and ($pr.Name -notin $script:PlatzhalterSpalten)) { continue }
        if (($pr.Value -is [string]) -and ($PlatzhalterTexte -ccontains $pr.Value)) { return $true }
    }
    ("$($o.Feststellung)" -clike 'Nicht geprüft - *')
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
function NichtGeprueft([string]$IdMuster, [string[]]$Namen, [string]$Verweis, [string]$Spalte = 'Server', [string]$Grund = 'Abfrage nicht möglich') {
    foreach ($x in ($Befunde | Where-Object { $_.Id -match $IdMuster })) {
        $war = $x.Anzahl
        $zeilen = @($Namen | ForEach-Object { [pscustomobject][ordered]@{ $Spalte = $_; Feststellung = "Nicht geprüft - $Grund (siehe $Verweis)"; Details = '' } })
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
    param([string]$Name, [string]$Datei, [string]$Bezug, [string]$BezugName = 'Domäne', [System.Collections.IDictionary]$Umgebung,
        [System.Collections.IDictionary]$Abschnitte = @{}, [System.Collections.IDictionary]$Anhang = @{}, [string]$Fuss)
    Write-Host '  Bericht schreiben ...'
    $jetzt = Get-Date
    $ziel = $script:Ausgabe
    if (-not $ziel) { $ziel = Join-Path (Get-Location).Path ("{0}_{1}_{2:yyyy-MM-dd_HHmm}.html" -f $Datei, $Bezug, $jetzt) }
    $ziel = [System.IO.Path]::GetFullPath($ziel)
    $ordner = Split-Path $ziel -Parent
    if (-not (Test-Path $ordner)) { New-Item -ItemType Directory -Path $ordner -Force | Out-Null }

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
    [void]$h.Append("<!DOCTYPE html><html lang='de'><head><meta charset='utf-8'><title>$(Html $Name) $(Html $Bezug)</title><style>$css</style></head><body><div class='seite'>")
    [void]$h.Append("<h1>$(Html $Name)</h1><div class='meta'>$(Html $BezugName) <b>$(Html $Bezug)</b> · erstellt am $($jetzt.ToString('dd.MM.yyyy HH:mm')) von $(Html "$env:USERDOMAIN\$env:USERNAME") auf $(Html $env:COMPUTERNAME)</div>")
    [void]$h.Append("<div class='kacheln'>")
    foreach ($s in 'Hoch', 'Mittel', 'Niedrig', 'Info', 'OK') {
        $text = @{ Hoch = 'Befunde mit hoher Priorität'; Mittel = 'mit mittlerer Priorität'; Niedrig = 'mit niedriger Priorität'; Info = 'Hinweise und nicht geprüft'; OK = 'Prüfpunkte ohne Befund' }[$s]
        [void]$h.Append("<div class='kachel $s'><b>$($zahl[$s])</b><span>$text</span></div>")
    }
    [void]$h.Append("</div><h2>Umgebung</h2><table><tr><th>Eigenschaft</th><th>Wert</th></tr>")
    foreach ($k in $Umgebung.Keys) { [void]$h.Append("<tr><td>$(Html $k)</td><td>$(Html $Umgebung[$k])</td></tr>") }
    [void]$h.Append("</table>")

    [void]$h.Append("<h2>Übersicht der Prüfpunkte</h2><table><tr><th>Nr.</th><th>Bereich</th><th>Prüfpunkt</th><th>Treffer</th><th>Bewertung</th></tr>")
    foreach ($x in $Befunde) {
        [void]$h.Append("<tr><td>$($x.Id)</td><td>$(Html $x.Bereich)</td><td><a href='#$($x.Id)'>$(Html $x.Titel)</a></td><td>$($x.Anzahl)</td><td><span class='marke m-$($x.Stufe)'>$($x.Stufe)</span></td></tr>")
    }
    [void]$h.Append("</table>")

    foreach ($k in $Abschnitte.Keys) {
        [void]$h.Append("<h2>$(Html $k)</h2>")
        $inhalt = $Abschnitte[$k]
        if ($inhalt -is [string]) { [void]$h.Append($inhalt) }
        elseif (@($inhalt).Count) { [void]$h.Append(((Einheitlich $inhalt | ConvertTo-Html -Fragment) -join "`n")) }
        else { [void]$h.Append("<p class='klein'>Keine Einträge.</p>") }
    }

    [void]$h.Append("<h2>Befunde</h2>")
    $offen = @($Befunde | Where-Object { $_.Stufe -ne 'OK' } | Sort-Object { $rang[$_.Stufe] }, { IdText $_.Id }, { IdZahl $_.Id })
    foreach ($x in $offen) {
        [void]$h.Append("<div class='befund $($x.Stufe)' id='$($x.Id)'><h3><span class='marke m-$($x.Stufe)'>$($x.Stufe)</span> $($x.Id) · $(Html $x.Titel) <span class='klein'>($($x.Anzahl) Treffer)</span></h3>")
        [void]$h.Append("<p>$(Html $x.Erklaerung)</p><p class='emp'><b>Empfehlung:</b> $(Html $x.Empfehlung)</p>")
        [void]$h.Append(((Einheitlich $x.Objekte | ConvertTo-Html -Fragment) -join "`n"))
        [void]$h.Append("</div>")
    }
    if (-not $offen.Count) { [void]$h.Append("<p>Keine Auffälligkeiten gefunden.</p>") }
    $ok = @($Befunde | Where-Object Stufe -eq 'OK')
    if ($ok.Count) {
        [void]$h.Append("<h2>Ohne Befund</h2><table><tr><th>Nr.</th><th>Prüfpunkt</th></tr>")
        foreach ($x in $ok) { [void]$h.Append("<tr id='$($x.Id)'><td>$($x.Id)</td><td>$(Html $x.Titel)</td></tr>") }
        [void]$h.Append("</table>")
    }
    foreach ($k in $Anhang.Keys) { [void]$h.Append("<h2>$(Html $k)</h2><pre>$(Html $Anhang[$k])</pre>") }
    [void]$h.Append("<div class='fuss'>$(Html $Name) Version $script:Version · IT Admin Tools by MM · Das Skript hat ausschließlich gelesen und nichts verändert. $(Html $Fuss)</div>")
    [void]$h.Append("</div></body></html>")
    [System.IO.File]::WriteAllText($ziel, $h.ToString(), (New-Object System.Text.UTF8Encoding($true)))

    $basis = [System.IO.Path]::ChangeExtension($ziel, $null).TrimEnd('.')
    if ($script:Csv) {
        foreach ($x in ($Befunde | Where-Object { $_.Anzahl -gt 0 })) {
            Einheitlich $x.Objekte | Export-Csv -Path "${basis}_$($x.Id).csv" -NoTypeInformation -Delimiter ';' -Encoding UTF8
        }
    }
    Write-Host ''
    Write-Host ("Fertig: {0} hoch, {1} mittel, {2} niedrig, {3} ohne Befund." -f $zahl.Hoch, $zahl.Mittel, $zahl.Niedrig, $zahl.OK) -ForegroundColor Green
    Write-Host "Bericht: $ziel"
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
    if ($u -band 0x10040) { return 'Löschen' }
    'Lesen'
}
$rang = @{ 'Lesen' = 0; 'Schreiben' = 1; 'Ändern' = 2; 'Vollzugriff' = 3 }

# ------------------------------------------------------------------ Server bestimmen
Write-Host 'Dateiserver-Bericht - es wird nur gelesen, nichts verändert.' -ForegroundColor Cyan
$lokaleNamen = @($env:COMPUTERNAME, 'localhost', '.', "$env:COMPUTERNAME.$env:USERDNSDOMAIN")
try { $lokaleNamen += [System.Net.Dns]::GetHostEntry('').HostName } catch { }
$lokaleNamen = @($lokaleNamen | Where-Object { $_ -and $_ -notlike '*.' } | ForEach-Object { $_.ToLower() } | Select-Object -Unique)
if (-not $Server) { $Server = @($env:COMPUTERNAME) }
$jahrText = if ($AlterJahre -eq 1) { 'Jahr' } else { 'Jahren' }      # "seit über 5 Jahren"
$jahreText = if ($AlterJahre -eq 1) { 'Jahr' } else { 'Jahre' }      # "älter als 5 Jahre"
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
        else { $e.Art = 'lokales Konto (nicht aufgelöst)' }
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
            elseif ($oc -contains 'group') { $e.Art = 'Gruppe' }
            elseif ($oc -contains 'user') { $e.Art = 'Benutzer' }
        } catch { }
        if ($e.Art -eq 'unbekannt') {
            try { if (Get-LocalUser -SID $sid -ErrorAction Stop) { $e.Art = 'Benutzer' } } catch { }
        }
        if ($e.Art -eq 'unbekannt') {
            try { if (Get-LocalGroup -SID $sid -ErrorAction Stop) { $e.Art = 'Gruppe' } } catch { }
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
    Write-Host "  Server $s ..."
    $z = [ordered]@{ Name = $s; Os = ''; Build = 0; Client = $false; OsFehler = $null; Smb = $null; Freigaben = @(); Zugriff = @{}; Pfad = @{}; Platten = @(); Schatten = @(); Volumes = @(); Fsrm = $null; Gast = $null; GastFehler = $false; Fehler = @(); Lokal = $null; MaschinenSid = $null; Verbindung = $null }
    try {
        $os = Cim $s Win32_OperatingSystem
        $z.Os = $os.Caption; $z.Build = [int]$os.BuildNumber; $z.Client = ([int]$os.ProductType -eq 1)
    } catch { $z.OsFehler = Kurz $_.Exception.Message; $z.Os = 'unbekannt' }
    $cs = $null
    try { $cs = Sitzung $s } catch { if (-not $z.OsFehler) { $z.OsFehler = Kurz $_.Exception.Message } }
    # Ohne Verbindung zu einem entfernten Server nie lokal weiterfragen (das würde die Daten des falschen Rechners unter dem Servernamen zeigen)
    if (-not (IstLokal $s) -and -not $cs) {
        $nichtErreicht += [pscustomobject][ordered]@{ Server = $s; Feststellung = 'Abfrage fehlgeschlagen - nicht geprüft'; Details = $(if ($z.OsFehler) { $z.OsFehler } else { 'Keine Verbindung zum Server' }) }
        continue
    }
    $cp = @{}; if ($cs) { $cp.CimSession = $cs }
    try { $z.Smb = Get-SmbServerConfiguration @cp -ErrorAction Stop } catch { $z.Fehler += "SMB-Einstellungen: $(Kurz $_.Exception.Message)" }
    $freigabenGelesen = $true
    try {
        $alle = @(Get-SmbShare @cp -ErrorAction Stop)
        $z.Freigaben = @($alle | Where-Object { -not $_.Special -and $_.Name -ne 'IPC$' })
        if ($Freigaben) { $z.Freigaben = @($z.Freigaben | Where-Object { $n = $_.Name; @($Freigaben | Where-Object { $n -like $_ }).Count }) }
    } catch { $freigabenGelesen = $false; $z.Fehler += "Freigaben: $(Kurz $_.Exception.Message)" }
    if ($z.OsFehler -and -not $freigabenGelesen -and -not $z.Smb) {
        $nichtErreicht += [pscustomobject][ordered]@{ Server = $s; Feststellung = 'Abfrage fehlgeschlagen - nicht geprüft'; Details = $z.OsFehler }
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
            foreach ($g in (Cim $s Win32_Group 'LocalAccount=True')) { $lk["$($g.SID)"] = [pscustomobject]@{ Name = "$($g.Domain)\$($g.Name)"; Art = 'Gruppe'; Verwaist = $false } }
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
                } catch { $z.Fehler += "Dateizugriff mit -Anmeldung auf $l`$: $(Kurz $_.Exception.Message)" }
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
            $f1 += [pscustomobject][ordered]@{ Server = $z.Name; Freigabe = $f.Name; Konto = ''; Recht = ''; Feststellung = 'Berechtigungen nicht lesbar - nicht geprüft'; Details = 'Freigaberechte sind ohne Administratorrechte nicht (vollständig) lesbar' }
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
                    if ($recht -ne 'Lesen' -and -not $standard) { $f1 += [pscustomobject][ordered]@{ Server = $z.Name; Freigabe = $f.Name; Konto = $k; Recht = $recht; Feststellung = 'Abfrage fehlgeschlagen - nicht geprüft'; Details = 'Konto nicht auflösbar' } }
                    continue
                }
                $aufgehoben = $verw.ContainsKey("$k") -and ($verw["$k"] -ge $rang[$recht])
                if ($sid -and ($sid -in $weit -or (IstDomaenenBenutzer $sid)) -and $recht -ne 'Lesen' -and -not $aufgehoben -and -not $standard) {
                    $f1 += [pscustomobject][ordered]@{ Server = $z.Name; Freigabe = $f.Name; Konto = $k; Recht = $recht; Feststellung = 'Freigabe erlaubt Schreiben für einen sehr großen Personenkreis' }
                }
                if ($sid -in $weit -or (IstDomaenenBenutzer $sid)) { $weitText += "$k ($recht); " }
            }
        }
        $pfad = SharePfad $z $f
        if (-not $pfad) {
            $ohnePfad += [pscustomobject][ordered]@{ Server = $z.Name; Freigabe = $f.Name }
            $ueberspringen["$($z.Name)|$($f.Name)"] = 'Pfad'
            $f5 += [pscustomobject][ordered]@{ Server = $z.Name; Freigabe = $f.Name; Pfad = ''; Feststellung = 'Berechtigungen nicht lesbar - nicht geprüft'; Details = 'Der Ordner der Freigabe ist nicht ermittelbar (Administratorrechte oder WMI-Zugriff fehlen)' }
        } else {
            if ($pfad -match '^[A-Za-z]:\\?$' -or $pfad -match '^[A-Za-z]:\\(Windows|Windows\\System32|Program Files|Program Files \(x86\)|ProgramData|Users)\\?$') {
                $f2 += [pscustomobject][ordered]@{ Server = $z.Name; Freigabe = $f.Name; Pfad = $pfad; Feststellung = 'Freigabe zeigt auf ein ganzes Laufwerk oder einen Systemordner' }
            }
            $wz = LeseWurzel $z $f
            $da = $true
            try { $da = [bool](Test-Path -LiteralPath $wz -ErrorAction Stop) } catch { $da = $null }
            if ($null -eq $da) {
                $f5 += [pscustomobject][ordered]@{ Server = $z.Name; Freigabe = $f.Name; Pfad = $pfad; Feststellung = 'Abfrage fehlgeschlagen - nicht geprüft'; Details = 'Zugriff auf den Ordner nicht möglich' }
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
            'Zugriff für große Gruppen' = $(if (-not $zugLesbar) { 'nicht lesbar' } elseif ($weitText) { $weitText.TrimEnd(' ', ';') } else { '-' })
            Verschlüsselt = JaNein $f.EncryptData; 'Zugriffsbasierte Aufzählung' = JaNein ("$($f.FolderEnumerationMode)" -eq 'AccessBased')
        }
    }
}
Befund 'F1' 'Freigaben' 'Freigaben mit Schreibrecht für Jeder, Authentifizierte Benutzer oder Benutzer' 'Hoch' $f1 `
    'Auf Freigabe-Ebene dürfen sehr viele Personen schreiben. Wirksam ist immer das strengere von Freigabe- und Dateisystemrechten; der Freigabeeintrag "Jeder" ist üblich, solange die Dateisystemrechte streng sind (siehe N1). Bleibt beides offen, kann jeder Mitarbeiter und im schlimmsten Fall Schadsoftware alle Dateien verändern oder verschlüsseln. Die Standardfreigaben SYSVOL und NETLOGON von Domänencontrollern werden nicht bewertet. Ein gleichzeitiges Verweigern für dasselbe Konto wird gegengerechnet.' `
    'Freigaberechte auf die zuständigen Gruppen beschränken, die eigentliche Steuerung über Dateisystemrechte mit Gruppen. "Jeder" und "Authentifizierte Benutzer" nicht mit Schreibrecht eintragen.'
Befund 'F3' 'Freigaben' 'Versteckte Freigaben (Name endet auf $) außer den Standardfreigaben' 'Info' $f3 `
    'Versteckte Freigaben erscheinen nicht in der Netzwerkumgebung, sind aber per Name erreichbar. Sie dienen oft Sicherung oder Software und geraten leicht in Vergessenheit.' `
    'Zweck jeder Freigabe klären und dokumentieren; nicht mehr benötigte Freigaben entfernen.'

Befund 'N2' 'Dateisystemrechte' 'Berechtigungen für nicht mehr vorhandene Konten (verwaiste SIDs)' 'Mittel' $n2 `
    'Die Einträge gehören zu gelöschten Konten. Sie bewirken nichts, sind aber Unordnung und ein Risiko, falls eine SID später wiederverwendet wird (z. B. nach einer Domänenmigration).' `
    'Verwaiste Einträge aufräumen (Sicherheitseinstellungen des Ordners) und Konten künftig erst nach dem Entfernen der Rechte löschen.'

# ------------------------------------------------------------------ Server-Einstellungen und Volumes
$s1 = @(); $s2 = @(); $s3 = @(); $s4 = @(); $s5 = @(); $s6 = @(); $s7 = @(); $s8 = @()
$nichtBewertet = @()
foreach ($z in $srvListe) {
    if ($z.Smb) {
        if ($z.Smb.EnableSMB1Protocol) { $s1 += [pscustomobject][ordered]@{ Server = $z.Name; Feststellung = 'SMB1 ist aktiv'; Details = '' } }
        if (-not $z.Smb.RequireSecuritySignature) { $s2 += [pscustomobject][ordered]@{ Server = $z.Name; Feststellung = 'SMB-Signierung wird nicht verlangt'; Details = "Signierung angeboten: $(JaNein $z.Smb.EnableSecuritySignature)" } }
        $ns = @($z.Smb.NullSessionShares | Where-Object { $_ })
        if ($ns.Count) { $s3 += [pscustomobject][ordered]@{ Server = $z.Name; Feststellung = 'Freigaben ohne Anmeldung erreichbar (Null-Sitzung)'; Details = ($ns -join ', ') } }
    } else {
        $grund = (($z.Fehler | Where-Object { $_ -like 'SMB*' }) -join ' ')
        $pl = [pscustomobject][ordered]@{ Server = $z.Name; Feststellung = 'Abfrage fehlgeschlagen - nicht geprüft'; Details = $grund }
        $s1 += $pl; $s2 += $pl; $s3 += $pl
    }
    if ($z.GastFehler) { $s4 += [pscustomobject][ordered]@{ Server = $z.Name; Feststellung = 'Abfrage fehlgeschlagen - nicht geprüft'; Details = 'Lokale Benutzer des Servers nicht abfragbar (Administratorrechte oder WMI-Zugriff fehlen)' } }
    elseif ($z.Gast -and -not $z.Gast.Disabled) { $s4 += [pscustomobject][ordered]@{ Server = $z.Name; Feststellung = 'Lokales Gastkonto ist aktiviert'; Details = $z.Gast.Name } }
    # Volumes mit Freigaben
    $laufwerke = @($z.Pfad.Values | ForEach-Object { if ("$_" -match '^([A-Za-z]):') { $Matches[1].ToUpper() + ':' } } | Select-Object -Unique)
    if ($null -eq $z.Platten -and ($laufwerke.Count -or $z.Freigaben.Count)) { $s5 += [pscustomobject][ordered]@{ Server = $z.Name; Volume = ''; 'Frei (GB)' = ''; 'Gesamt (GB)' = ''; 'Frei (%)' = ''; Feststellung = 'Abfrage fehlgeschlagen - nicht geprüft' } }
    if ($null -eq $z.Schatten -and -not $laufwerke.Count -and $z.Freigaben.Count) { $s6 += [pscustomobject][ordered]@{ Server = $z.Name; Volume = ''; Feststellung = 'Abfrage fehlgeschlagen - nicht geprüft' } }
    foreach ($lw in $laufwerke) {
        $pl = @($z.Platten) | Where-Object { $_.DeviceID -eq $lw } | Select-Object -First 1
        if ($pl -and $pl.Size) {
            $proz = [math]::Round(100 * $pl.FreeSpace / $pl.Size, 1)
            if ($proz -lt $FreiProzent) { $s5 += [pscustomobject][ordered]@{ Server = $z.Name; Volume = $lw; 'Frei (GB)' = [math]::Round($pl.FreeSpace / 1GB, 1); 'Gesamt (GB)' = [math]::Round($pl.Size / 1GB, 1); 'Frei (%)' = $proz } }
        }
        if ($null -eq $z.Schatten) {
            $s6 += [pscustomobject][ordered]@{ Server = $z.Name; Volume = $lw; Feststellung = 'Abfrage fehlgeschlagen - nicht geprüft' }
        } else {
            $vol = @($z.Volumes | Where-Object { $_.DriveLetter -eq $lw } | Select-Object -First 1)
            $anz = 0
            if ($vol.Count -and $vol[0].DeviceID) { $anz = @($z.Schatten | Where-Object { $_.VolumeName -eq $vol[0].DeviceID }).Count }
            if (-not $anz -and $lw -ne 'C:') { $s6 += [pscustomobject][ordered]@{ Server = $z.Name; Volume = $lw; Feststellung = 'Keine Schattenkopien (Vorgängerversionen)' } }
        }
    }
    if ($null -eq $z.Fsrm) { $s7 += [pscustomobject][ordered]@{ Server = $z.Name; Feststellung = 'Abfrage fehlgeschlagen - nicht geprüft'; Details = '' } }
    elseif ($z.Fsrm -eq $false -and $z.Freigaben.Count) { $s7 += [pscustomobject][ordered]@{ Server = $z.Name; Feststellung = 'Kein Ressourcen-Manager für Dateiserver (Kontingente und Dateiprüfung) installiert'; Details = '' } }
    # Betriebssystem
    if ($z.OsFehler) {
        $s8 += [pscustomobject][ordered]@{ Server = $z.Name; Betriebssystem = ''; 'Support-Ende' = ''; Feststellung = 'Abfrage fehlgeschlagen - nicht geprüft'; Details = $z.OsFehler }
    } elseif ($z.Client) {
        $nichtBewertet += "$($z.Name) ($($z.Os))"
    } else {
        $tab = $ServerTabelle | Where-Object { $z.Build -ge $_[0] } | Select-Object -First 1
        if ($tab) {
            $ende = [datetime]$tab[2]
            if ($ende -lt (Get-Date)) { $s8 += [pscustomobject][ordered]@{ Server = $z.Name; Betriebssystem = $tab[1]; 'Support-Ende' = Datum $ende; Feststellung = 'Support abgelaufen'; Details = '' } }
            elseif ($ende -lt (Get-Date).AddDays(365)) { $s8 += [pscustomobject][ordered]@{ Server = $z.Name; Betriebssystem = $tab[1]; 'Support-Ende' = Datum $ende; Feststellung = 'Support endet innerhalb eines Jahres'; Details = '' } }
        }
    }
}
Befund 'S1' 'Server' 'SMB1 aktiv' 'Hoch' $s1 `
    'SMB1 ist ein veraltetes Protokoll mit schweren Sicherheitslücken (u. a. von der Ransomware WannaCry ausgenutzt). Heutige Windows-Versionen brauchen es nicht mehr.' `
    'SMB1 abschalten, vorher mit der Überwachung (Audit) prüfen, ob noch alte Geräte (Scanner, NAS) es verwenden.'
Befund 'S2' 'Server' 'SMB-Signierung nicht erforderlich' 'Mittel' $s2 `
    'Ohne verlangte Signierung können Angreifer im Netzwerk Datenverkehr manipulieren oder umleiten (Relay-Angriffe).' `
    'SMB-Signierung verlangen (Gruppenrichtlinie "Microsoft-Netzwerk (Server): Kommunikation digital signieren (immer)"), vorher Kompatibilität alter Geräte prüfen.'

Befund 'D3' 'Datenbestand' 'Ausführbare Dateien und Skripte in Datenfreigaben' 'Niedrig' $d3 `
    'Programme und Skripte auf Datenfreigaben (.exe, .bat, .ps1, .vbs, .msi ...) sind oft Installationsreste. Schreibbare Freigaben mit ausführbaren Dateien sind ein typischer Weg, Schadsoftware zu verteilen.' `
    'Zweck prüfen. Installationsmedien auf eine schreibgeschützte Software-Freigabe legen, Reste löschen.'

# Server, die nicht abgefragt werden konnten, und gekürzte Listen dürfen nicht als "ohne Befund" erscheinen
if ($nichtErreicht.Count) {
    $namen = @($nichtErreicht | ForEach-Object { $_.Server })
    NichtGeprueft '^[FNSD]\d+$' $namen 'Abschnitt "Nicht abgefragte Server"' 'Server' 'Server nicht abfragbar'
}
foreach ($x in ($Befunde | Where-Object { $_.Id -match '^N\d$' })) {
    foreach ($g in $gekuerztOrdner) {
        $x.Objekte = @($x.Objekte) + [pscustomobject][ordered]@{ Server = ''; Freigabe = $g; Ordner = ''; Konto = ''; Recht = ''; Feststellung = 'Nur die ersten Ordner geprüft - Grenze erreicht (-MaxOrdner erhöhen)'; Details = '' }
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
    'Freigaben ausgewertet'     = @($uebersicht).Count
    'Ordner auf Rechte geprüft' = "$ordnerGeprueft (Tiefe $Tiefe, höchstens $MaxOrdner je Freigabe)"
    'Dateien'                   = $(if ($OhneDateien) { 'nicht ausgewertet (-OhneDateien)' } else { "höchstens $MaxDateien je Freigabe, alt = älter als $AlterJahre $jahreText" })
    'Ausgeführt als'            = "$env:USERDOMAIN\$env:USERNAME"
}
if ($Anmeldung -and $laufwerkeVerbunden.Count) { $umgebung['Dateizugriff'] = "bei entfernten Servern als $($Anmeldung.UserName)" }
if ($verknuepfungen.Count) { $umgebung['Übersprungene Verknüpfungen (Junctions/Symlinks)'] = (($verknuepfungen.Keys | Sort-Object | ForEach-Object { "$($_ -replace '^.*\|', ''): $($verknuepfungen[$_])" }) -join '; ') }
if ($mehrfach.Count) { $umgebung['Mehrfach freigegebene Ordner (nur einmal gezählt)'] = $mehrfach -join '; ' }
if ($ohnePfad.Count) { $umgebung['Freigabepfade nicht sichtbar'] = "$($ohnePfad.Count) Freigaben (ohne Administratorrechte oder WMI-Zugriff liefert Windows keine Pfade)" }
if ($nichtBewertet.Count) { $umgebung['Kein Serverbetriebssystem (Support nicht bewertet)'] = $nichtBewertet -join '; ' }
if ($nichtErreicht.Count) { $umgebung['Nicht erreicht'] = ($nichtErreicht | ForEach-Object { $_.Server }) -join ', ' }
$abschn = [ordered]@{ 'Freigaben im Überblick' = @($uebersicht) }
if ($nichtErreicht.Count) { $abschn['Nicht abgefragte Server'] = @($nichtErreicht) }
if ($gesamt.Count) { $abschn['Datenbestand je Freigabe'] = $gesamt }
$bezug = ($srvListe | ForEach-Object { $_.Name }) -join '_'
if (-not $bezug) { $bezug = ($Server -join '_') }
$basis = BerichtSchreiben -Name 'Dateiserver-Bericht Light' -Datei 'Dateiserver-Bericht' -Bezug $bezug -BezugName 'Server' -Umgebung $umgebung -Abschnitte $abschn `
    -Fuss 'Light-Version mit 6 von 24 Prüfpunkten der Vollversion (https://mueller86.gumroad.com/l/wbbwj). Grundlage sind die Freigaben, Berechtigungen und Dateieigenschaften zum Zeitpunkt der Abfrage; Dateiinhalte wurden nicht gelesen. Der Bericht ersetzt keine Sicherheitsprüfung und keine Datensicherung.'
if ($Csv) {
    if ($uebersicht.Count) { Einheitlich $uebersicht | Export-Csv -Path "${basis}_Freigaben.csv" -NoTypeInformation -Delimiter ';' -Encoding UTF8 }
    if ($gesamt.Count) { Einheitlich $gesamt | Export-Csv -Path "${basis}_Datenbestand.csv" -NoTypeInformation -Delimiter ';' -Encoding UTF8 }
}
foreach ($cs in $sitzungen.Values) { try { Remove-CimSession -CimSession $cs } catch { } }
foreach ($l in $laufwerkeVerbunden) { try { Remove-PSDrive -Name $l -Force -ErrorAction SilentlyContinue } catch { } }
if ($PassThru) { $Befunde }
