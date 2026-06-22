#Requires -Version 7.0

<#
.SYNOPSIS
    Schreibt das Zuteilungs-Array (Get-SPSiteBusinessUnit.ps1 -> $r.Zuteilung)
    in eine SQL-Tabelle. Nutzt die Invoke-SQL-Funktion (muss vorher geladen sein).

.DESCRIPTION
    - legt die Zieltabelle an, falls sie noch nicht existiert
    - escaped alle Textwerte (Apostrophe verdoppeln) -> kein Bruch bei Titeln
      wie "O'Brien Team"
    - schreibt Zahlen kulturunabhaengig (Punkt als Dezimaltrenner) -> kein
      Fehler auf Servern mit deutscher/schweizer Locale
    - haengt standardmaessig an (Spalte ImportDatum); -Leeren leert vorher

.PARAMETER Data
    Das Array, typischerweise $r.Zuteilung. Objekte mit den Eigenschaften
    SiteName, SiteUrl, BusinessUnit, Begruendung, SpeicherMB, SpeicherGB.

.PARAMETER Env
    'Prod' oder 'Test' - wird an Invoke-SQL durchgereicht.

.PARAMETER Table
    Zieltabelle, Default 'dbo.SiteZuteilung'.

.PARAMETER Leeren
    Tabelle vor dem Import leeren (aktueller Snapshot statt Historie).

.EXAMPLE
    . .\Invoke-SQL.ps1            # Funktion des Chefs laden
    . .\Import-SPSiteToSql.ps1    # diese Funktion laden
    $r = .\Get-SPSiteBusinessUnit.ps1 -GraphOnly -ClientId <id> -Tenant <t> -CertificateThumbprint <tp>
    Import-SPSiteToSql -Data $r.Zuteilung -Env Test
#>
function Import-SPSiteToSql {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Data,
        [Parameter(Mandatory)][ValidateSet('Prod', 'Test')][string]$Env,
        [string]$Table = 'dbo.SiteZuteilung',
        [switch]$Leeren
    )

    if (-not (Get-Command Invoke-SQL -ErrorAction SilentlyContinue)) {
        throw "Invoke-SQL ist nicht geladen. Zuerst die Datei mit der Invoke-SQL-Funktion dot-sourcen, z.B.: . .\Invoke-SQL.ps1"
    }

    $rows = @($Data)
    if ($rows.Count -eq 0) { Write-Warning 'Keine Daten zum Importieren.'; return 0 }

    # Text SQL-sicher machen: Apostrophe verdoppeln, NULL, Unicode (N'...')
    function Esc($v) {
        if ($null -eq $v -or "$v" -eq '') { return 'NULL' }
        return "N'" + ([string]$v).Replace("'", "''") + "'"
    }
    # Zahl kulturunabhaengig (immer Punkt als Dezimaltrenner)
    function Num($v) {
        if ($null -eq $v -or "$v" -eq '') { return 'NULL' }
        return [string]::Format([Globalization.CultureInfo]::InvariantCulture, '{0}', $v)
    }

    # 1) Tabelle anlegen, falls nicht vorhanden
    $createSql = @"
IF OBJECT_ID('$Table','U') IS NULL
CREATE TABLE $Table (
    SiteName     NVARCHAR(400),
    SiteUrl      NVARCHAR(1000),
    BusinessUnit NVARCHAR(200),
    Begruendung  NVARCHAR(MAX),
    SpeicherMB   BIGINT,
    SpeicherGB   DECIMAL(18,2),
    ImportDatum  DATETIME2 NOT NULL DEFAULT SYSUTCDATETIME()
);
"@
    Invoke-SQL -env $Env -command $createSql -WriteWithRowsAff | Out-Null

    # 2) Optional leeren
    if ($Leeren) {
        Invoke-SQL -env $Env -command "DELETE FROM $Table;" -WriteWithRowsAff | Out-Null
    }

    # 3) Zeilen einfuegen
    $imported = 0
    foreach ($s in $rows) {
        $insert = "INSERT INTO $Table (SiteName, SiteUrl, BusinessUnit, Begruendung, SpeicherMB, SpeicherGB) VALUES ($(Esc $s.SiteName), $(Esc $s.SiteUrl), $(Esc $s.BusinessUnit), $(Esc $s.Begruendung), $(Num $s.SpeicherMB), $(Num $s.SpeicherGB));"
        $aff = Invoke-SQL -env $Env -command $insert -WriteWithRowsAff
        if ($aff) { $imported += [int]$aff }
    }

    Write-Host ("{0} von {1} Zeilen in {2} ({3}) importiert." -f $imported, $rows.Count, $Table, $Env) -ForegroundColor Green
    return $imported
}
