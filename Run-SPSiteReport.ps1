#Requires -Version 7.0

<#
.SYNOPSIS
    Wrapper fuer die geplante, unbeaufsichtigte Ausfuehrung von
    Get-SPSiteBusinessUnit.ps1 (Windows-Aufgabenplanung).

.DESCRIPTION
    - erzwingt App-Only-Authentifizierung per Zertifikat (kein Benutzer-Login)
    - schreibt die CSV mit Zeitstempel in einen Report-Ordner
    - protokolliert den kompletten Lauf als Transcript-Log
    - raeumt Reports/Logs aelter als -RetentionDays auf
    - Exit-Code 0 = Erfolg, 1 = Fehler (fuer die Ueberwachung der Aufgabe)

.EXAMPLE
    .\Run-SPSiteReport.ps1 -GraphOnly -ClientId <AppId> -Tenant contoso.onmicrosoft.com -CertificateThumbprint <Thumbprint>
#>
[CmdletBinding()]
param(
    # =========================================================================
    # >>> KONFIGURATION: DIESE WERTE AUSFUELLEN <<<
    # Alle Werte koennen alternativ als Parameter uebergeben werden (so macht
    # es die von Register-SPSiteReportTask.ps1 erstellte Aufgabe).
    # =========================================================================

    # App-ID (Client-ID) der Entra-App-Registrierung:
    [string]$ClientId = '',

    # Tenant, z.B. 'contoso.onmicrosoft.com':
    [string]$Tenant = '',

    # Thumbprint des App-Zertifikats (App-Only ist fuer geplante Laeufe Pflicht):
    [string]$CertificateThumbprint = '',

    # SharePoint Admin Center URL - nur noetig, wenn GraphOnly = $false:
    [string]$TenantAdminUrl = '',

    # $true = Least-Privilege-Modus (nur Graph, keine SharePoint-Berechtigung):
    [switch]$GraphOnly = $false,

    # $true = zusaetzlich Site Collection Admins + SP-Besitzergruppen (langsamer):
    [switch]$DeepScan = $false,

    # $true = persoenliche OneDrive-Sites mit auswerten:
    [switch]$IncludeOneDrive = $false,

    # Ablageordner fuer CSV-Reports und Transcript-Logs:
    [string]$OutputFolder = 'C:\Reports\SPSiteReport',

    # Pfad zum Hauptskript; leer = gleiche Ablage wie dieser Runner:
    [string]$ScriptPath = '',

    # Reports/Logs aelter als n Tage werden geloescht:
    [int]$RetentionDays = 90
)

$ErrorActionPreference = 'Stop'

if (-not $ScriptPath) {
    $ScriptPath = Join-Path -Path $PSScriptRoot -ChildPath 'Get-SPSiteBusinessUnit.ps1'
}

$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$csvPath   = Join-Path -Path $OutputFolder -ChildPath ('SPSite-BU-Report_{0}.csv' -f $timestamp)
$logPath   = Join-Path -Path $OutputFolder -ChildPath ('Transcript_{0}.log' -f $timestamp)

New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
Start-Transcript -Path $logPath | Out-Null

try {
    if (-not $ClientId)              { throw 'ClientId fehlt - oben im Skript ausfuellen oder als Parameter uebergeben.' }
    if (-not $Tenant)                { throw 'Tenant fehlt - oben im Skript ausfuellen oder als Parameter uebergeben.' }
    if (-not $CertificateThumbprint) { throw 'CertificateThumbprint fehlt - oben im Skript ausfuellen oder als Parameter uebergeben.' }
    if (-not (Test-Path -Path $ScriptPath)) {
        throw ("Hauptskript nicht gefunden: {0}" -f $ScriptPath)
    }
    if (-not $GraphOnly -and -not $TenantAdminUrl) {
        throw 'Ohne GraphOnly muss TenantAdminUrl angegeben werden.'
    }

    $reportParams = @{
        ClientId              = $ClientId
        Tenant                = $Tenant
        CertificateThumbprint = $CertificateThumbprint
        OutputCsv             = $csvPath
        GraphOnly             = $GraphOnly
        DeepScan              = $DeepScan
        IncludeOneDrive       = $IncludeOneDrive
    }
    if ($TenantAdminUrl) { $reportParams['TenantAdminUrl'] = $TenantAdminUrl }

    & $ScriptPath @reportParams

    # Alte Reports/Logs aufraeumen (nur eigene Dateimuster anfassen)
    $cutoff = (Get-Date).AddDays(-$RetentionDays)
    Get-ChildItem -Path $OutputFolder -File |
        Where-Object {
            ($_.Name -like 'SPSite-BU-Report_*.csv' -or $_.Name -like 'Transcript_*.log') -and
            $_.LastWriteTime -lt $cutoff
        } |
        Remove-Item -Force

    Write-Host ("Lauf erfolgreich. Report: {0}" -f $csvPath) -ForegroundColor Green
    exit 0
}
catch {
    Write-Warning ("Report-Lauf fehlgeschlagen: {0}" -f $_.Exception.Message)
    exit 1
}
finally {
    Stop-Transcript | Out-Null
}
