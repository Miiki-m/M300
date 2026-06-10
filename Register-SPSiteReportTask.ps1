<#
.SYNOPSIS
    Registriert eine geplante Aufgabe in der Windows-Aufgabenplanung, die den
    SharePoint-Site-Business-Unit-Report regelmaessig unbeaufsichtigt ausfuehrt
    (via Run-SPSiteReport.ps1, App-Only mit Zertifikat).

.DESCRIPTION
    Standard: woechentlich montags 06:00 Uhr als SYSTEM-Konto.
    Mit -ServiceAccount laeuft die Aufgabe stattdessen unter einem Dienstkonto
    (Passwort wird bei der Registrierung abgefragt; das Konto braucht das Recht
    "Anmelden als Stapelverarbeitungsauftrag").

    Voraussetzungen fuer das ausfuehrende Konto (SYSTEM oder Dienstkonto):
      - Module systemweit installiert (als Admin in pwsh):
            Install-Module Microsoft.Graph.Authentication -Scope AllUsers
            Install-Module PnP.PowerShell -Scope AllUsers     # entfaellt bei -GraphOnly
      - Zertifikat importiert in Cert:\LocalMachine\My (das Graph-SDK und
        aktuelle PnP-Versionen durchsuchen CurrentUser\My UND LocalMachine\My).
        Bei Dienstkonto: in certlm.msc ueber "Private Schluessel verwalten"
        dem Konto Lesezugriff auf den privaten Schluessel geben.

.EXAMPLE
    # Woechentlich montags 06:00, Least Privilege (GraphOnly), als SYSTEM
    .\Register-SPSiteReportTask.ps1 -GraphOnly -ClientId <AppId> -Tenant contoso.onmicrosoft.com -CertificateThumbprint <Thumbprint>

.EXAMPLE
    # Taeglich 05:30 unter einem Dienstkonto, voller Funktionsumfang
    .\Register-SPSiteReportTask.ps1 -ClientId <AppId> -Tenant contoso.onmicrosoft.com -CertificateThumbprint <Thumbprint> -TenantAdminUrl https://contoso-admin.sharepoint.com -Frequency Daily -Time 05:30 -ServiceAccount 'CONTOSO\svc-spreport'
#>
[CmdletBinding()]
param(
    # =========================================================================
    # >>> KONFIGURATION: DIESE WERTE AUSFUELLEN <<<
    # Alle Werte koennen alternativ auch als Parameter uebergeben werden.
    # =========================================================================

    # App-ID (Client-ID) der Entra-App-Registrierung:
    [string]$ClientId = '',

    # Tenant, z.B. 'contoso.onmicrosoft.com':
    [string]$Tenant = '',

    # Thumbprint des App-Zertifikats (muss in Cert:\LocalMachine\My liegen):
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

    # Name der geplanten Aufgabe:
    [string]$TaskName = 'SPSite-BusinessUnit-Report',

    # Zeitplan: 'Weekly' (mit $DaysOfWeek) oder 'Daily':
    [ValidateSet('Daily', 'Weekly')]
    [string]$Frequency = 'Weekly',

    # Wochentag(e) bei Frequency = 'Weekly':
    [System.DayOfWeek[]]$DaysOfWeek = @([System.DayOfWeek]::Monday),

    # Startzeit (24h-Format):
    [string]$Time = '06:00',

    # Dienstkonto 'DOMAIN\benutzer' (Passwort wird abgefragt); leer = SYSTEM:
    [string]$ServiceAccount = '',

    # Pfad zu PowerShell 7:
    [string]$PwshPath = "$env:ProgramFiles\PowerShell\7\pwsh.exe"
)

$ErrorActionPreference = 'Stop'

if (-not $ClientId)              { throw 'ClientId fehlt - oben im Skript ausfuellen oder als Parameter uebergeben.' }
if (-not $Tenant)                { throw 'Tenant fehlt - oben im Skript ausfuellen oder als Parameter uebergeben.' }
if (-not $CertificateThumbprint) { throw 'CertificateThumbprint fehlt - oben im Skript ausfuellen oder als Parameter uebergeben.' }
if (-not (Test-Path -Path $PwshPath)) {
    throw ("pwsh.exe nicht gefunden unter '{0}'. PowerShell 7 installieren oder -PwshPath angeben." -f $PwshPath)
}
if (-not $GraphOnly -and -not $TenantAdminUrl) {
    throw 'Ohne -GraphOnly muss -TenantAdminUrl angegeben werden.'
}

$runnerPath = Join-Path -Path $PSScriptRoot -ChildPath 'Run-SPSiteReport.ps1'
if (-not (Test-Path -Path $runnerPath)) {
    throw ("Run-SPSiteReport.ps1 nicht gefunden in {0} - beide Skripte in denselben Ordner legen." -f $PSScriptRoot)
}

# --- Argumentliste fuer pwsh zusammenbauen ---
$taskArgs = @(
    '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
    '-File', ('"{0}"' -f $runnerPath),
    '-ClientId', ('"{0}"' -f $ClientId),
    '-Tenant', ('"{0}"' -f $Tenant),
    '-CertificateThumbprint', ('"{0}"' -f $CertificateThumbprint),
    '-OutputFolder', ('"{0}"' -f $OutputFolder)
)
if ($TenantAdminUrl)  { $taskArgs += @('-TenantAdminUrl', ('"{0}"' -f $TenantAdminUrl)) }
if ($GraphOnly)       { $taskArgs += '-GraphOnly' }
if ($DeepScan)        { $taskArgs += '-DeepScan' }
if ($IncludeOneDrive) { $taskArgs += '-IncludeOneDrive' }

$action = New-ScheduledTaskAction -Execute $PwshPath -Argument ($taskArgs -join ' ')

if ($Frequency -eq 'Daily') {
    $trigger = New-ScheduledTaskTrigger -Daily -At $Time
}
else {
    $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $DaysOfWeek -At $Time
}

$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Hours 6)

if ($ServiceAccount) {
    $credential = Get-Credential -UserName $ServiceAccount -Message 'Passwort des Dienstkontos fuer die geplante Aufgabe'
    Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings `
        -User $credential.UserName -Password $credential.GetNetworkCredential().Password `
        -RunLevel Highest -Force | Out-Null
}
else {
    $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
    Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings `
        -Principal $principal -Force | Out-Null
}

Write-Host ''
Write-Host ("Aufgabe '{0}' registriert." -f $TaskName) -ForegroundColor Green
if ($Frequency -eq 'Daily') {
    Write-Host ("Zeitplan: taeglich um {0}" -f $Time)
}
else {
    Write-Host ("Zeitplan: woechentlich ({0}) um {1}" -f ($DaysOfWeek -join ', '), $Time)
}
Write-Host ("Konto:    {0}" -f $(if ($ServiceAccount) { $ServiceAccount } else { 'SYSTEM' }))
Write-Host ("Ablage:   {0}" -f $OutputFolder)
Write-Host ''
Write-Host 'Testlauf jetzt starten:' -ForegroundColor Cyan
Write-Host ("    Start-ScheduledTask -TaskName '{0}'" -f $TaskName)
Write-Host 'Letztes Ergebnis pruefen (0 = OK):' -ForegroundColor Cyan
Write-Host ("    (Get-ScheduledTaskInfo -TaskName '{0}').LastTaskResult" -f $TaskName)
