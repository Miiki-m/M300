<#
.SYNOPSIS
    Diagnose fuer die App-Only-Authentifizierung des SP-Site-Reports.
    Prueft Schritt fuer Schritt: Zertifikat -> Graph-Permissions ->
    SharePoint-Permission, und zeigt die Rollen im ausgestellten Token an.

.DESCRIPTION
    Typische Befunde:
      - Graph-Tests scheitern  -> Graph-APPLICATION-Permissions fehlen oder
                                  Admin Consent nicht erteilt
      - SharePoint-Test 401    -> Sites.FullControl.All wurde unter der
                                  falschen API hinzugefuegt (Microsoft Graph
                                  statt "SharePoint") oder Consent fehlt
      - 'roles' im Token leer  -> Admin Consent fehlt oder alte Session
                                  (neue PowerShell-Session starten!)
#>
[CmdletBinding()]
param(
    # =========================================================================
    # >>> KONFIGURATION: DIESE WERTE AUSFUELLEN <<<
    # =========================================================================

    # App-ID (Client-ID) der Entra-App-Registrierung:
    [string]$ClientId = '',

    # Tenant, z.B. 'contoso.onmicrosoft.com':
    [string]$Tenant = '',

    # Thumbprint des App-Zertifikats:
    [string]$CertificateThumbprint = '',

    # SharePoint Admin Center URL, z.B. 'https://contoso-admin.sharepoint.com'.
    # Leer lassen, wenn nur GraphOnly genutzt wird (SharePoint-Test wird dann
    # uebersprungen):
    [string]$TenantAdminUrl = ''
)

$ErrorActionPreference = 'Stop'
$script:Problems = [System.Collections.Generic.List[string]]::new()

function Write-Ok   { param([string]$Text) Write-Host ("  [OK]     {0}" -f $Text) -ForegroundColor Green }
function Write-Bad  { param([string]$Text) Write-Host ("  [FEHLER] {0}" -f $Text) -ForegroundColor Red; $script:Problems.Add($Text) }
function Write-Info { param([string]$Text) Write-Host ("  [INFO]   {0}" -f $Text) -ForegroundColor Yellow }

function ConvertFrom-JwtPayload {
    <# Dekodiert den Payload eines JWT (ohne Signaturpruefung, nur Anzeige). #>
    param([string]$Token)
    $payload = $Token.Split('.')[1].Replace('-', '+').Replace('_', '/')
    switch ($payload.Length % 4) {
        2 { $payload += '==' }
        3 { $payload += '=' }
    }
    $json = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($payload))
    return $json | ConvertFrom-Json
}

if (-not $ClientId -or -not $Tenant -or -not $CertificateThumbprint) {
    throw 'ClientId, Tenant und CertificateThumbprint oben im Skript ausfuellen.'
}

# =============================================================================
Write-Host ''
Write-Host '=== 1) Zertifikat ===' -ForegroundColor Cyan
# =============================================================================
$storePaths = @(
    "Cert:\CurrentUser\My\$CertificateThumbprint",
    "Cert:\LocalMachine\My\$CertificateThumbprint"
)
$foundStores = @($storePaths | Where-Object { Test-Path -Path $_ })
if ($foundStores.Count -gt 0) {
    Write-Ok ("Zertifikat gefunden in: {0}" -f (($foundStores | Split-Path -Parent) -join ' und '))
    $cert = Get-Item -Path $foundStores[0]
    if ($cert.NotAfter -lt (Get-Date)) {
        Write-Bad ("Zertifikat ist ABGELAUFEN ({0})" -f $cert.NotAfter)
    }
    else {
        Write-Ok ("Gueltig bis {0}" -f $cert.NotAfter)
    }
}
else {
    Write-Bad ("Zertifikat mit Thumbprint {0} weder in CurrentUser\My noch LocalMachine\My gefunden." -f $CertificateThumbprint)
}

# =============================================================================
Write-Host ''
Write-Host '=== 2) Microsoft Graph (App-Only) ===' -ForegroundColor Cyan
# =============================================================================
$graphConnected = $false
try {
    $graphParams = @{ ClientId = $ClientId; TenantId = $Tenant; CertificateThumbprint = $CertificateThumbprint }
    if ((Get-Command Connect-MgGraph).Parameters.ContainsKey('NoWelcome')) { $graphParams['NoWelcome'] = $true }
    Connect-MgGraph @graphParams
    Write-Ok 'Graph-Verbindung hergestellt (Token wurde ausgestellt)'
    $graphConnected = $true
}
catch {
    Write-Bad ("Graph-Verbindung fehlgeschlagen: {0}" -f $_.Exception.Message)
    Write-Info 'AADSTS700016 = ClientId/Tenant falsch | AADSTS700027 = Zertifikat passt nicht zur App'
}

if ($graphConnected) {
    $graphTests = @(
        @{ Uri = 'v1.0/users?$top=1&$select=id';  Permission = 'User.Read.All' },
        @{ Uri = 'v1.0/groups?$top=1&$select=id'; Permission = 'Group.Read.All' },
        @{ Uri = 'v1.0/sites/getAllSites';        Permission = 'Sites.Read.All' }
    )
    foreach ($test in $graphTests) {
        try {
            $response = Invoke-MgGraphRequest -Method GET -Uri $test.Uri -ErrorAction Stop
            $count = @($response.value).Count
            Write-Ok ("{0} wirkt ({1} -> {2} Objekt(e) auf erster Seite)" -f $test.Permission, $test.Uri, $count)
        }
        catch {
            Write-Bad ("{0} FEHLT oder kein Admin Consent ({1}): {2}" -f $test.Permission, $test.Uri, $_.Exception.Message)
        }
    }
}

# =============================================================================
Write-Host ''
Write-Host '=== 3) SharePoint Admin (PnP, App-Only) ===' -ForegroundColor Cyan
# =============================================================================
if (-not $TenantAdminUrl) {
    Write-Info 'Uebersprungen (TenantAdminUrl leer - reines GraphOnly-Setup braucht keine SharePoint-Permission).'
}
elseif (-not (Get-Module -ListAvailable -Name 'PnP.PowerShell')) {
    Write-Info 'Uebersprungen (Modul PnP.PowerShell nicht installiert).'
}
else {
    try {
        Connect-PnPOnline -Url $TenantAdminUrl -ClientId $ClientId -Tenant $Tenant -Thumbprint $CertificateThumbprint
        Write-Ok 'PnP-Verbindung hergestellt (Token wurde ausgestellt)'

        # Token dekodieren und Rollen anzeigen
        try {
            $rawToken = $null
            try { $rawToken = Get-PnPAccessToken -ResourceTypeName SharePoint }
            catch { $rawToken = Get-PnPAccessToken }
            if ($rawToken) {
                $claims = ConvertFrom-JwtPayload -Token $rawToken
                Write-Info ("Token-Audience (aud): {0}" -f $claims.aud)
                $roles = @($claims.roles)
                if ($roles.Count -gt 0) {
                    Write-Info ("Token-Rollen (roles): {0}" -f ($roles -join ', '))
                }
                else {
                    Write-Bad 'Token enthaelt KEINE Rollen (roles-Claim leer) -> Application-Permission fehlt im Token: falsche API gewaehlt, Admin Consent fehlt, oder alte Session (neue PowerShell-Session starten).'
                }
                if ($claims.aud -like '*sharepoint.com*' -and $roles -notcontains 'Sites.FullControl.All') {
                    Write-Bad 'Sites.FullControl.All ist NICHT im SharePoint-Token. Wichtig: Die Permission muss im Portal unter der API "SharePoint" stehen - Sites.FullControl.All unter "Microsoft Graph" hilft fuer Get-PnPTenantSite NICHT.'
                }
            }
        }
        catch {
            Write-Info ("Token-Analyse nicht moeglich: {0}" -f $_.Exception.Message)
        }

        # Der entscheidende Test:
        try {
            $testSites = @(Get-PnPTenantSite -ErrorAction Stop | Select-Object -First 3)
            Write-Ok ("Get-PnPTenantSite funktioniert ({0}+ Sites lesbar) - SharePoint-Permission ist korrekt." -f $testSites.Count)
        }
        catch {
            Write-Bad ("Get-PnPTenantSite: {0}" -f $_.Exception.Message)
            Write-Info 'Bei "Unauthorized": Portal -> App -> API permissions -> "Add a permission" -> Kachel "SharePoint" (NICHT Microsoft Graph) -> Application permissions -> Sites.FullControl.All -> hinzufuegen -> "Grant admin consent" klicken -> NEUE PowerShell-Session starten.'
        }
    }
    catch {
        Write-Bad ("PnP-Verbindung fehlgeschlagen: {0}" -f $_.Exception.Message)
    }
}

# =============================================================================
Write-Host ''
Write-Host '=== Fazit ===' -ForegroundColor Cyan
# =============================================================================
if ($script:Problems.Count -eq 0) {
    Write-Host '  Alle Tests bestanden - der Report-Lauf sollte funktionieren.' -ForegroundColor Green
}
else {
    Write-Host ("  {0} Problem(e) gefunden:" -f $script:Problems.Count) -ForegroundColor Red
    $script:Problems | ForEach-Object { Write-Host ("   - {0}" -f $_) -ForegroundColor Red }
    Write-Host ''
    Write-Host '  Haeufigste Ursachen in dieser Reihenfolge pruefen:' -ForegroundColor Yellow
    Write-Host '   1. Permission unter der FALSCHEN API hinzugefuegt: Get-PnPTenantSite braucht'
    Write-Host '      "SharePoint -> Application -> Sites.FullControl.All", NICHT Microsoft Graph.'
    Write-Host '   2. "Grant admin consent for <Tenant>" nicht geklickt (gruene Haekchen pruefen).'
    Write-Host '   3. Alte Session: nach Berechtigungsaenderungen IMMER neue PowerShell-Session.'
    Write-Host '   4. GraphOnly-Alternative: Wenn die drei Graph-Tests OK sind, laeuft der Report'
    Write-Host '      mit $GraphOnly = $true sofort - ganz ohne SharePoint-Permission.'
}
