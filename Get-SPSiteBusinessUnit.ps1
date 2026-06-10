<#
.SYNOPSIS
    Inventarisiert alle SharePoint-Online-Sites des Tenants, loest deren
    Owner/Admins auf (inkl. Owner-Gruppen), ermittelt pro Person den
    Office-Standort (physicalDeliveryOfficeName / officeLocation) und leitet
    daraus per Mehrheitsentscheid die Business Unit der Site ab.
    Das Resultat wird als CSV exportiert.

.DESCRIPTION
    Ablauf:
      1. Liest alle Site Collections des Tenants (Get-PnPTenantSite).
      2. Loest pro Site die verantwortlichen Personen auf:
         - Microsoft-365-Gruppen-Sites (Teams, GROUP#0): Owner der M365-Gruppe
           via Microsoft Graph. Damit wird das Problem geloest, dass der Owner
           nur als "Owner-Gruppe von Site xyz" angezeigt wird
           (Claim: c:0o.c|federateddirectoryclaimprovider|<GroupId>_o).
         - Entra-ID-Security-Gruppen als Owner: transitive Aufloesung aller
           Mitglieder (inkl. verschachtelter Gruppen).
         - Einzelbenutzer: direkte Aufloesung.
         - Mit -DeepScan zusaetzlich pro Site: Site Collection Administratoren
           und die Mitglieder der zugehoerigen SharePoint-Besitzergruppe
           ("xyz - Besitzer" / "xyz Owners").
      3. Liest pro Benutzer das Graph-Attribut 'officeLocation'. Dieses
         entspricht dem AD-Attribut 'physicalDeliveryOfficeName' (wird von
         Entra Connect dorthin synchronisiert).
      4. Bildet die Mehrheit ueber alle Offices der aufgeloesten Personen und
         bestimmt darueber die Business Unit der Site. Das Mapping
         Office -> Business Unit kann in der Tabelle $OfficeToBusinessUnit
         (siehe unten) gepflegt werden; ohne Eintrag gilt das Office selbst
         als Business Unit. Gleichstaende werden ausgewiesen.
      5. Exportiert das Ergebnis als CSV (Standard: Semikolon-getrennt,
         UTF-8 mit BOM -> direkt in Excel verwendbar).

.PARAMETER TenantAdminUrl
    URL des SharePoint Admin Centers, z.B. https://contoso-admin.sharepoint.com
    (nicht noetig im -GraphOnly-Modus).

.PARAMETER ClientId
    App-ID (Client-ID) einer Entra-ID-App-Registrierung fuer PnP.PowerShell.
    Seit PnP.PowerShell 2.12 zwingend fuer interaktive Anmeldung. Eine eigene
    App laesst sich einmalig erstellen mit:
        Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnP-Reporting" -Tenant contoso.onmicrosoft.com

.PARAMETER Tenant
    Tenant fuer Cert-Authentifizierung (App-Only), z.B. contoso.onmicrosoft.com
    oder die Tenant-ID. Zwingend zusammen mit -CertificateThumbprint/-CertificatePath.

.PARAMETER CertificateThumbprint
    Thumbprint eines Zertifikats im Store CurrentUser\My, dessen oeffentlicher
    Teil (.cer) in der App-Registrierung hochgeladen ist. Aktiviert App-Only-Auth
    (unbeaufsichtigt, kein Benutzer-Login).

.PARAMETER CertificatePath
    Alternativ zu -CertificateThumbprint: Pfad zu einer PFX-Datei.

.PARAMETER CertificatePassword
    Passwort der PFX-Datei (SecureString), falls -CertificatePath verwendet wird.

.PARAMETER GraphOnly
    Inventarisiert die Sites rein ueber Microsoft Graph (sites/getAllSites)
    statt ueber das SharePoint Admin Center. Benoetigt KEINE SharePoint-
    Berechtigung auf der App - nur die Graph-Application-Permissions
    Sites.Read.All, Group.Read.All, User.Read.All (alle read-only).
    Nur mit Cert-Auth moeglich (getAllSites ist App-Only-only).
    Einschraenkungen: kein DeepScan, keine Site-Vorlage, Owner klassischer
    Sites via Eigentuemer der Standard-Dokumentbibliothek (Best Effort).

.PARAMETER OutputCsv
    Pfad der CSV-Ausgabedatei. Standard: .\SPSite-BusinessUnit-Report.csv

.PARAMETER CsvDelimiter
    CSV-Trennzeichen. Standard: ';' (Excel mit deutschen/schweizer
    Regionseinstellungen).

.PARAMETER IncludeOneDrive
    Nimmt auch die persoenlichen OneDrive-Sites in die Auswertung auf.

.PARAMETER DeepScan
    Verbindet sich zusaetzlich mit jeder einzelnen Site und liest dort die
    Site Collection Administratoren sowie die Mitglieder der
    SharePoint-Besitzergruppe. Genauer, aber deutlich langsamer.

.PARAMETER Limit
    Verarbeitet nur die ersten n Sites (fuer Tests). 0 = alle.

.PARAMETER ExcludeTemplates
    Site-Vorlagen, die uebersprungen werden (Systemsites).

.EXAMPLE
    .\Get-SPSiteBusinessUnit.ps1 -TenantAdminUrl https://contoso-admin.sharepoint.com -ClientId 11111111-2222-3333-4444-555555555555

.EXAMPLE
    .\Get-SPSiteBusinessUnit.ps1 -TenantAdminUrl https://contoso-admin.sharepoint.com -ClientId 1111... -DeepScan -OutputCsv C:\Temp\report.csv

.EXAMPLE
    # Testlauf mit 20 Sites
    .\Get-SPSiteBusinessUnit.ps1 -TenantAdminUrl https://contoso-admin.sharepoint.com -ClientId 1111... -Limit 20

.EXAMPLE
    # App-Only mit Zertifikat (unbeaufsichtigt, kein Benutzer-Login)
    .\Get-SPSiteBusinessUnit.ps1 -TenantAdminUrl https://contoso-admin.sharepoint.com -ClientId 1111... -Tenant contoso.onmicrosoft.com -CertificateThumbprint ABCDEF1234567890...

.EXAMPLE
    # Least Privilege: rein Graph-basiert, App braucht KEINE SharePoint-Berechtigung
    .\Get-SPSiteBusinessUnit.ps1 -GraphOnly -ClientId 1111... -Tenant contoso.onmicrosoft.com -CertificateThumbprint ABCDEF1234567890...

.NOTES
    Benoetigte Module (PowerShell 7.4+, empfohlen):
        Install-Module PnP.PowerShell                 -Scope CurrentUser
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
    Benoetigte Module bei PowerShell 7.0/7.1 (aeltere kompatible Versionen pinnen,
    aktuelle PnP-Versionen benoetigen PS 7.4+):
        Install-Module PnP.PowerShell                 -RequiredVersion 1.12.0 -Scope CurrentUser
        Install-Module Microsoft.Graph.Authentication -RequiredVersion 1.28.0 -Scope CurrentUser
        Hinweis: Register-PnPEntraIDAppForInteractiveLogin existiert in 1.12 nicht -
        die Entra-App-Registrierung dann manuell im Portal anlegen (siehe README).
    Im GraphOnly-Modus wird PnP.PowerShell nicht benoetigt.

    Benoetigte Berechtigungen je Modus:
      Interaktiv (delegiert, Standard):
        - Benutzer braucht Rolle SharePoint-Administrator
        - App: SharePoint delegiert AllSites.FullControl (effektive Rechte =
          Schnittmenge App + Benutzer); Graph-Login laeuft ueber die
          Microsoft-Graph-PowerShell-App (User.Read.All, Group.Read.All,
          Sites.Read.All werden beim Login konsentiert)
      App-Only mit Zertifikat (voller Funktionsumfang):
        - App: SharePoint APPLICATION Sites.FullControl.All (fuer
          Get-PnPTenantSite/Admin-CSOM zwingend - granularer geht es dort nicht)
        - App: Graph APPLICATION User.Read.All, Group.Read.All, Sites.Read.All
      App-Only -GraphOnly (Least Privilege, alles read-only):
        - App: NUR Graph APPLICATION Sites.Read.All, Group.Read.All, User.Read.All
        - keinerlei SharePoint-Berechtigung noetig
#>
[CmdletBinding()]
param(
    # =========================================================================
    # >>> KONFIGURATION: DIESE WERTE AUSFUELLEN <<<
    # Alle Werte koennen alternativ auch als Parameter uebergeben werden,
    # z.B. .\Get-SPSiteBusinessUnit.ps1 -ClientId xxx -GraphOnly
    # Das Office->Business-Unit-Mapping steht direkt unter dem param-Block.
    # =========================================================================

    # App-ID (Client-ID) der eigenen Entra-App-Registrierung:
    [string]$ClientId = '',

    # Eigener Tenant, z.B. 'contoso.onmicrosoft.com' oder die Tenant-ID
    # (zwingend bei Cert-Authentifizierung):
    [string]$Tenant = '',

    # SharePoint Admin Center, z.B. 'https://contoso-admin.sharepoint.com'
    # (darf leer bleiben, wenn GraphOnly = $true):
    [string]$TenantAdminUrl = '',

    # Thumbprint des App-Zertifikats im Zertifikatsspeicher.
    # Leer lassen = interaktiver Browser-Login statt App-Only:
    [string]$CertificateThumbprint = '',

    # Alternativ zum Thumbprint: Pfad zu einer PFX-Datei. Passwort beim Aufruf
    # mitgeben: -CertificatePassword (Read-Host -AsSecureString)
    [string]$CertificatePath = '',
    [securestring]$CertificatePassword,

    # $true = Least-Privilege-Modus rein ueber Microsoft Graph: braucht KEINE
    # SharePoint-Berechtigung auf der App, nur die Graph-Application-Permissions
    # Sites.Read.All, Group.Read.All, User.Read.All. Erfordert Cert-Auth.
    # Einschraenkungen: kein DeepScan, keine Site-Vorlage in der Ausgabe.
    [switch]$GraphOnly = $false,

    # Pfad der CSV-Ausgabedatei:
    [string]$OutputCsv = '.\SPSite-BusinessUnit-Report.csv',

    # CSV-Trennzeichen (';' = Excel mit deutschen/schweizer Einstellungen):
    [string]$CsvDelimiter = ';',

    # $true = persoenliche OneDrive-Sites mit auswerten:
    [switch]$IncludeOneDrive = $false,

    # $true = zusaetzlich Site Collection Admins + SP-Besitzergruppen pro Site
    # einlesen (genauer, aber deutlich langsamer; nicht mit GraphOnly moeglich):
    [switch]$DeepScan = $false,

    # Nur die ersten n Sites verarbeiten (fuer Tests; 0 = alle):
    [int]$Limit = 0,

    # Site-Vorlagen, die uebersprungen werden (Systemsites):
    [string[]]$ExcludeTemplates = @(
        'SRCHCEN#0',            # Suchcenter
        'SPSMSITEHOST#0',       # MySite-Host
        'APPCATALOG#0',         # App-Katalog
        'POINTPUBLISHINGHUB#0', # Video/Stream
        'POINTPUBLISHINGTOPIC#0',
        'EDISC#0',              # eDiscovery
        'RedirectSite#0'        # Umleitungs-Sites
    )
)

# =============================================================================
# Mapping Office-Standort -> Business Unit
# Hier die eigenen Standorte pflegen. Offices ohne Eintrag werden 1:1 als
# Business Unit uebernommen.
# =============================================================================
$OfficeToBusinessUnit = @{
    # 'Zuerich HQ'   = 'BU Corporate'
    # 'Basel'        = 'BU Pharma'
    # 'Geneve'       = 'BU Finance'
}

# =============================================================================
# Caches, damit Benutzer und Gruppen nur einmal via Graph abgefragt werden
# =============================================================================
$script:UserCache  = @{}   # Key: UPN oder ObjectId -> Benutzerobjekt (oder $null)
$script:GroupCache = @{}   # Key: "<GroupId>|Owners" / "<GroupId>|Members" -> Benutzerliste

# =============================================================================
# Hilfsfunktionen
# =============================================================================

function Invoke-GraphGetAll {
    <# Holt eine Graph-Collection inkl. aller Folgeseiten (Paging). #>
    param([Parameter(Mandatory)][string]$Uri)
    $items = @()
    $next = $Uri
    while ($next) {
        $response = Invoke-MgGraphRequest -Method GET -Uri $next -ErrorAction Stop
        if ($response.value) { $items += $response.value }
        $next = $response.'@odata.nextLink'
    }
    return $items
}

function ConvertTo-UserObject {
    <# Normalisiert eine Graph-User-Antwort und befuellt den Cache. #>
    param($GraphUser)
    $office = ('' + $GraphUser.officeLocation).Trim()
    if (-not $office) { $office = $null }
    $user = [pscustomobject]@{
        Id          = $GraphUser.id
        DisplayName = $GraphUser.displayName
        Upn         = $GraphUser.userPrincipalName
        Office      = $office   # = AD-Attribut physicalDeliveryOfficeName
    }
    if ($user.Id)  { $script:UserCache[$user.Id]  = $user }
    if ($user.Upn) { $script:UserCache[$user.Upn] = $user }
    return $user
}

function Get-CachedUser {
    <# Liest einen Benutzer (inkl. officeLocation) via Graph, mit Cache. #>
    param([Parameter(Mandatory)][string]$IdOrUpn)
    if ($script:UserCache.ContainsKey($IdOrUpn)) { return $script:UserCache[$IdOrUpn] }
    $user = $null
    try {
        $uri = 'v1.0/users/{0}?$select=id,displayName,userPrincipalName,officeLocation' -f [uri]::EscapeDataString($IdOrUpn)
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
        $user = ConvertTo-UserObject -GraphUser $response
    }
    catch {
        Write-Verbose ("Benutzer '{0}' nicht aufloesbar: {1}" -f $IdOrUpn, $_.Exception.Message)
    }
    $script:UserCache[$IdOrUpn] = $user
    return $user
}

function Get-GroupUsers {
    <#
        Loest eine Entra-/M365-Gruppe in Benutzer auf.
        Owners  = direkte Besitzer der Gruppe
        Members = transitive Mitglieder (inkl. verschachtelter Gruppen)
    #>
    param(
        [Parameter(Mandatory)][string]$GroupId,
        [Parameter(Mandatory)][ValidateSet('Owners', 'Members')][string]$Mode
    )
    $cacheKey = '{0}|{1}' -f $GroupId, $Mode
    if ($script:GroupCache.ContainsKey($cacheKey)) { return $script:GroupCache[$cacheKey] }

    $segment = 'owners'
    if ($Mode -eq 'Members') { $segment = 'transitiveMembers' }

    $users = @()
    try {
        # Cast auf microsoft.graph.user: liefert nur Benutzer und erlaubt $select auf officeLocation
        $uri = 'v1.0/groups/{0}/{1}/microsoft.graph.user?$select=id,displayName,userPrincipalName,officeLocation&$top=999' -f $GroupId, $segment
        $rawUsers = Invoke-GraphGetAll -Uri $uri
        foreach ($raw in $rawUsers) {
            $users += ConvertTo-UserObject -GraphUser $raw
        }
    }
    catch {
        Write-Warning ("Gruppe {0} ({1}) konnte nicht aufgeloest werden: {2}" -f $GroupId, $Mode, $_.Exception.Message)
    }
    $script:GroupCache[$cacheKey] = $users
    return $users
}

function Resolve-LoginToUsers {
    <#
        Loest einen SharePoint-LoginName/Claim in echte Benutzer auf.
        Unterstuetzt:
          c:0o.c|federateddirectoryclaimprovider|<guid>_o  -> Owner einer M365-Gruppe
          c:0o.c|federateddirectoryclaimprovider|<guid>    -> Mitglieder einer M365-Gruppe
          c:0t.c|tenant|<guid>                             -> Entra-Security-Gruppe (transitiv)
          i:0#.f|membership|user@contoso.com               -> Einzelbenutzer
          user@contoso.com                                 -> Einzelbenutzer
    #>
    param([string]$LoginName)

    $result = [pscustomobject]@{ Users = @(); Source = 'Unbekannt' }
    if ([string]::IsNullOrWhiteSpace($LoginName)) { return $result }

    # --- M365-Gruppe ("Owner-Gruppe von Site xyz") ---
    if ($LoginName -match 'federateddirectoryclaimprovider\|([0-9a-fA-F\-]{36})(_o)?') {
        $groupId  = $Matches[1]
        $isOwners = [bool]$Matches[2]
        if ($isOwners) {
            $result.Users  = @(Get-GroupUsers -GroupId $groupId -Mode Owners)
            $result.Source = 'M365-Gruppe (Owner)'
            if ($result.Users.Count -eq 0) {
                # verwaiste Gruppe ohne Owner -> Mitglieder als Fallback
                $result.Users  = @(Get-GroupUsers -GroupId $groupId -Mode Members)
                $result.Source = 'M365-Gruppe (Mitglieder, Fallback ohne Owner)'
            }
        }
        else {
            $result.Users  = @(Get-GroupUsers -GroupId $groupId -Mode Members)
            $result.Source = 'M365-Gruppe (Mitglieder)'
        }
        return $result
    }

    # --- Entra-ID-Security-Gruppe ---
    if ($LoginName -match 'tenant\|([0-9a-fA-F\-]{36})') {
        $result.Users  = @(Get-GroupUsers -GroupId $Matches[1] -Mode Members)
        $result.Source = 'Security-Gruppe (Mitglieder)'
        return $result
    }

    # --- Einzelbenutzer (Claim oder reine UPN) ---
    $upn = $LoginName
    if ($LoginName -match '\|membership\|(.+)$') { $upn = $Matches[1] }
    if ($upn -like '*@*') {
        $user = Get-CachedUser -IdOrUpn $upn
        if ($user) {
            $result.Users  = @($user)
            $result.Source = 'Einzelbenutzer'
        }
        return $result
    }

    # --- Letzter Versuch: unbekannter Claim mit GUID -> als Gruppe behandeln ---
    if ($LoginName -match '([0-9a-fA-F\-]{36})') {
        $users = @(Get-GroupUsers -GroupId $Matches[1] -Mode Members)
        if ($users.Count -gt 0) {
            $result.Users  = $users
            $result.Source = 'Gruppe (Mitglieder)'
        }
    }
    return $result
}

function Resolve-SiteOwners {
    <#
        Sammelt fuer eine Site alle verantwortlichen Personen:
          - M365-Gruppen-Site: Owner der Gruppe (Fallback: Mitglieder)
          - klassische Site:   Owner-Feld der Site Collection
          - DeepScan:          + Site Collection Admins
                               + Mitglieder der SharePoint-Besitzergruppe
        Gibt deduplizierte Benutzerliste, Quellen und Hinweise zurueck.
    #>
    param($Site, $SiteConnection)

    $pool    = [System.Collections.Generic.List[object]]::new()
    $sources = [System.Collections.Generic.List[string]]::new()
    $notes   = [System.Collections.Generic.List[string]]::new()

    # --- 1) Tenant-Ebene: M365-Gruppe oder Owner-Feld ---
    if ($Site.GroupId -and $Site.GroupId -ne [guid]::Empty) {
        $users = @(Get-GroupUsers -GroupId $Site.GroupId -Mode Owners)
        if ($users.Count -gt 0) {
            $sources.Add('M365-Gruppe (Owner)')
        }
        else {
            $users = @(Get-GroupUsers -GroupId $Site.GroupId -Mode Members)
            if ($users.Count -gt 0) {
                $sources.Add('M365-Gruppe (Mitglieder, Fallback ohne Owner)')
                $notes.Add('M365-Gruppe hat keine Owner - Mitglieder verwendet')
            }
            else {
                $notes.Add('M365-Gruppe konnte nicht aufgeloest werden (geloescht?)')
            }
        }
        foreach ($u in $users) { $pool.Add($u) }
    }
    else {
        $ownerLogin = $Site.OwnerLoginName
        if (-not $ownerLogin) { $ownerLogin = $Site.Owner }
        if ($ownerLogin) {
            $resolved = Resolve-LoginToUsers -LoginName $ownerLogin
            if ($resolved.Users.Count -gt 0) {
                $sources.Add($resolved.Source)
                foreach ($u in $resolved.Users) { $pool.Add($u) }
            }
            else {
                $notes.Add(("Owner '{0}' konnte nicht aufgeloest werden" -f $ownerLogin))
            }
        }
        else {
            $notes.Add('Kein Owner auf Tenant-Ebene hinterlegt')
        }
    }

    # --- 2) DeepScan: Site Collection Admins + SharePoint-Besitzergruppe ---
    if ($SiteConnection) {
        try {
            $admins = Get-PnPSiteCollectionAdmin -Connection $SiteConnection
            foreach ($admin in $admins) {
                $resolved = Resolve-LoginToUsers -LoginName $admin.LoginName
                if ($resolved.Users.Count -gt 0) {
                    $sources.Add(('Site-Admin: {0}' -f $resolved.Source))
                    foreach ($u in $resolved.Users) { $pool.Add($u) }
                }
            }
        }
        catch {
            $notes.Add(('Site-Admins nicht lesbar: {0}' -f $_.Exception.Message))
        }

        try {
            $ownerGroup = Get-PnPGroup -AssociatedOwnerGroup -Connection $SiteConnection
            if ($ownerGroup) {
                $members = Get-PnPGroupMember -Group $ownerGroup -Connection $SiteConnection
                foreach ($member in $members) {
                    $resolved = Resolve-LoginToUsers -LoginName $member.LoginName
                    if ($resolved.Users.Count -gt 0) {
                        $sources.Add(('SP-Besitzergruppe: {0}' -f $resolved.Source))
                        foreach ($u in $resolved.Users) { $pool.Add($u) }
                    }
                }
            }
        }
        catch {
            $notes.Add(('SP-Besitzergruppe nicht lesbar: {0}' -f $_.Exception.Message))
        }
    }

    # --- Deduplizieren (gleicher Benutzer aus mehreren Quellen zaehlt 1x) ---
    $unique = @($pool | Where-Object { $_ } | Group-Object -Property Id | ForEach-Object { $_.Group[0] })

    return [pscustomobject]@{
        Users   = $unique
        Sources = @($sources | Sort-Object -Unique)
        Notes   = @($notes)
    }
}

function Get-MajorityResult {
    <#
        Mehrheitsentscheid ueber die Offices der uebergebenen Benutzer.
        Prozentbasis sind nur Benutzer MIT gepflegtem Office; Benutzer ohne
        Office werden separat gezaehlt. Gleichstaende werden markiert.
    #>
    param($Users)

    $result = [pscustomobject]@{
        OfficeCounts   = ''       # z.B. "Zuerich (4) | Basel (2)"
        MajorityOffice = ''
        SharePercent   = 0
        IsTie          = $false
        NoOfficeCount  = 0
    }

    $all        = @($Users)
    $withOffice = @($all | Where-Object { $_.Office })
    $result.NoOfficeCount = $all.Count - $withOffice.Count
    if ($withOffice.Count -eq 0) { return $result }

    $grouped = @($withOffice | Group-Object -Property Office |
        Sort-Object -Property @{ Expression = 'Count'; Descending = $true }, @{ Expression = 'Name'; Descending = $false })

    $top  = $grouped[0]
    $tied = @($grouped | Where-Object { $_.Count -eq $top.Count })

    $result.OfficeCounts = ($grouped | ForEach-Object { '{0} ({1})' -f $_.Name, $_.Count }) -join ' | '
    $result.IsTie        = ($tied.Count -gt 1)
    if ($result.IsTie) {
        $result.MajorityOffice = ($tied | ForEach-Object { $_.Name }) -join ' | '
    }
    else {
        $result.MajorityOffice = $top.Name
    }
    $result.SharePercent = [math]::Round(100 * $top.Count / $withOffice.Count, 1)
    return $result
}

function Get-BusinessUnit {
    <#
        Mappt Office(s) auf die Business Unit. Bei Gleichstand: wenn alle
        beteiligten Offices auf dieselbe BU zeigen, gilt diese - sonst
        "Unbestimmt (Gleichstand: ...)".
    #>
    param([string[]]$Offices)

    $valid = @($Offices | Where-Object { $_ })
    if ($valid.Count -eq 0) { return 'Unbestimmt (kein Office)' }

    $units = @($valid | ForEach-Object {
            if ($OfficeToBusinessUnit.ContainsKey($_)) { $OfficeToBusinessUnit[$_] } else { $_ }
        } | Sort-Object -Unique)

    if ($units.Count -eq 1) { return [string]$units[0] }
    return ('Unbestimmt (Gleichstand: {0})' -f ($units -join ' | '))
}

function Get-SiteCreator {
    <#
        Best-Effort-Ermittlung des Site-Erstellers via Graph (createdBy).
        Graph liefert dieses Feld nicht fuer alle Sites - dann bleibt die
        Spalte leer. Verlaesslich ist der Ersteller nur im Unified Audit Log.
    #>
    param([string]$SiteUrl)
    try {
        $uri = [uri]$SiteUrl
        $sitePath = $uri.AbsolutePath.TrimEnd('/')
        $graphSiteId = $uri.Host
        if ($sitePath) { $graphSiteId = '{0}:{1}' -f $uri.Host, $sitePath }
        $response = Invoke-MgGraphRequest -Method GET -Uri ('v1.0/sites/{0}?$select=createdBy,createdDateTime' -f $graphSiteId) -ErrorAction Stop
        $creator = $response.createdBy.user
        if ($creator -and ($creator.displayName -or $creator.email)) {
            return ('{0} <{1}>' -f $creator.displayName, $creator.email).Trim()
        }
    }
    catch {
        Write-Verbose ("Ersteller fuer {0} nicht ermittelbar: {1}" -f $SiteUrl, $_.Exception.Message)
    }
    return ''
}

function Get-GraphSites {
    <#
        Site-Inventar rein ueber Microsoft Graph (sites/getAllSites, nur App-Only).
        Benoetigt nur die Graph-Application-Permission Sites.Read.All.
        Owner-Ermittlung ueber den Eigentuemer der Standard-Dokumentbibliothek
        (drive.owner): bei Gruppen-Sites die M365-Gruppe, sonst ein Benutzer.
        Liefert Objekte mit denselben Properties, die die Hauptschleife auch
        von Get-PnPTenantSite erwartet (Url, Title, Template, GroupId, Owner,
        OwnerLoginName).
    #>
    param([bool]$IncludePersonal)

    $sites = @()
    $raw = Invoke-GraphGetAll -Uri 'v1.0/sites/getAllSites'
    foreach ($s in $raw) {
        $isPersonal = ($s.isPersonalSite -eq $true) -or ($s.webUrl -like '*-my.sharepoint.com*')
        if ($isPersonal -and -not $IncludePersonal) { continue }

        $groupId    = [guid]::Empty
        $ownerRaw   = ''
        $ownerLogin = ''
        try {
            $drive = Invoke-MgGraphRequest -Method GET -Uri ('v1.0/sites/{0}/drive?$select=owner' -f $s.id) -ErrorAction Stop
            if ($drive.owner.group.id) {
                $groupId  = [guid]$drive.owner.group.id
                $ownerRaw = ('M365-Gruppe: {0}' -f $drive.owner.group.displayName)
            }
            elseif ($drive.owner.user) {
                $ownerRaw = ('Benutzer: {0}' -f $drive.owner.user.displayName)
                $ownerUser = $null
                if ($drive.owner.user.email) { $ownerUser = Get-CachedUser -IdOrUpn $drive.owner.user.email }
                elseif ($drive.owner.user.id) { $ownerUser = Get-CachedUser -IdOrUpn $drive.owner.user.id }
                if ($ownerUser) { $ownerLogin = $ownerUser.Upn }
            }
        }
        catch {
            Write-Verbose ("Drive-Owner fuer {0} nicht lesbar: {1}" -f $s.webUrl, $_.Exception.Message)
        }

        $sites += [pscustomobject]@{
            Url            = $s.webUrl
            Title          = $s.displayName
            Template       = ''
            GroupId        = $groupId
            Owner          = $ownerRaw
            OwnerLoginName = $ownerLogin
        }
    }
    return $sites
}

# =============================================================================
# Validierung & Modul-Check
# =============================================================================
$AppOnly = [bool]($CertificateThumbprint -or $CertificatePath)

if ($AppOnly -and -not ($ClientId -and $Tenant)) {
    throw 'Cert-Authentifizierung (App-Only) benoetigt zusaetzlich -ClientId und -Tenant (z.B. contoso.onmicrosoft.com).'
}
if ($GraphOnly -and -not $AppOnly) {
    throw 'Der GraphOnly-Modus benoetigt App-Only-Auth: -ClientId, -Tenant und -CertificateThumbprint oder -CertificatePath angeben (sites/getAllSites ist nur App-Only verfuegbar).'
}
if (-not $GraphOnly -and -not $TenantAdminUrl) {
    throw 'Ohne -GraphOnly muss -TenantAdminUrl angegeben werden (https://<tenant>-admin.sharepoint.com).'
}
if ($GraphOnly -and $DeepScan) {
    Write-Warning 'DeepScan ist im GraphOnly-Modus nicht verfuegbar (kein SharePoint-CSOM) und wird ignoriert.'
    $DeepScan = $false
}

$requiredModules = @('Microsoft.Graph.Authentication')
if (-not $GraphOnly) { $requiredModules += 'PnP.PowerShell' }
foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        throw ("Benoetigtes Modul '{0}' fehlt. Installation: Install-Module {0} -Scope CurrentUser" -f $module)
    }
}

# =============================================================================
# Verbindungen herstellen
# =============================================================================
if ($CertificatePath) { $CertificatePath = (Resolve-Path -Path $CertificatePath).Path }

# PnP-Auth-Parameter zentral aufbauen - werden auch fuer die DeepScan-
# Verbindungen pro Site wiederverwendet
$script:PnPAuthParams = @{}
if ($ClientId) { $script:PnPAuthParams['ClientId'] = $ClientId }
if ($AppOnly) {
    $script:PnPAuthParams['Tenant'] = $Tenant
    if ($CertificateThumbprint) {
        $script:PnPAuthParams['Thumbprint'] = $CertificateThumbprint
    }
    else {
        $script:PnPAuthParams['CertificatePath'] = $CertificatePath
        if ($CertificatePassword) { $script:PnPAuthParams['CertificatePassword'] = $CertificatePassword }
    }
}
else {
    $script:PnPAuthParams['Interactive'] = $true
}

if (-not $GraphOnly) {
    Write-Host ("Verbinde mit SharePoint Admin Center: {0}" -f $TenantAdminUrl) -ForegroundColor Cyan
    $pnpParams = $script:PnPAuthParams.Clone()
    $pnpParams['Url'] = $TenantAdminUrl
    Connect-PnPOnline @pnpParams
}

Write-Host 'Verbinde mit Microsoft Graph...' -ForegroundColor Cyan
if ($AppOnly) {
    $graphParams = @{ ClientId = $ClientId; TenantId = $Tenant }
    if ($CertificateThumbprint) {
        $graphParams['CertificateThumbprint'] = $CertificateThumbprint
    }
    elseif ($CertificatePassword) {
        $graphParams['Certificate'] = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($CertificatePath, $CertificatePassword)
    }
    else {
        $graphParams['Certificate'] = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($CertificatePath)
    }
}
else {
    $graphParams = @{ Scopes = @('User.Read.All', 'Group.Read.All', 'Sites.Read.All') }
}
# -NoWelcome gibt es erst ab Microsoft.Graph 2.x - auf aelteren SDKs weglassen
if ((Get-Command Connect-MgGraph).Parameters.ContainsKey('NoWelcome')) { $graphParams['NoWelcome'] = $true }
Connect-MgGraph @graphParams

# =============================================================================
# Sites einlesen
# =============================================================================
Write-Host 'Lese Site Collections des Tenants...' -ForegroundColor Cyan
if ($GraphOnly) {
    $sites = @(Get-GraphSites -IncludePersonal $IncludeOneDrive.IsPresent |
        Where-Object { $_.Template -notin $ExcludeTemplates })
}
else {
    $sites = @(Get-PnPTenantSite -IncludeOneDriveSites:$IncludeOneDrive |
        Where-Object { $_.Template -notin $ExcludeTemplates })
}

if ($Limit -gt 0) { $sites = @($sites | Select-Object -First $Limit) }
Write-Host ("{0} Sites werden analysiert." -f $sites.Count) -ForegroundColor Cyan

# =============================================================================
# Hauptschleife
# =============================================================================
$results = [System.Collections.Generic.List[object]]::new()
$counter = 0

foreach ($site in $sites) {
    $counter++
    Write-Progress -Activity 'Analysiere SharePoint-Sites' -Status ("[{0}/{1}] {2}" -f $counter, $sites.Count, $site.Url) `
        -PercentComplete ([int](100 * $counter / [math]::Max(1, $sites.Count)))

    # Optionale Site-Verbindung fuer DeepScan (Token wird wiederverwendet,
    # es erfolgt keine erneute Anmeldung pro Site)
    $siteConnection = $null
    $deepScanNote = ''
    if ($DeepScan) {
        try {
            $deepParams = $script:PnPAuthParams.Clone()
            $deepParams['Url'] = $site.Url
            $deepParams['ReturnConnection'] = $true
            $siteConnection = Connect-PnPOnline @deepParams
        }
        catch {
            $deepScanNote = ('DeepScan-Verbindung fehlgeschlagen: {0}' -f $_.Exception.Message)
        }
    }

    $owners   = Resolve-SiteOwners -Site $site -SiteConnection $siteConnection
    $majority = Get-MajorityResult -Users $owners.Users

    if ($majority.MajorityOffice) {
        $businessUnit = Get-BusinessUnit -Offices ($majority.MajorityOffice -split ' \| ')
    }
    else {
        $businessUnit = 'Unbestimmt (kein Office)'
    }

    $notes = @($owners.Notes)
    if ($deepScanNote) { $notes += $deepScanNote }

    $ownerRaw = $site.OwnerLoginName
    if (-not $ownerRaw) { $ownerRaw = $site.Owner }

    $results.Add([pscustomobject]@{
            SiteUrl               = $site.Url
            SiteTitel             = $site.Title
            Vorlage               = $site.Template
            GroupId               = if ($site.GroupId -and $site.GroupId -ne [guid]::Empty) { [string]$site.GroupId } else { '' }
            OwnerRoh              = $ownerRaw
            OwnerQuelle           = ($owners.Sources -join ' | ')
            AnzahlOwner           = $owners.Users.Count
            Owners                = ($owners.Users | ForEach-Object { '{0} <{1}>' -f $_.DisplayName, $_.Upn }) -join ' | '
            OfficeVerteilung      = $majority.OfficeCounts
            OwnerOhneOffice       = $majority.NoOfficeCount
            MehrheitsOffice       = $majority.MajorityOffice
            MehrheitAnteilProzent = $majority.SharePercent
            Gleichstand           = if ($majority.IsTie) { 'Ja' } else { 'Nein' }
            BusinessUnit          = $businessUnit
            Ersteller             = Get-SiteCreator -SiteUrl $site.Url
            Hinweis               = ($notes -join ' | ')
        })
}

Write-Progress -Activity 'Analysiere SharePoint-Sites' -Completed

# =============================================================================
# CSV-Export
# =============================================================================
$csvParams = @{
    Path              = $OutputCsv
    NoTypeInformation = $true
    Delimiter         = $CsvDelimiter
}
if ($PSVersionTable.PSVersion.Major -ge 6) { $csvParams['Encoding'] = 'utf8BOM' }
else { $csvParams['Encoding'] = 'UTF8' }

$results | Export-Csv @csvParams

Write-Host ''
Write-Host ("Fertig: {0} Sites ausgewertet." -f $results.Count) -ForegroundColor Green
Write-Host ("CSV-Report: {0}" -f (Resolve-Path -Path $OutputCsv)) -ForegroundColor Green

# Kurze Zusammenfassung der BU-Verteilung in der Konsole
$results | Group-Object -Property BusinessUnit | Sort-Object -Property Count -Descending |
    Select-Object @{ Name = 'BusinessUnit'; Expression = { $_.Name } }, Count |
    Format-Table -AutoSize
