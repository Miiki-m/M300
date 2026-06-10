# M300

## SharePoint Site → Business Unit Report (`Get-SPSiteBusinessUnit.ps1`)

PowerShell-Skript, das alle SharePoint-Online-Sites des Tenants inventarisiert,
deren Owner/Admins auflöst, pro Person den Office-Standort
(`physicalDeliveryOfficeName` / Graph: `officeLocation`) ermittelt und daraus
per **Mehrheitsentscheid** die Business Unit der Site ableitet. Resultat: CSV.

### Was das Skript löst

In SharePoint Online wird der Owner einer Site oft nur als
**"Owner-Gruppe von Site xyz"** angezeigt (technisch der Claim
`c:0o.c|federateddirectoryclaimprovider|<GroupId>_o`). Das Skript löst solche
Gruppen via Microsoft Graph in echte Benutzer auf und bildet dann die Mehrheit
über deren Office-Standorte:

1. **Alle Sites lesen** – `Get-PnPTenantSite` (OneDrive optional über `-IncludeOneDrive`)
2. **Owner auflösen** – je nach Site-Typ:
   - M365-Gruppen-Site (Teams, `GROUP#0`): Owner der M365-Gruppe; hat die
     Gruppe keine Owner mehr (verwaist), Fallback auf die Mitglieder
   - Entra-Security-Gruppe als Owner: transitive Auflösung (inkl. verschachtelter Gruppen)
   - Einzelbenutzer: direkte Auflösung
   - mit `-DeepScan` zusätzlich: Site Collection Admins + Mitglieder der
     SharePoint-Besitzergruppe («xyz – Besitzer»)
3. **Office pro Person** – Graph-Attribut `officeLocation`
   (= AD-Attribut `physicalDeliveryOfficeName`, via Entra Connect synchronisiert)
4. **Mehrheitsrechnung** – häufigstes Office unter allen aufgelösten Personen
   bestimmt die Business Unit; Gleichstände werden markiert
5. **CSV-Export** – Semikolon-getrennt, UTF-8 mit BOM (direkt Excel-tauglich)

### Voraussetzungen

```powershell
Install-Module PnP.PowerShell                 -Scope CurrentUser
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
```

- PowerShell 7.4+ empfohlen (Voraussetzung aktueller PnP.PowerShell-Versionen)
- Rolle **SharePoint-Administrator** (für `Get-PnPTenantSite`)
- Graph-Scopes (delegiert, Admin Consent nötig): `User.Read.All`, `Group.Read.All`, `Sites.Read.All`
- Eigene Entra-App-Registrierung für PnP (seit PnP.PowerShell 2.12 Pflicht), einmalig:

```powershell
Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnP-Reporting" -Tenant contoso.onmicrosoft.com
```

### Verwendung

```powershell
# Standardlauf (schnell, Tenant-Ebene)
.\Get-SPSiteBusinessUnit.ps1 -TenantAdminUrl https://contoso-admin.sharepoint.com -ClientId <App-Id>

# Gründlich: zusätzlich Site Collection Admins + SP-Besitzergruppen pro Site
.\Get-SPSiteBusinessUnit.ps1 -TenantAdminUrl https://contoso-admin.sharepoint.com -ClientId <App-Id> -DeepScan

# Testlauf mit 20 Sites, eigener Output-Pfad
.\Get-SPSiteBusinessUnit.ps1 -TenantAdminUrl https://contoso-admin.sharepoint.com -ClientId <App-Id> -Limit 20 -OutputCsv C:\Temp\report.csv
```

### Authentifizierung & Berechtigungen

Das Skript unterstützt drei Modi:

| Modus | Aufruf | Benötigte App-Berechtigungen |
|---|---|---|
| **Interaktiv** (delegiert) | `-ClientId` | SharePoint *delegiert* `AllSites.FullControl` – effektive Rechte = Schnittmenge aus App und angemeldetem Benutzer (der SharePoint-Admin sein muss). Graph-Login läuft über die Microsoft-Graph-PowerShell-App. |
| **App-Only mit Zertifikat** (voller Funktionsumfang) | `-ClientId -Tenant -CertificateThumbprint` | SharePoint **Application** `Sites.FullControl.All` **+** Graph **Application** `Sites.Read.All`, `Group.Read.All`, `User.Read.All` |
| **App-Only `-GraphOnly`** (Least Privilege) | `-GraphOnly -ClientId -Tenant -CertificateThumbprint` | **Nur** Graph **Application** `Sites.Read.All`, `Group.Read.All`, `User.Read.All` (alles read-only) – keinerlei SharePoint-Berechtigung |

**Wichtig:** Für App-Only zählen ausschliesslich *Application*-Permissions
(delegierte Scopes sind dann wirkungslos), und `Get-PnPTenantSite` braucht
App-Only zwingend SharePoint `Sites.FullControl.All` – eine granularere
Admin-Berechtigung existiert nicht. Wer das (zu Recht) als überprivilegiert
einstuft, nutzt den `-GraphOnly`-Modus: Site-Inventar via Graph
`sites/getAllSites`, Owner-Ermittlung über die M365-Gruppe bzw. den
Eigentümer der Standard-Dokumentbibliothek. Einschränkungen: kein `-DeepScan`
(keine Site Collection Admins / SP-Besitzergruppen), keine Site-Vorlage in der
Ausgabe.

**Zertifikat erstellen und hochladen (einmalig, auf Windows):**

```powershell
# Self-signed Zertifikat erzeugen (3 Jahre gueltig, landet in CurrentUser\My)
$cert = New-SelfSignedCertificate -Subject "CN=PnP-Reporting" `
    -CertStoreLocation "Cert:\CurrentUser\My" -KeyExportPolicy Exportable `
    -KeySpec Signature -KeyLength 2048 -NotAfter (Get-Date).AddYears(3)

# Oeffentlichen Teil exportieren -> im Portal unter
# App registrations -> (App) -> "Certificates & secrets" -> "Upload certificate" hochladen
Export-Certificate -Cert $cert -FilePath .\PnP-Reporting.cer

# Thumbprint anzeigen (wird dem Skript uebergeben)
$cert.Thumbprint
```

Aufruf danach:

```powershell
# Voller Funktionsumfang (App braucht SharePoint Sites.FullControl.All):
.\Get-SPSiteBusinessUnit.ps1 -TenantAdminUrl https://contoso-admin.sharepoint.com `
    -ClientId <AppId> -Tenant contoso.onmicrosoft.com -CertificateThumbprint <Thumbprint>

# Least Privilege ohne SharePoint-Berechtigung:
.\Get-SPSiteBusinessUnit.ps1 -GraphOnly `
    -ClientId <AppId> -Tenant contoso.onmicrosoft.com -CertificateThumbprint <Thumbprint>
```

Statt Thumbprint (Zertifikatsspeicher) geht auch eine PFX-Datei:
`-CertificatePath .\cert.pfx -CertificatePassword (Read-Host -AsSecureString)`.

### Office → Business-Unit-Mapping

Im Skript die Tabelle `$OfficeToBusinessUnit` pflegen:

```powershell
$OfficeToBusinessUnit = @{
    'Zuerich HQ' = 'BU Corporate'
    'Basel'      = 'BU Pharma'
}
```

Offices ohne Mapping-Eintrag werden 1:1 als Business Unit übernommen. Bei
Gleichstand zweier Offices, die auf **dieselbe** BU zeigen, gilt diese BU –
sonst `Unbestimmt (Gleichstand: …)`.

### CSV-Spalten

| Spalte | Inhalt |
|---|---|
| `SiteUrl`, `SiteTitel`, `Vorlage`, `GroupId` | Stammdaten der Site |
| `OwnerRoh` | Owner-Wert wie im Tenant hinterlegt (z. B. Gruppen-Claim) |
| `OwnerQuelle` | Wie die Personen ermittelt wurden (M365-Gruppe, Security-Gruppe, Site-Admin, …) |
| `AnzahlOwner` / `Owners` | Aufgelöste Personen (dedupliziert) |
| `OfficeVerteilung` | Alle Offices mit Anzahl, z. B. `Zuerich (4) \| Basel (2)` |
| `OwnerOhneOffice` | Personen ohne gepflegtes Office-Attribut |
| `MehrheitsOffice` | Gewinner der Mehrheitsrechnung (bei Gleichstand alle Beteiligten) |
| `MehrheitAnteilProzent` | Anteil des Mehrheits-Office (Basis: Personen mit Office) |
| `Gleichstand` | `Ja`/`Nein` |
| `BusinessUnit` | Abgeleitete Business Unit |
| `Ersteller` | Site-Ersteller via Graph `createdBy` (Best Effort, siehe Hinweise) |
| `Hinweis` | Auflösungsprobleme, Fallbacks usw. |

### Troubleshooting

**Welche URL gehört in `-TenantAdminUrl`?**
Die URL des SharePoint **Admin Centers**: `https://<tenantname>-admin.sharepoint.com`.
Den Tenantnamen siehst du in jeder normalen SharePoint-URL – liegen deine Sites
unter `https://contoso.sharepoint.com/sites/...`, lautet die Admin-URL
`https://contoso-admin.sharepoint.com`. **Nicht** die Root-Site und nicht
`/admin` anhängen.

**`Register-PnPEntraIDAppForInteractiveLogin` wird nicht erkannt («Die Benennung
wurde nicht als Name eines Cmdlet erkannt»)**
Das Cmdlet existiert erst seit PnP.PowerShell 2.12, und PnP 2.x/3.x läuft nur
auf PowerShell 7 – in der klassischen «Windows PowerShell 5.1» (blaues Fenster)
fehlt es immer. Prüfen mit `$PSVersionTable.PSVersion` und
`Get-Module PnP.PowerShell -ListAvailable`. Lösung:

```powershell
# 1) PowerShell 7 installieren (einmalig):
winget install --id Microsoft.PowerShell --source winget
# 2) Neues Terminal "PowerShell 7" öffnen (pwsh, NICHT "Windows PowerShell"), dann:
Install-Module PnP.PowerShell -Scope CurrentUser -Force
Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnP-Reporting" -Tenant <tenant>.onmicrosoft.com
```

Ohne `winget`: MSI von https://aka.ms/powershell herunterladen. Wer kein
PowerShell 7 installieren kann/darf, erstellt die App-Registrierung manuell im
Entra Admin Center (siehe oben) – dafür braucht es kein PnP-Cmdlet.

**Ich habe nur PowerShell 7.0/7.1 – welche Module?**
Aktuelle Modulversionen brauchen neuere Hosts (PnP.PowerShell 2.x/3.x → PS 7.2/7.4+,
Microsoft.Graph 2.x → PS 7.2+). Auf PS 7.1 die letzten kompatiblen Versionen pinnen:

```powershell
Install-Module PnP.PowerShell                 -RequiredVersion 1.12.0 -Scope CurrentUser
Install-Module Microsoft.Graph.Authentication -RequiredVersion 1.28.0 -Scope CurrentUser
```

Das Skript läuft damit unverändert. Einschränkungen: In PnP 1.12 gibt es
`Register-PnPEntraIDAppForInteractiveLogin` nicht – die App-Registrierung manuell
im Entra Admin Center anlegen (siehe unten) und die AppId als `-ClientId`
übergeben. PowerShell 7.1 ist seit 2022 End-of-Life; mittelfristig auf 7.4+
aktualisieren (`winget install Microsoft.PowerShell` aktualisiert in-place).

**Login mit «PnP Management Shell» schlägt fehl (z. B. `AADSTS700016: Application
with identifier '31359c7f-...' was not found`)**
Die multi-tenant App «PnP Management Shell» wurde vom PnP-Team im September 2024
**gelöscht**. Ein früher erteilter Admin Consent im eigenen Tenant ändert daran
nichts – die App existiert nicht mehr. Seit PnP.PowerShell 2.12 ist deshalb eine
**eigene App-Registrierung** Pflicht (`-ClientId`):

```powershell
# Einmalig als Admin ausführen (erstellt die App und fragt den Admin Consent ab):
Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnP-Reporting" -Tenant contoso.onmicrosoft.com
# Ausgegebene AppId/ClientId notieren und dem Skript mitgeben:
.\Get-SPSiteBusinessUnit.ps1 -TenantAdminUrl https://contoso-admin.sharepoint.com -ClientId <AppId>
```

Alternativ manuell im Entra Admin Center: *App registrations → New registration*
(Single Tenant) → unter *Authentication* Plattform «Mobile and desktop
applications» mit Redirect URI `http://localhost` hinzufügen und «Allow public
client flows» aktivieren → unter *API permissions* die **delegierte**
SharePoint-Berechtigung `AllSites.FullControl` ergänzen und Admin Consent
erteilen → die *Application (client) ID* als `-ClientId` verwenden.

Voraussetzungen: Das Konto braucht das Recht, App-Registrierungen zu erstellen,
und für den Consent einen Admin (z. B. Global Administrator). Für
`Get-PnPTenantSite` muss der angemeldete Benutzer zusätzlich
SharePoint-Administrator sein. Bei alten PnP-Versionen (< 2.x auf Windows
PowerShell 5.1) zuerst auf PowerShell 7 + aktuelles `PnP.PowerShell` wechseln
(`Get-Module PnP.PowerShell -ListAvailable` zeigt die Version).

### Hinweise / Grenzen

- **Ersteller:** Graph liefert `createdBy` nicht für jede Site; zuverlässig ist
  der Ersteller nur über das Unified Audit Log (`Search-UnifiedAuditLog`,
  Aufbewahrung standardmässig 180 Tage).
- **DeepScan** ist deutlich langsamer (eine Verbindung pro Site), liefert aber
  auch Site Collection Admins und die klassischen SharePoint-Besitzergruppen.
- Benutzer und Gruppen werden **gecacht** – jede Person/Gruppe wird nur einmal
  via Graph abgefragt.
- Gesperrte Sites (`NoAccess`) erzeugen im DeepScan einen Eintrag in `Hinweis`
  statt das Skript abzubrechen.
