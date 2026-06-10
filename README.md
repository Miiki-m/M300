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
