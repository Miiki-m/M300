# Schritt-für-Schritt: SharePoint-Site → Business-Unit-Report

Zielbild: Der Report läuft **App-Only mit Zertifikat** im **GraphOnly-Modus**
(keine SharePoint-Berechtigung auf der App, nur read-only Graph-Permissions)
und wird **automatisch** über die Windows-Aufgabenplanung ausgeführt.

---

## Schritt 1 – PowerShell und Module vorbereiten

```powershell
# Version prüfen:
$PSVersionTable.PSVersion
```

**PowerShell 7.4+** (empfohlen):

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
```

**PowerShell 7.0/7.1** (ältere kompatible Version pinnen):

```powershell
Install-Module Microsoft.Graph.Authentication -RequiredVersion 1.28.0 -Scope CurrentUser
```

`PnP.PowerShell` wird im GraphOnly-Modus **nicht** benötigt.

---

## Schritt 2 – App-Registrierung fertigstellen (Entra Admin Center)

Die App existiert bereits. Jetzt prüfen/ergänzen unter
**Entra Admin Center → App registrations → deine App**:

1. **API permissions** → *Add a permission* → **Microsoft Graph** →
   **Application permissions** (nicht Delegated!) → diese drei hinzufügen:
   - `Sites.Read.All`
   - `Group.Read.All`
   - `User.Read.All`
2. **Grant admin consent for &lt;Tenant&gt;** klicken → alle drei müssen ein
   grünes Häkchen haben.
3. SharePoint-Berechtigungen werden **keine** benötigt – `Sites.FullControl.All`
   kann entfernt bleiben.
4. Notieren von der **Overview**-Seite:
   - *Application (client) ID* → wird `$ClientId`
   - Tenant: `<tenant>.onmicrosoft.com` (Entra → Overview → *Primary domain*)
     oder die *Directory (tenant) ID* → wird `$Tenant`

---

## Schritt 3 – Zertifikat erstellen (auf der Maschine, die den Report ausführt)

```powershell
# Self-signed Zertifikat, 3 Jahre gueltig (landet in CurrentUser\My):
$cert = New-SelfSignedCertificate -Subject "CN=SPSiteReport" `
    -CertStoreLocation "Cert:\CurrentUser\My" -KeyExportPolicy Exportable `
    -KeySpec Signature -KeyLength 2048 -NotAfter (Get-Date).AddYears(3)

# Oeffentlicher Teil (.cer) - wird im Portal hochgeladen:
Export-Certificate -Cert $cert -FilePath .\SPSiteReport.cer

# Privater Teil (.pfx) - Backup + spaeter Import in den Maschinenspeicher:
$pfxPwd = Read-Host -AsSecureString "PFX-Passwort waehlen"
Export-PfxCertificate -Cert $cert -FilePath .\SPSiteReport.pfx -Password $pfxPwd

# Thumbprint notieren -> wird $CertificateThumbprint:
$cert.Thumbprint
```

---

## Schritt 4 – Zertifikat in der App hochladen

**App registrations → deine App → Certificates & secrets → Certificates →
Upload certificate** → `SPSiteReport.cer` auswählen → *Add*.

Der angezeigte Thumbprint im Portal muss mit `$cert.Thumbprint` übereinstimmen.

---

## Schritt 5 – Skripte ablegen und ausfüllen

1. Ordner anlegen, z. B. `C:\Scripts\SPSiteReport`, und die drei Dateien
   hineinlegen:
   - `Get-SPSiteBusinessUnit.ps1` (Hauptskript)
   - `Run-SPSiteReport.ps1` (Runner für geplante Läufe)
   - `Register-SPSiteReportTask.ps1` (richtet die Aufgabe ein)
2. In **allen drei** Skripten oben im Block
   `>>> KONFIGURATION: DIESE WERTE AUSFUELLEN <<<` eintragen:

   | Variable | Wert |
   |---|---|
   | `$ClientId` | Application (client) ID aus Schritt 2 |
   | `$Tenant` | `<tenant>.onmicrosoft.com` |
   | `$CertificateThumbprint` | Thumbprint aus Schritt 3 |
   | `$GraphOnly` | `$true` |
   | `$TenantAdminUrl` | darf leer bleiben (nur ohne GraphOnly nötig) |

3. Die Mapping-Tabelle `$OfficeToBusinessUnit` im Hauptskript vorerst leer
   lassen – sie wird in Schritt 7 anhand echter Daten gefüllt.

---

## Schritt 6 – Manueller Testlauf

```powershell
cd C:\Scripts\SPSiteReport
.\Get-SPSiteBusinessUnit.ps1 -Limit 10
```

Erwartung: `Verbinde mit Microsoft Graph...` → `n Sites werden analysiert` →
CSV `SPSite-BusinessUnit-Report.csv` entsteht. In Excel öffnen und prüfen:

- `Owners` enthält echte Personen (nicht mehr «Owner-Gruppe von Site xyz»)
- `OfficeVerteilung` zeigt die Standorte mit Anzahl, z. B. `Zuerich (3) | Basel (1)`
- `OwnerOhneOffice` > 0 bedeutet: bei diesen Personen ist
  `physicalDeliveryOfficeName` im AD/Entra nicht gepflegt

---

## Schritt 7 – Business-Unit-Mapping pflegen und Volllauf

1. Aus der Spalte `OfficeVerteilung`/`MehrheitsOffice` die vorkommenden
   Office-Namen ablesen (exakte Schreibweise!).
2. Im Hauptskript die Tabelle füllen:

   ```powershell
   $OfficeToBusinessUnit = @{
       'Zuerich HQ' = 'BU Corporate'
       'Basel'      = 'BU Pharma'
       'Geneve'     = 'BU Finance'
   }
   ```

3. Volllauf ohne Limit: `.\Get-SPSiteBusinessUnit.ps1`
4. Nacharbeiten: Zeilen mit `Gleichstand = Ja` oder
   `BusinessUnit = Unbestimmt (...)` manuell zuordnen bzw. Office-Attribute
   im AD/Entra nachpflegen.

---

## Schritt 8 – Automatisierung vorbereiten (einmalig, als Administrator)

```powershell
# 1) Module SYSTEMWEIT installieren - das Task-Konto (SYSTEM) sieht
#    CurrentUser-Installationen NICHT:
Install-Module Microsoft.Graph.Authentication -Scope AllUsers
# (bei PS 7.0/7.1: zusaetzlich -RequiredVersion 1.28.0)

# 2) Zertifikat in den Maschinenspeicher importieren:
Import-PfxCertificate -FilePath .\SPSiteReport.pfx `
    -CertStoreLocation Cert:\LocalMachine\My `
    -Password (Read-Host -AsSecureString "PFX-Passwort")
```

Falls die Aufgabe später unter einem **Dienstkonto** (statt SYSTEM) laufen
soll: `certlm.msc` → Zertifikat → *Alle Aufgaben* → *Private Schlüssel
verwalten* → dem Konto Lesezugriff geben. Das Konto braucht zudem das Recht
«Anmelden als Stapelverarbeitungsauftrag».

---

## Schritt 9 – Geplante Aufgabe registrieren (als Administrator)

```powershell
cd C:\Scripts\SPSiteReport
.\Register-SPSiteReportTask.ps1            # Werte stehen oben im Skript
```

Standard: wöchentlich montags 06:00 als SYSTEM. Varianten:

```powershell
.\Register-SPSiteReportTask.ps1 -Frequency Daily -Time 05:30
.\Register-SPSiteReportTask.ps1 -DaysOfWeek Monday,Thursday -Time 07:00
.\Register-SPSiteReportTask.ps1 -ServiceAccount 'DOMAIN\svc-spreport'
```

---

## Schritt 10 – Automatischen Lauf testen

```powershell
Start-ScheduledTask -TaskName 'SPSite-BusinessUnit-Report'

# kurz warten (Laufzeit haengt von der Anzahl Sites ab), dann:
(Get-ScheduledTaskInfo -TaskName 'SPSite-BusinessUnit-Report').LastTaskResult   # 0 = OK

# Ergebnis ansehen:
Get-ChildItem C:\Reports\SPSiteReport
```

Pro Lauf entstehen `SPSite-BU-Report_<Zeitstempel>.csv` und
`Transcript_<Zeitstempel>.log`; Dateien älter als 90 Tage werden automatisch
gelöscht. Bei `LastTaskResult` ≠ 0 das neueste Transcript-Log lesen – dort
steht die genaue Fehlermeldung.

---

## Häufige Fehler

| Fehler | Ursache / Lösung |
|---|---|
| `AADSTS700016: Application ... was not found` | `$ClientId` oder `$Tenant` falsch |
| `AADSTS700027: ... invalid client assertion / certificate` | Hochgeladenes `.cer` passt nicht zum verwendeten Zertifikat – Thumbprint vergleichen |
| `Authorization_RequestDenied` / `Access denied` bei Graph | Application-Permissions fehlen oder **Admin Consent nicht erteilt** (Schritt 2) |
| Task schlägt fehl, manuell klappt es | Module nicht mit `-Scope AllUsers` installiert oder Zertifikat nicht in `Cert:\LocalMachine\My` (Schritt 8) |
| `Certificate with thumbprint ... not found` | Zertifikat liegt nicht im Store des ausführenden Kontos – Schritt 8 Punkt 2 |
| Viele `OwnerOhneOffice` | `physicalDeliveryOfficeName` im AD/Entra nicht gepflegt – Datenqualität, kein Skriptfehler |
