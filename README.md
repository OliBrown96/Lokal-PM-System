# Kalender-Export, Web-Viewer & lokaler Outlook-Sync

Ein PowerShell-Skript zum automatisierten Auslesen von Kalenderdaten (Termine, Abwesenheiten und Sonderzeiten inkl. Detail-Modals und Änderungshistorie) aus dem PM-Webportal, Export als strukturierte JSON-Datei, Darstellung in einer interaktiven Browser-Übersicht (`index.html`) sowie Synchronisation eigener Einträge in das lokale Microsoft Outlook.

---

## ✨ Funktionen

- **Pre-Flight VPN-Check:** Prüft vor dem Verbindungsaufbau per DNS-Auflösung, ob der Ziel-Host auf eine interne VPN-IP (`10.x.x.x`) auflöst.
- **Single-Config & DPAPI-Verschlüsselung:** Speichert URL, Workspace-ID, Zugangsdaten und Filter in einer einzigen `credentials.txt`. Passwörter werden über die Windows Data Protection API (DPAPI) verschlüsselt und sind an den aktuellen Windows-Benutzer und Rechner gebunden.
- **Gleitende oder feste Zeiträume:** Unterstützt feste Monatsangaben (`2026-11`) sowie relative Monats-Offsets (`+0` bis `+1`), die jeden Monat automatisch mitwandern.
- **Flexible Benutzerfilter:** Abruf aller Mitarbeiter, nur der in PM vorausgewählten Benutzer oder einer individuell definierten Auswahl (inkl. lokalem User-Caching).
- **Tiefen-Extraktion (Detail-Modals):** Liest neben der Monatsübersicht alle Detail-Dialoge (`Mode=4`) für **Meetings** (inkl. Historie), **Abwesenheiten** und **Sonderzeiten** aus und berechnet exakte Start-/Endzeiten sowie Stundensummen.
- **JSON-Export:** Erzeugt pro abgerufenem Monat eine sauber strukturierte Datei (`PM_Kalender_YYYY-MM.json`).
- **Interaktive Teamübersicht (`index.html`):** Visualisiert die exportierte JSON-Datei direkt im Browser als übersichtliche Monatsmatrix pro Mitarbeiter – inklusive Echtzeit-Suche, Kategoriefilter, Stundenauswertung und Detail-Popup für jeden Eintrag.
- **Lokaler Outlook-Sync (COM):** Überträgt eigene Termine wahlweise in den Outlook-Hauptkalender oder einen eigenen Unterkalender (z. B. `PM`). Erkennt bestehende Einträge anhand einer benutzerdefinierten Eigenschaft (`PMKey`) und aktualisiert sie, statt Duplikate anzulegen.

---

## 📋 Voraussetzungen

- **Betriebssystem:** Windows 10 / 11
- **PowerShell:** Version 5.1 oder höher
- **Abhängigkeit:** `HtmlAgilityPack.dll` (muss im selben Verzeichnis wie das Skript liegen)
- **Netzwerk:** Aktive Verbindung zum Firmen-VPN (Ziel-Server muss auf `10.x.x.x` auflösen)
- **Optional:**
  - Klassisches Microsoft Outlook (Desktop-Installation) für den lokalen Kalender-Sync
  - Moderner Webbrowser (Edge, Chrome, Firefox) für die Teamübersicht (`index.html`)

---

## 🚀 Installation & Ersteinrichtung

1. **Skript, Web-Viewer & DLL bereitstellen**  
   Lege das PowerShell-Skript (`Sync-PMCalendar.ps1`) und die `index.html` in einem Ordner ab und besorge die `HtmlAgilityPack.dll` über das offizielle [NuGet-Paket](https://www.nuget.org/packages/HtmlAgilityPack/) (*Download package* → `.nupkg` in `.zip` umbenennen → Datei aus `lib/netstandard2.0/HtmlAgilityPack.dll` entpacken).

   Alternativ kannst du die DLL direkt per PowerShell in den Skript-Ordner herunterladen:
   ```powershell
   Invoke-WebRequest "https://www.nuget.org/api/v2/package/HtmlAgilityPack" -OutFile "$env:TEMP\hap.zip"
   Expand-Archive "$env:TEMP\hap.zip" -DestinationPath "$env:TEMP\hap" -Force
   Copy-Item "$env:TEMP\hap\lib\netstandard2.0\HtmlAgilityPack.dll" -Destination .
   Remove-Item "$env:TEMP\hap.zip", "$env:TEMP\hap" -Recurse -Force
   ```

   > **Hinweis:** Falls Windows die heruntergeladene DLL blockiert, einmalig `Unblock-File -Path .\HtmlAgilityPack.dll` ausführen.

   Die Ordnerstruktur sieht anschließend so aus:
   ```text
   PM-Calendar-Sync/
   ├── Sync-PMCalendar.ps1
   ├── index.html
   ├── HtmlAgilityPack.dll
   ├── credentials.txt              # Wird bei der Ersteinrichtung automatisch erzeugt
   └── PM_Kalender_YYYY-MM.json     # Wird beim Export automatisch erzeugt
   ```

2. **Skript ausführen**  
   Starte das Skript in PowerShell:
   ```powershell
   .\Sync-PMCalendar.ps1
   ```
   *(Falls die Ausführung von Skripten auf dem System eingeschränkt ist, einmalig `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` ausführen.)*

3. **Interaktiver Einrichtungsassistent**  
   Beim ersten Start fragt das Skript alle benötigten Parameter ab und speichert sie auf Wunsch in der `credentials.txt`:
   - Base-URL (z. B. `https://server.firma.de/XX/`)
   - Workspace-ID (Standardwert per `Enter` übernehmbar)
   - Benutzername & Passwort (wird maskiert eingelesen und DPAPI-verschlüsselt gespeichert)
   - Gewünschter Zeitraum (gleitend oder fest)
   - Benutzerfilter für den Abruf
   - Eigener PM-Benutzer & Ziel-Kalender in Outlook

---

## 🌐 Nutzung des Web-Viewers (`index.html`)

Die `index.html` benötigt keinen Webserver und läuft vollständig lokal im Browser:

1. Öffne die `index.html` per Doppelklick im Browser.
2. Wähle über den Button **„JSON-Datei laden“** (oder per **Drag & Drop** ins Fenster) eine der exportierten Monatsdateien (`PM_Kalender_YYYY-MM.json`) aus. *(Beim lokalen Öffnen über `file://` blockieren Browser aus Sicherheitsgründen das automatische Nachladen lokaler Dateien.)*
3. **Funktionen in der Ansicht:**
   - **Monatsmatrix:** Zeigt alle Mitarbeiter zeilenweise mit Wochenend-Hervorhebung, Terminanzahl und Stundensumme.
   - **Filter & Suche:** Filtert live nach Mitarbeiternamen, Kunden, Beschreibungen oder Terminkategorien.
   - **Detail-Modal:** Ein Klick auf einen Termin öffnet alle erfassten Einsatz-, Abwesenheits- oder Sonderzeitdetails inklusive Änderungshistorie und aufklappbaren JSON-Rohdaten.

---

## ⚙️ Konfiguration (`credentials.txt`)

Alle Einstellungen werden im `Key=Value`-Format in der `credentials.txt` gespeichert und können bei Bedarf manuell angepasst werden:

| Schlüssel | Beschreibung | Beispiel / Mögliche Werte |
| :--- | :--- | :--- |
| `BaseUrl` | Basis-URL der PM-Instanz | `https://server.firma.de/XX/` |
| `WorkspaceId` | GUID des PM-Arbeitsbereichs | `00000000-0000-0000-0000-000000000000` |
| `Username` | PM-Anmeldename | `max.mustermann` |
| `EncryptedPassword` | DPAPI-verschlüsseltes Passwort | `01000000d08c9dPM...` |
| `StartMonth` | Startmonat (fest oder relativ) | `+0` oder `2026-11` |
| `EndMonth` | Endmonat (fest oder relativ) | `+1` oder `2026-12` |
| `UserFilter` | Filter für abzurufende Kalender | `ALL`, `SELECTED` oder kommagetrennte User-IDs |
| `OwnUserId` | Eigene PM-User-ID für den Outlook-Sync | `12345` |
| `OutlookCalendar` | Ziel-Kalender in Outlook | `DEFAULT`, `PM` (Unterkalender) oder `NONE` (deaktiviert) |
| `CachedUsers` | JSON-Cache der verfügbaren PM-Benutzer | `[{"userId":"...","displayName":"..."}]` |

> **Tipp:** Um eine Abfrage (z. B. Zeitraum oder Benutzerfilter) erneut interaktiv auszulösen, einfach die entsprechende Zeile aus der `credentials.txt` entfernen.

---

## 📅 Funktionsweise des Outlook-Syncs

- **Duplikat-Schutz:** Jeder Termin erhält in Outlook die versteckte Eigenschaft `PMKey` (`<UserId>_<DB_ID>`). Bei erneuten Durchläufen werden bestehende Termine aktualisiert statt neu angelegt.
- **Kategorien & Status:**
  - `MEETING` → Kategorie `PM; Termin` (Status: *Gebucht*, bei Absage *Frei* + Präfix `[Abgesagt]`, bei Tentativ *Mit Vorbehalt*)
  - `ABSENCE` → Kategorie `PM; Fehlzeit` (Status: *Abwesend*, ganztägig oder als 1./2. Halbtag)
  - `SPECIALTIME` → Kategorie `PM; Sonderzeit` (Status: *Gebucht*)
- **Inhalt:** Betreff, Ort (`Vor Ort`, `Remote` oder Kundenadresse) sowie Notizen inklusive Ticketnummer, Auftrag und vollständiger Änderungshistorie werden automatisch in den Outlook-Termin übernommen.

---

## 📄 Lizenz & Drittanbieter

- Dieses Projekt steht unter der [GNU General Public License v3.0 (GPLv3)](LICENSE).
- Nutzt [HtmlAgilityPack](https://html-agility-pack.net/) (lizenziert unter der MIT-Lizenz, © ZZZ Projects).
