# Kalender-Export & lokaler Outlook-Sync

Ein PowerShell-Skript zum automatisierten Auslesen von Kalenderdaten (Termine, Abwesenheiten und Sonderzeiten inkl. Detail-Modals und Änderungshistorie) aus dem PM-Webportal, Export als strukturierte JSON-Datei sowie Synchronisation eigener Einträge in das lokale Microsoft Outlook.

---

## ✨ Funktionen

* **Pre-Flight VPN-Check:** Prüft vor dem Verbindungsaufbau per DNS-Auflösung, ob der Ziel-Host auf eine interne VPN-IP (`10.x.x.x`) auflöst.
* **Single-Config & DPAPI-Verschlüsselung:** Speichert URL, Workspace-ID, Zugangsdaten und Filter in einer einzigen `credentials.txt`. Passwörter werden über die Windows Data Protection API (DPAPI) verschlüsselt und sind an den aktuellen Windows-Benutzer und Rechner gebunden.
* **Gleitende oder feste Zeiträume:** Unterstützt feste Monatsangaben (`2026-11`) sowie relative Monats-Offsets (`+0` bis `+1`), die jeden Monat automatisch mitwandern.
* **Flexible Benutzerfilter:** Abruf aller Mitarbeiter, nur der in PM vorausgewählten Benutzer oder einer individuell definierten Auswahl (inkl. lokalem User-Caching).
* **Tiefen-Extraktion (Detail-Modals):** Liest neben der Monatsübersicht alle Detail-Dialoge (`Mode=4`) für **Meetings** (inkl. Historie), **Abwesenheiten** und **Sonderzeiten** aus und berechnet exakte Start-/Endzeiten sowie Stundensummen.
* **JSON-Export:** Erzeugt pro abgerufenem Monat eine sauber strukturierte Datei (`PM_Kalender_YYYY-MM.json`).
* **Lokaler Outlook-Sync (COM):** Überträgt eigene Termine wahlweise in den Outlook-Hauptkalender oder einen eigenen Unterkalender (z. B. `PM`). Erkennt bestehende Einträge anhand einer benutzerdefinierten Eigenschaft (`PMKey`) und aktualisiert sie, statt Duplikate anzulegen.

---

## 📋 Voraussetzungen

* **Betriebssystem:** Windows 10 / 11
* **PowerShell:** Version 5.1 oder höher
* **Abhängigkeit:** `HtmlAgilityPack.dll` (muss im selben Verzeichnis wie das Skript liegen)
* **Netzwerk:** Aktive Verbindung zum Firmen-VPN (Ziel-Server muss auf `10.x.x.x` auflösen)
* **Optional:** Klassisches Microsoft Outlook (Desktop-Installation) für den lokalen Kalender-Sync

---

## 🚀 Installation & Ersteinrichtung

1. **Skript & DLL bereitstellen**  
   Lege das PowerShell-Skript (z. B. `Sync-PMCalendar.ps1`) in einem Ordner ab und kopiere die `HtmlAgilityPack.dll` (z. B. aus einem bestehenden .NET-Projekt oder dem [NuGet-Paket](https://www.nuget.org/packages/HtmlAgilityPack/)) in denselben Ordner:
   ```text
   PM-Calendar-Sync/
   ├── Sync-PMCalendar.ps1
   ├── HtmlAgilityPack.dll
   └── credentials.txt        # Wird bei der Ersteinrichtung automatisch erzeugt
   ```

2. **Skript ausführen**  
   Starte das Skript in PowerShell:
   ```powershell
   .\Sync-PMCalendar.ps1
   ```
   *(Falls die Ausführung von Skripten auf dem System eingeschränkt ist, einmalig `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` ausführen.)*

3. **Interaktiver Einrichtungsassistent**  
   Beim ersten Start fragt das Skript alle benötigten Parameter ab und speichert sie auf Wunsch in der `credentials.txt`:
   * Base-URL (z. B. `https://server.firma.de/XX/`)
   * Workspace-ID (Standardwert per `Enter` übernehmbar)
   * Benutzername & Passwort (wird maskiert eingelesen und DPAPI-verschlüsselt gespeichert)
   * Gewünschter Zeitraum (gleitend oder fest)
   * Benutzerfilter für den Abruf
   * Eigener PM-Benutzer & Ziel-Kalender in Outlook

---

## ⚙️ Konfiguration (`credentials.txt`)

Alle Einstellungen werden im `Key=Value`-Format in der `credentials.txt` gespeichert und können bei Bedarf manuell angepasst werden:

| Schlüssel | Beschreibung | Beispiel / Mögliche Werte |
| :--- | :--- | :--- |
| `BaseUrl` | Basis-URL der PM-Instanz | `https://server.firma.de/XX/` |
| `WorkspaceId` | GUID des PM-Arbeitsbereichs | `4DD936F6-33D2-4D56-AB08-B0D23FEE4B9D` |
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

* **Duplikat-Schutz:** Jeder Termin erhält in Outlook die versteckte Eigenschaft `PMKey` (`<UserId>_<DB_ID>`). Bei erneuten Durchläufen werden bestehende Termine aktualisiert statt neu angelegt.
* **Kategorien & Status:**
  * `MEETING` → Kategorie `PM; Termin` (Status: *Gebucht*, bei Absage *Frei* + Präfix `[Abgesagt]`, bei Tentativ *Mit Vorbehalt*)
  * `ABSENCE` → Kategorie `PM; Fehlzeit` (Status: *Abwesend*, ganztägig oder als 1./2. Halbtag)
  * `SPECIALTIME` → Kategorie `PM; Sonderzeit` (Status: *Gebucht*)
* **Inhalt:** Betreff, Ort (`Vor Ort`, `Remote` oder Kundenadresse) sowie Notizen inklusive Ticketnummer, Auftrag und vollständiger Änderungshistorie werden automatisch in den Outlook-Termin übernommen.

---
