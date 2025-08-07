# Server Maintenance Tools

This repository contains PowerShell tools for system administration and maintenance on Windows Servers.

<a href="#english">English</a> | <a href="#deutsch">Deutsch</a>

---

## <a name="english"></a>English

### Maintenance Tool (`Wartungs Tool.ps1`)

A PowerShell script with a graphical user interface (GUI) for performing various server maintenance and information-gathering tasks.

**Features:**
- **System Information:** Check free disk space (`C:`), server uptime.
- **Active Directory:** Get a list of all AD computers, check the uptime of servers and clients in AD.
- **WSUS (Windows Server Update Services):**
  - Check the size of the WSUS content folder.
  - View WSUS synchronization errors and general errors.
  - Shrink the WSUS content by declining superseded updates and cleaning up unneeded files.
- **Hyper-V:** Check for existing VM snapshots and `.avhdx` files.
- **DFSR:** Check for DFSR replication errors (Event IDs 2212, 4012).
- **Event Logs:** Get a detailed overview of Application and System event logs (Errors and Warnings).
- **Updates:** List recently installed Windows updates.
- **Export:** Export the output to a timestamped log file (`wtlog_...txt`).
- **Beta Mode:** Run the script with the `-beta` argument to access experimental features.

**Prerequisites:**
- Windows Server with PowerShell.
- The script must be run with **Administrator privileges**. It includes a self-elevation prompt if not run as admin.
- Some features require specific Windows Roles to be installed (e.g., WSUS, Active Directory Domain Services, Hyper-V). The script will automatically show/hide buttons based on installed roles.

**Usage:**
```powershell
.\Wartungs Tool.ps1
```
To run with experimental features enabled:
```powershell
.\Wartungs Tool.ps1 -beta
```

---

### Quick Check Script (`quickcheck.ps1`)

A command-line script that provides a quick summary of the server's status.

**Features:**
- Displays current server uptime.
- Shows storage information for all logical disks (free space and total size).
- Reports the most recently installed Windows update.
- Summarizes Error and Warning events from the Application and System logs for a specified number of days.

**Usage:**
Run the script from a PowerShell console.
```powershell
.\quickcheck.ps1
```
To specify the number of days to look back for event logs (default is 30):
```powershell
.\quickcheck.ps1 -<days>
# Example for the last 10 days:
.\quickcheck.ps1 -10
```

---

### ToDo
- Refactor code: Migrate all checks into a function-based format, similar to the `Eventoverview` function.

---

<details>
<summary><b>Changelog</b></summary>

**v0.9.1.1**
- (Extended) Eventoverview: The total number of events is now also displayed.

**v0.9.1**
- (New) Eventoverview: Provides an overview (sorted by frequency) of Application & System Events (Error and Warning) from the last 30 days.

**v0.9.0.1**
- Fixed typos.
- Adjusted text outputs.

**v0.9b1**
- Shrink WSUS Content: Declines superseded updates and performs a cleanup of unneeded content files.

**v0.8b2**
- DFSR Replication Check for Event IDs 2212 and 4012.

**v0.8b1**
- Snapshot Check (Requires Admin rights) for Hyper-V environments added.

**v0.7b2**
- Fixed Beta-Args and added an indicator in the window title for beta mode.
- `-beta` now shows all functions, including non-usable ones.

**v0.7b**
- Self-elevate prompt if the tool is run without admin rights.
- `wtlog` is now timestamped.

**v0.6b**
- Added WSUS Sync Errors.
- Added Eventlog check.

**v0.5b2**
- Added last Update KB number.

**v0.5b**
- Beta functions are enabled via the `-beta` argument.
- Server uptime check is now limited to reachable servers.

**v0.5**
- Added client uptime check.
- Minor formatting corrections.
- Added server name and network to the top line of the GUI.

**v0.4**
- Added server uptime query from AD.
- Added AD Check.

**v0.3**
- WSUS Content Winver Check.
- Added comments to the code.
- Added "Last Updates" check.
- Added "Export to log" feature.
- WSUS functions are now only available if WSUS is installed.

**v0.2**
- Initial version.
</details>

<br>

---

## <a name="deutsch"></a>Deutsch

### Wartungstool (`Wartungs Tool.ps1`)

Ein PowerShell-Skript mit einer grafischen Benutzeroberfläche (GUI) zur Durchführung verschiedener Server-Wartungs- und Informationsabfrageaufgaben.

**Funktionen:**
- **Systeminformationen:** Überprüfung des freien Speicherplatzes (`C:`), Server-Uptime.
- **Active Directory:** Abrufen einer Liste aller AD-Computer, Überprüfung der Uptime von Servern und Clients im AD.
- **WSUS (Windows Server Update Services):**
  - Überprüfung der Größe des WSUS-Content-Ordners.
  - Anzeige von WSUS-Synchronisationsfehlern und allgemeinen Fehlern.
  - Verkleinern des WSUS-Contents durch Ablehnen veralteter Updates und Bereinigen nicht mehr benötigter Dateien.
- **Hyper-V:** Suche nach vorhandenen VM-Snapshots und `.avhdx`-Dateien.
- **DFSR:** Prüfung auf DFSR-Replikationsfehler (Ereignis-IDs 2212, 4012).
- **Ereignisprotokolle:** Detaillierte Übersicht über Anwendungs- und Systemereignisprotokolle (Fehler und Warnungen).
- **Updates:** Auflisten der zuletzt installierten Windows-Updates.
- **Export:** Exportieren der Ausgabe in eine Log-Datei mit Zeitstempel (`wtlog_...txt`).
- **Beta-Modus:** Ausführen des Skripts mit dem `-beta`-Argument, um auf experimentelle Funktionen zuzugreifen.

**Voraussetzungen:**
- Windows Server mit PowerShell.
- Das Skript muss mit **Administratorrechten** ausgeführt werden. Es enthält eine Abfrage zur Rechteerweiterung, falls es nicht als Administrator gestartet wird.
- Einige Funktionen erfordern die Installation bestimmter Windows-Rollen (z. B. WSUS, Active Directory Domain Services, Hyper-V). Das Skript zeigt die entsprechenden Schaltflächen automatisch an oder verbirgt sie, je nach installierten Rollen.

**Anwendung:**
```powershell
.\Wartungs Tool.ps1
```
Um das Skript mit aktivierten experimentellen Funktionen auszuführen:
```powershell
.\Wartungs Tool.ps1 -beta
```

---

### Quickcheck Skript (`quickcheck.ps1`)

Ein Kommandozeilen-Skript, das eine schnelle Zusammenfassung des Serverstatus liefert.

**Funktionen:**
- Zeigt die aktuelle Server-Uptime an.
- Zeigt Speicherinformationen für alle logischen Laufwerke (freier Speicher und Gesamtgröße).
- Meldet das zuletzt installierte Windows-Update.
- Fasst Fehler- und Warnungsereignisse aus den Anwendungs- und Systemprotokollen für eine bestimmte Anzahl von Tagen zusammen.

**Anwendung:**
Führen Sie das Skript in einer PowerShell-Konsole aus.
```powershell
.\quickcheck.ps1
```
Um die Anzahl der Tage für die Rückschau der Ereignisprotokolle festzulegen (Standard ist 30):
```powershell
.\quickcheck.ps1 -<Tage>
# Beispiel für die letzten 10 Tage:
.\quickcheck.ps1 -10
```

---

### ToDo
- Code-Refactoring: Alle Prüfungen in ein funktionsbasiertes Format überführen, ähnlich der `Eventoverview`-Funktion.

---

<details>
<summary><b>Changelog</b></summary>

**v0.9.1.1**
- (Erweitert) Eventoverview: Die Gesamtzahl der Events wird nun auch angezeigt.

**v0.9.1**
- (Neu) Eventoverview: Gibt einen Überblick (sortiert nach Häufigkeit) der Application & System Events (Error und Warning) der letzten 30 Tage.

**v0.9.0.1**
- Tippfehler behoben.
- Textausgaben angepasst.

**v0.9b1**
- Shrink WSUS Content: Lehnt veraltete (superseded) Updates ab und führt eine Bereinigung der nicht mehr benötigten Content-Dateien durch.

**v0.8b2**
- DFSR Replication Check für die Event IDs 2212 und 4012.

**v0.8b1**
- Snapshot Check (erfordert Admin-Rechte) für Hyper-V-Umgebungen hinzugefügt.

**v0.7b2**
- Beta-Args behoben und einen Indikator im Fenstertitel für den Beta-Modus hinzugefügt.
- `-beta` zeigt nun alle Funktionen an, auch die nicht nutzbaren.

**v0.7b**
- Self-Elevate-Prompt, falls das Tool ohne Admin-Rechte ausgeführt wurde.
- `wtlog` ist jetzt mit einem Zeitstempel versehen.

**v0.6b**
- WSUS Sync Errors hinzugefügt.
- Eventlog-Prüfung hinzugefügt.

**v0.5b2**
- Letzte Update-KB-Nummer hinzugefügt.

**v0.5b**
- Beta-Funktionen werden über das `-beta`-Argument aktiviert.
- Server-Uptime-Prüfung auf erreichbare Server beschränkt.

**v0.5**
- Client-Uptime-Prüfung hinzugefügt.
- Kleinere Formatierungskorrekturen.
- Servername und Netzwerk in die oberste Zeile der GUI eingefügt.

**v0.4**
- Server-Uptime-Abfrage von AD hinzugefügt.
- AD Check hinzugefügt.

**v0.3**
- WSUS Content Winver Check.
- Kommentare im Code hinzugefügt.
- "Last Updates"-Prüfung hinzugefügt.
- "Export to log"-Funktion hinzugefügt.
- WSUS-Funktionen sind nur verfügbar, wenn WSUS installiert ist.

**v0.2**
- Initiale Version.
</details>
