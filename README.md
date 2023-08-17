# wartungstool
Simple GUI Anwendung zur Abfrage grundlegender Informationen auf Windows Servern

# quickcheck script
Knappe Zusammenfassung wichtiger Informationen und Event Logs Einträgen.
./quickcheck.ps1 -{tage} um einen anderen Zeitraum als die default 30 Tage zu wählen.

# todo:


# Changelog Wartungstool:
v0.8b2
-DFSR Replication Check für Event IDs 2212 und 4012

v0.8b1
-Snapshot Check (Erfordert Admin Rechte) für Hyper-V Umgebungen hinzugefügt. Prüft via Get-VMSnapshot ob Snapshots vorhanden sind und zusätzlich im Hyper-V Virtual Machines Ordner ob *.avhdx Dateien existieren.

v0.7b2
-Beta-Args behoben, Indikator im Window-Text hinzugefügt
-beta Zeigt alle Funktionen an - auch unbenutzbare!

v0.7b
-Self-Elevate prompt falls Tool ohne Adminrechte ausgeführt wurde
-wtlog mit Timestamp

v0.6b
-WSUS Sync Errors hinzugefügt
-Eventlog Hinzugefügt

v0.5b2
-last Update KB Nummer hinzugefügt.

v0.5b
-beta Funktionen durch beta Argument beim ausführen.
-Server Uptime auf erreichbare Server beschränkt.

v0.5
-Client Uptime hinzugefügt
-Formatierung etwas korrigiert
-Servername und Netzwerk in Topzeile eingefügt.


v0.4
-Server Uptime Abfrage von AD hinzugefügt. Noch nicht 100% auf aktive Server beschränkt.
-AD Check hinzugefügt.

v0.3
-WSUS Content Winver Check.
-Kommentare im Code
-Lastupdates
-Export to log
-WSUS Funktionen nur verfügbar, wenn WSUS installiert ist.

v0.2
-Initial
