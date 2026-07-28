# NinjaOne Veeam Backup Check

PowerShell-Script zur Ueberwachung von Veeam Backup & Replication ueber NinjaOne.
Es liest den Status aller aktivierten Jobs aus, schreibt eine Zusammenfassung in
benutzerdefinierte Felder und liefert dabei die **Klartext-Ursache** aus dem
Veeam-Log statt nur "Warning" oder "Failed".

Autor: Sebastian Herrmann

## Was wird geprueft

| # | Bereich | Cmdlets |
|---|---------|---------|
| 1 | VM Backups | `Get-VBRJob`, `Get-VBRBackupSession` |
| 2 | Platform Backups (Proxmox VE, Nutanix AHV, oVirt/RHV/OLVM, Scale, Morpheus) | `Get-VBRSession -Type PlatformBackupJob`, Fallback `[Veeam.Backup.Core.CBackupJob]` |
| 3 | Backup Copy Jobs | `Get-VBRBackupCopyJob` |
| 4 | Agent Backups (Windows/Linux/Mac) | `Get-VBRComputerBackupJob` |
| 5 | Tape/Band Jobs | `Get-VBRTapeJob`, `Get-VBRTapeBackupSession` |
| 6 | Microsoft 365 (VBO) | `Get-VBOJobSession` |

## Besonderheiten

### Proxmox VE und andere Platform-Jobs

Veeam liefert Proxmox-VE-Jobs **nicht** ueber `Get-VBRJob` zurueck. Skripte, die
ihre Sessions ueber `Get-VBRJob -Name` gegenpruefen, filtern diese Jobs deshalb
unbemerkt heraus. Das Script nutzt daher einen eigenen Abschnitt mit zwei Wegen:

* **Weg A** (offiziell, ab VBR 12.3): `Get-VBRSession -Type PlatformBackupJob`,
  danach die vollstaendige Session per `Get-VBRBackupSession -Id` nachladen.
* **Weg B** (Fallback): `[Veeam.Backup.Core.CBackupJob]::GetAll()`, gefiltert
  ueber `TypeToString`.

### Klartext-Meldungen aus dem Log

Statt nur des Job-Ergebnisses werden die tatsaechlichen Warn- und Fehlerdetails
ausgelesen (`Logger.GetLog().GetAttentionRecords()` mit `Title` + `Description`,
inkl. mehrerer Fallbacks). Nichtssagende Standardzeilen wie
"Job finished with warning" werden nach hinten sortiert:

```
Tape-Job: Wochensicherung | Status: Warning | Zeitpunkt: 27.07.2026 22:00:00
    > No suitable media found :: Tape is not inserted into the drive
    > Job finished with warning
```

Das gilt fuer alle Ebenen: Job-Session, einzelne VMs, Rechner und Tape-Objekte.

### Tape-Jobs, die auf ein Band warten

Sessions im Zustand `WaitingTape` / `Idle` / `Pending` haben den Status
"Stopped" nie erreicht und sind in einfachen Abfragen unsichtbar. Sie werden
erkannt und als Warnung gemeldet, inklusive der letzten Log-Eintraege.

## Voraussetzungen

* Ausfuehrung auf dem **Veeam Backup & Replication Server**
* Veeam V10 bis V13
  * V11/V12: `Veeam.Backup.PowerShell` Modul oder `VeeamPSSnapin`
  * V13: zusaetzlich **PowerShell 7**, das Script startet sich dort selbst neu
* Ausfuehrung als **System** bzw. mit Rechten auf die Veeam-Konsole

## Benutzerdefinierte Felder in NinjaOne

Vor dem ersten Lauf anzulegen (Scope: Device):

| Feldname | Typ | Inhalt |
|----------|-----|--------|
| `veeamlog` | Multiline Text | Log, max. 10.000 Zeichen |
| `veeamsuccess` | Integer | 0 / 1 |
| `veeamwarning` | Integer | 0 / 1 |
| `veeamfail` | Integer | 0 / 1 |

Das Script muss auf die Felder Schreibrechte haben.

## Einrichtung

1. Die vier benutzerdefinierten Felder anlegen.
2. `NinjaOne Veeam Backup Script.ps1` als Script in NinjaOne importieren
   (Sprache: PowerShell, Ausfuehrung als System).
3. Einen Zeitplan setzen, der nach dem Backup-Fenster liegt.
4. Conditions auf `veeamfail = 1` bzw. `veeamwarning = 1` aufbauen.

## Bekannte Einschraenkungen

* **Geraet offline:** Ist die Maschine zum Ausfuehrungszeitpunkt aus, laeuft das
  Script gar nicht. Die Felder behalten ihre alten Werte, ein Ausfall kann im
  Dashboard also weiterhin gruen aussehen. Dagegen hilft nur eine separate
  Condition auf "Device Offline" bzw. "Last Contact".
* **`[Veeam.Backup.Core.CBackupJob]`** ist eine interne Klasse und von Veeam
  ausdruecklich **nicht supportet** ("only use this for reporting"). Sie kann mit
  einem Update wegbrechen. Der Aufruf ist gekapselt: faellt er aus, fehlt nur der
  Platform-Abschnitt, der Rest laeuft weiter. Veeam hat offizielle
  PowerShell-Cmdlets fuer Proxmox angekuendigt.
* **10.000-Zeichen-Limit:** Reicht der Platz nicht, werden zuerst die Zeilen
  erfolgreicher VMs ausgeblendet. Fehlermeldungen bleiben erhalten.

## Codierung

Die Datei ist bewusst reines ASCII (`ae`/`oe`/`ue` statt Umlaute), damit sie beim
Copy/Paste in den NinjaOne-Script-Editor nicht beschaedigt wird.
