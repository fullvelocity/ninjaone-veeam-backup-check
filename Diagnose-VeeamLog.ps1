# =============================================================================
# DIAGNOSE-SCRIPT FUER DEN NINJAONE VEEAM CHECK
# =============================================================================
#
# Autor: Sebastian Herrmann
#
# ZWECK
# Dieses Script aendert NICHTS. Es liest nur aus und gibt auf der Konsole aus,
# welche Objekte und Eigenschaften die installierte Veeam-Version tatsaechlich
# liefert. Damit laesst sich klaeren:
#
#   1. Woher die Warn-/Fehlermeldung eines Tape-Jobs wirklich kommt
#      (.Log, .Logger.GetLog(), GetDetails() oder Info.Reason)
#   2. Woran sich interne Kind-Jobs erkennen lassen, die Veeam pro VM anlegt
#      und die sonst das Log zumuellen
#
# AUSFUEHRUNG
#   Direkt auf dem Veeam Backup & Replication Server, als Administrator:
#     powershell -ExecutionPolicy Bypass -File .\Diagnose-VeeamLog.ps1
#   Bei Veeam V13 stattdessen mit PowerShell 7:
#     pwsh -File .\Diagnose-VeeamLog.ps1
#
#   Ausgabe komplett kopieren. Enthalten sind nur Job-/VM-Namen und
#   Veeam-Meldungstexte - keine Passwoerter, keine Pfade zu Backupdaten.
#
# =============================================================================

$WarningPreference  = 'SilentlyContinue'
$ProgressPreference = 'SilentlyContinue'

function Write-Head {
    param([string]$Text)
    Write-Host ""
    Write-Host "=============================================================" -ForegroundColor Cyan
    Write-Host $Text -ForegroundColor Cyan
    Write-Host "=============================================================" -ForegroundColor Cyan
}

function Write-Sub {
    param([string]$Text)
    Write-Host ""
    Write-Host "--- $Text " -ForegroundColor Yellow
}

# Kuerzt lange Texte fuer die Konsolenausgabe
function Short {
    param([string]$Text, [int]$Max = 200)
    if ([string]::IsNullOrEmpty($Text)) { return "<leer>" }
    $t = ($Text -replace '<[^>]+>', ' ') -replace '\s+', ' '
    $t = $t.Trim()
    if ($t.Length -gt $Max) { $t = $t.Substring(0, $Max) + " [...]" }
    if ([string]::IsNullOrEmpty($t)) { return "<leer>" }
    return $t
}

# -----------------------------------------------------------------------------
# 0. UMGEBUNG
# -----------------------------------------------------------------------------
Write-Head "0. UMGEBUNG"
Write-Host "PowerShell-Version : $($PSVersionTable.PSVersion)"
Write-Host "PowerShell-Edition : $($PSVersionTable.PSEdition)"

if (-not (Get-Command Get-VBRJob -ErrorAction SilentlyContinue)) {
    try {
        if (Get-Module -ListAvailable -Name "Veeam.Backup.PowerShell") {
            Import-Module "Veeam.Backup.PowerShell" -DisableNameChecking -ErrorAction Stop | Out-Null
        }
    } catch {}
    if (-not (Get-Command Get-VBRJob -ErrorAction SilentlyContinue)) {
        try { Add-PSSnapin VeeamPSSnapin -ErrorAction SilentlyContinue | Out-Null } catch {}
    }
}

if (-not (Get-Command Get-VBRJob -ErrorAction SilentlyContinue)) {
    Write-Host "ABBRUCH: Veeam PowerShell steht in dieser Sitzung nicht zur Verfuegung." -ForegroundColor Red
    Write-Host "Bei Veeam V13 dieses Script mit pwsh.exe starten, nicht mit powershell.exe."
    return
}

# Laedt die interne Core-Assembly (siehe Hauptscript)
try { Get-VBRLicenseAutoUpdateStatus -ErrorAction SilentlyContinue | Out-Null } catch {}

try {
    $v = (Get-Module Veeam.Backup.PowerShell -ErrorAction SilentlyContinue).Version
    if ($v) { Write-Host "Veeam-Modulversion : $v" }
} catch {}

$coreOk = $false
try {
    $null = [Veeam.Backup.Core.CBackupJob]
    $coreOk = $true
} catch {}
Write-Host "Zugriff auf Veeam.Backup.Core : $coreOk"

# -----------------------------------------------------------------------------
# 1. TAPE-JOBS: WO STECKT DIE MELDUNG?
# -----------------------------------------------------------------------------
Write-Head "1. TAPE-JOBS - QUELLE DER WARN-/FEHLERMELDUNG"

$tapeJobs = @()
try { $tapeJobs = @(Get-VBRTapeJob -ErrorAction SilentlyContinue) } catch {}

if ($tapeJobs.Count -eq 0) {
    Write-Host "Keine Tape-Jobs gefunden."
} else {
    foreach ($tj in $tapeJobs) {
        Write-Sub "Tape-Job: $($tj.Name)"
        Write-Host "  Job-Objekttyp : $($tj.GetType().FullName)"
        try { Write-Host "  Enabled       : $($tj.Enabled)" } catch {}
        try { Write-Host "  LastResult    : $($tj.LastResult)" } catch {}
        try { Write-Host "  LastState     : $($tj.LastState)" } catch {}

        # --- Sessions holen ---
        $sessions = @()
        try { $sessions = @(Get-VBRTapeBackupSession -Job $tj -ErrorAction SilentlyContinue) } catch {}
        if ($sessions.Count -eq 0) {
            try { $sessions = @(Get-VBRTapeBackupSession -Name $tj.Name -ErrorAction SilentlyContinue) } catch {}
        }
        Write-Host "  Sessions gefunden : $($sessions.Count)"
        if ($sessions.Count -eq 0) { continue }

        $sess = @($sessions | Sort-Object CreationTime -Descending)[0]
        Write-Host "  Session-Objekttyp : $($sess.GetType().FullName)"
        try { Write-Host "  State / Result    : $($sess.State) / $($sess.Result)" } catch {}
        try { Write-Host "  CreationTime      : $($sess.CreationTime)" } catch {}

        # --- Welche Properties hat das Objekt ueberhaupt? ---
        Write-Host ""
        Write-Host "  [A] Properties der Session:"
        try {
            $props = @($sess | Get-Member -MemberType Property, NoteProperty -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name)
            Write-Host ("      " + ($props -join ", "))
        } catch { Write-Host "      <nicht auslesbar>" }

        Write-Host "  [B] Methoden der Session (nur interessante):"
        try {
            $meths = @($sess | Get-Member -MemberType Method -ErrorAction SilentlyContinue |
                       Where-Object { $_.Name -match 'Log|Detail|Task|Reason|Record' } |
                       Select-Object -ExpandProperty Name)
            if ($meths.Count -eq 0) { Write-Host "      <keine passenden>" }
            else { Write-Host ("      " + ($meths -join ", ")) }
        } catch { Write-Host "      <nicht auslesbar>" }

        # --- Quelle 1: .Log-Property ---
        Write-Host ""
        Write-Host "  [1] .Log-Property"
        try {
            $logItems = @($sess.Log)
            Write-Host "      Anzahl Eintraege : $($logItems.Count)"
            if ($logItems.Count -gt 0) {
                Write-Host "      Eintrags-Typ     : $($logItems[0].GetType().FullName)"
                $lprops = @($logItems[0] | Get-Member -MemberType Property, NoteProperty -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name)
                Write-Host ("      Felder           : " + ($lprops -join ", "))
                Write-Host "      --- Eintraege mit Status Warning/Error/Failed ---"
                $hits = @($logItems | Where-Object { "$($_.Status)" -match 'Warning|Error|Failed' })
                if ($hits.Count -eq 0) {
                    Write-Host "      <keine> - hier die letzten 5 Eintraege stattdessen:"
                    $tail = @($logItems | Select-Object -Last 5)
                    foreach ($t in $tail) {
                        Write-Host "        [$($t.Status)] $(Short ([string]$t.Title))"
                    }
                } else {
                    foreach ($h in @($hits | Select-Object -First 8)) {
                        Write-Host "        [$($h.Status)] $(Short ([string]$h.Title))"
                        $d = [string]$h.Description
                        if ($d) { Write-Host "                  Desc: $(Short $d)" }
                    }
                }
            }
        } catch { Write-Host "      FEHLER: $($_.Exception.Message)" }

        # --- Quelle 2: .Logger.GetLog() ---
        Write-Host ""
        Write-Host "  [2] .Logger.GetLog()"
        try {
            $lg = $sess.Logger.GetLog()
            if ($null -eq $lg) {
                Write-Host "      <null>"
            } else {
                Write-Host "      Log-Objekttyp : $($lg.GetType().FullName)"
                $att = $null
                try { $att = @($lg.GetAttentionRecords()) } catch { Write-Host "      GetAttentionRecords() nicht verfuegbar" }
                if ($att) {
                    Write-Host "      GetAttentionRecords() : $($att.Count) Eintraege"
                    foreach ($a in @($att | Select-Object -First 8)) {
                        Write-Host "        [$($a.Status)] $(Short ([string]$a.Title))"
                        $d = [string]$a.Description
                        if ($d) { Write-Host "                  Desc: $(Short $d)" }
                    }
                }
                try {
                    $upd = @($lg.UpdatedRecords)
                    Write-Host "      UpdatedRecords        : $($upd.Count) Eintraege"
                    $uh = @($upd | Where-Object { "$($_.Status)" -match 'Warning|Error|Failed' })
                    Write-Host "      davon Warning/Error   : $($uh.Count)"
                    foreach ($u in @($uh | Select-Object -First 8)) {
                        Write-Host "        [$($u.Status)] $(Short ([string]$u.Title))"
                    }
                } catch { Write-Host "      UpdatedRecords nicht verfuegbar" }
            }
        } catch { Write-Host "      FEHLER: $($_.Exception.Message)" }

        # --- Quelle 3: Volles Core-Objekt ueber die Session-Id ---
        Write-Host ""
        Write-Host "  [3] Get-VBRBackupSession -Id (volles Core-Objekt)"
        try {
            $full = Get-VBRBackupSession -Id $sess.Id -ErrorAction SilentlyContinue
            if ($null -eq $full) {
                Write-Host "      <nichts zurueckgeliefert>"
            } else {
                Write-Host "      Objekttyp : $($full.GetType().FullName)"
                try {
                    $fl = $full.Logger.GetLog()
                    $fatt = @($fl.GetAttentionRecords())
                    Write-Host "      GetAttentionRecords() : $($fatt.Count) Eintraege"
                    foreach ($a in @($fatt | Select-Object -First 8)) {
                        Write-Host "        [$($a.Status)] $(Short ([string]$a.Title))"
                        $d = [string]$a.Description
                        if ($d) { Write-Host "                  Desc: $(Short $d)" }
                    }
                } catch { Write-Host "      Logger nicht nutzbar: $($_.Exception.Message)" }
            }
        } catch { Write-Host "      FEHLER: $($_.Exception.Message)" }

        # --- Quelle 4: GetDetails() / Info.Reason ---
        Write-Host ""
        Write-Host "  [4] GetDetails() / Info.Reason"
        try { Write-Host "      GetDetails() : $(Short ([string]$sess.GetDetails()))" } catch { Write-Host "      GetDetails() nicht verfuegbar" }
        try { Write-Host "      Info.Reason  : $(Short ([string]$sess.Info.Reason))" }  catch { Write-Host "      Info.Reason nicht verfuegbar" }

        # --- Quelle 5: TaskSessions ---
        Write-Host ""
        Write-Host "  [5] TaskSessions (einzelne Objekte des Jobs)"
        $tasks = $null
        try { $tasks = @($sess.GetTaskSessions()) } catch {}
        if (-not $tasks -or $tasks.Count -eq 0) {
            try { $tasks = @(Get-VBRTaskSession -Session $sess -ErrorAction SilentlyContinue) } catch {}
        }
        if (-not $tasks -or $tasks.Count -eq 0) {
            Write-Host "      <keine gefunden>"
        } else {
            Write-Host "      Anzahl : $($tasks.Count), Typ: $($tasks[0].GetType().FullName)"
            foreach ($t in @($tasks | Select-Object -First 10)) {
                $st = ""
                try { $st = [string]$t.Result } catch {}
                if (-not $st) { try { $st = [string]$t.Status } catch {} }
                Write-Host "        - $($t.Name) : $st"
                try {
                    $tr = @($t.Logger.GetLog().GetAttentionRecords())
                    foreach ($x in @($tr | Select-Object -First 3)) {
                        Write-Host "            [$($x.Status)] $(Short ([string]$x.Title))"
                    }
                } catch {}
            }
        }
    }
}

# -----------------------------------------------------------------------------
# 2. PLATFORM-JOBS: WIE ERKENNT MAN KIND-JOBS?
# -----------------------------------------------------------------------------
Write-Head "2. PLATFORM-JOBS - ERKENNUNG INTERNER KIND-JOBS"

if (-not $coreOk) {
    Write-Host "Kein Zugriff auf Veeam.Backup.Core - Abschnitt uebersprungen."
} else {
    $allJobs = @()
    try {
        $allJobs = @([Veeam.Backup.Core.CBackupJob]::GetAll() | Where-Object {
            "$($_.TypeToString)" -match 'Proxmox|Nutanix|AHV|oVirt|RHV|OLVM|Scale Computing|Morpheus'
        })
    } catch { Write-Host "FEHLER: $($_.Exception.Message)" }

    Write-Host "Gefundene Platform-Jobs (inkl. Kindjobs): $($allJobs.Count)"

    if ($allJobs.Count -gt 0) {
        Write-Sub "Verfuegbare Properties am Job-Objekt"
        try {
            $jprops = @($allJobs[0] | Get-Member -MemberType Property, NoteProperty -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name)
            Write-Host ("  " + ($jprops -join ", "))
        } catch { Write-Host "  <nicht auslesbar>" }

        Write-Sub "Kandidaten-Properties zur Kind-Erkennung"
        Write-Host ("  {0,-45} {1,-22} {2,-38} {3,-8} {4}" -f "Name", "TypeToString", "ParentJobId", "IsChild", "Enabled")
        Write-Host ("  " + ("-" * 130))
        foreach ($j in $allJobs) {
            $nm = "$($j.Name)"
            if ($nm.Length -gt 44) { $nm = $nm.Substring(0, 41) + "..." }

            $tp = ""
            try { $tp = "$($j.TypeToString)".Trim() } catch {}

            $parent = "<n/a>"
            try { if ($null -ne $j.ParentJobId) { $parent = "$($j.ParentJobId)" } } catch {}

            $child = "<n/a>"
            try { if ($null -ne $j.IsChild) { $child = "$($j.IsChild)" } } catch {}

            $en = ""
            try { $en = "$($j.IsScheduleEnabled)" } catch {}

            Write-Host ("  {0,-45} {1,-22} {2,-38} {3,-8} {4}" -f $nm, $tp, $parent, $child, $en)
        }

        Write-Sub "Weitere moegliche Kind-Marker am ersten Job"
        $probe = @('ParentJobId','ParentScheduleId','IsChild','IsChildJob','LinkedJobIds','SourceType','JobType','ObjectId')
        foreach ($p in $probe) {
            $val = "<nicht vorhanden>"
            try {
                $v2 = $allJobs[0].$p
                if ($null -ne $v2) { $val = Short ([string]$v2) 60 }
                else { $val = "<null>" }
            } catch {}
            Write-Host ("  {0,-20} = {1}" -f $p, $val)
        }
    }
}

Write-Head "ENDE - bitte die gesamte Ausgabe kopieren"
