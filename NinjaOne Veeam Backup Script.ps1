# =============================================================================
# NINJAONE VEEAM UNIVERSAL CHECK (V10 - V13 KOMPATIBEL)
# =============================================================================
#
# Autor: Sebastian Herrmann
#
# HINWEIS ZUR CODIERUNG: Diese Datei ist bewusst reines ASCII (ae/oe/ue statt
# Umlaute), damit sie beim Copy/Paste in NinjaOne nicht zerschossen wird.
#
# DURCHGEFUEHRTE CHECKS:
# 1. VM Backups (Get-VBRJob / Get-VBRBackupSession):
#    - Prueft alle *aktivierten* Jobs (IsScheduleEnabled = True).
#    - Holt die allerletzte *abgeschlossene* Session des Jobs (State = Stopped).
#    - Meldet Status (Success, Warning, Failed) und prueft einzelne VMs.
#
# 2. Platform Backups (Proxmox VE, Nutanix AHV, oVirt/RHV/OLVM, Scale, Morpheus):
#    - WICHTIG: Diese Jobs werden von Get-VBRJob NICHT zurueckgeliefert!
#    - Weg A (offiziell, ab V12.3): Get-VBRSession -Type PlatformBackupJob
#    - Weg B (Fallback, .NET):      [Veeam.Backup.Core.CBackupJob]::GetAll()
#
# 3. Backup Copy Jobs (Get-VBRBackupCopyJob)
# 4. Agent Backups (Get-VBRComputerBackupJob)
# 5. Tape/Band Jobs (Get-VBRTapeJob):
#    - Inkl. Erkennung von Sessions die auf ein Band warten (WaitingTape).
# 6. M365 / VBO (Get-VBOJobSession)
#
# NEU: KLARTEXT-MELDUNGEN AUS DEM VEEAM-LOG
#    Bei jedem Job/Objekt das NICHT "Success" ist, werden die eigentlichen
#    Warn-/Fehlermeldungen aus dem Session-Log ausgelesen
#    (Logger.GetLog().GetAttentionRecords() -> Title + Description).
#    Damit erscheint z.B. statt nur "Warning" die Ursache wie
#    "No suitable media found" / "Tape is not available".
#
# =============================================================================

# 1. Ninja-Properties abrufen
$successProp = Ninja-Property-Get veeamsuccess
$warningProp = Ninja-Property-Get veeamwarning
$failProp    = Ninja-Property-Get veeamfail
$logProp     = Ninja-Property-Get veeamlog

# -----------------------------------------------------------------------------
# DER CHECK-CODE (In einem Block, damit er ueberall laufen kann)
# -----------------------------------------------------------------------------
$VeeamCheckLogic = @'
    $WarningPreference = 'SilentlyContinue'
    $ProgressPreference = 'SilentlyContinue'

    $res = @{
        Success = $false
        Warning = $false
        Failed  = $false
        Log     = ""
    }

    $LogListFull    = New-Object System.Collections.Generic.List[string]
    $LogListCompact = New-Object System.Collections.Generic.List[string]

    # Jobnamen die bereits gemeldet wurden (verhindert Doppel-Meldungen)
    $HandledJobNames = @{}

    function Add-Log {
        param([string]$msg, [switch]$IsDetail)
        $LogListFull.Add($msg)
        if (-not $IsDetail) { $LogListCompact.Add($msg) }
    }

    function Invoke-WithRetry {
        param([ScriptBlock]$Command, [int]$MaxRetries = 3, [int]$DelaySeconds = 5)
        $retryCount = 0
        $lastException = $null

        while ($retryCount -lt $MaxRetries) {
            try { return (& $Command) }
            catch {
                $lastException = $_
                $retryCount++
                if ($retryCount -lt $MaxRetries) { Start-Sleep -Seconds $DelaySeconds }
            }
        }
        throw $lastException
    }

    # =========================================================================
    # HILFSFUNKTIONEN FUER KLARTEXT-MELDUNGEN AUS DEM VEEAM-LOG
    # =========================================================================

    # Bereinigt eine Veeam-Meldung (HTML-Tags, Zeilenumbrueche, Laenge)
    function Format-Msg {
        param([string]$Text, [int]$MaxLen = 260)
        if ([string]::IsNullOrEmpty($Text)) { return "" }
        $t = $Text
        $t = $t -replace '<br\s*/?>', ' '
        $t = $t -replace '<[^>]+>', ' '
        $t = $t -replace '&nbsp;', ' '
        $t = $t -replace '&quot;', '"'
        $t = $t -replace '&apos;', "'"
        $t = $t -replace '&#39;', "'"
        $t = $t -replace '&lt;', '<'
        $t = $t -replace '&gt;', '>'
        $t = $t -replace '&amp;', '&'
        $t = $t -replace '[\r\n\t]+', ' '
        $t = $t -replace '\s{2,}', ' '
        $t = $t.Trim()
        if ([string]::IsNullOrEmpty($t)) { return "" }
        if ($t.Length -gt $MaxLen) { $t = $t.Substring(0, $MaxLen) + " [...]" }
        return $t
    }

    # Nichtssagende Standard-Abschlussmeldungen nach hinten sortieren
    function Test-IsNoiseMsg {
        param([string]$Text)
        if ([string]::IsNullOrEmpty($Text)) { return $true }
        return ($Text -match '(finished with|completed with|Processing finished|Job finished|abgeschlossen mit)')
    }

    # Holt die eigentlichen Warn-/Fehlermeldungen aus einer Session oder TaskSession
    function Get-DetailMessages {
        param($Obj, [int]$Max = 5)

        $primary   = New-Object System.Collections.Generic.List[string]
        $secondary = New-Object System.Collections.Generic.List[string]
        $out       = New-Object System.Collections.Generic.List[string]

        if ($null -eq $Obj) { return ,$out }

        # --- 1) Logger-Records: hier steht die Ursache im Klartext ------------
        $records = $null
        try {
            $log = $Obj.Logger.GetLog()
            if ($log) {
                try { $records = $log.GetAttentionRecords() } catch {}
                if (-not $records) {
                    try {
                        $records = @($log.UpdatedRecords | Where-Object { "$($_.Status)" -match 'Warning|Failed|Error' })
                    } catch {}
                }
            }
        } catch {}

        foreach ($r in $records) {
            $title = ""
            $desc  = ""
            try { $title = Format-Msg -Text ([string]$r.Title) } catch {}
            try { $desc  = Format-Msg -Text ([string]$r.Description) } catch {}

            $line = $title
            if ($desc -and $desc -ne $title) {
                if ($line) { $line = "$line :: $desc" } else { $line = $desc }
            }
            if (-not $line) { continue }

            if (Test-IsNoiseMsg -Text $line) {
                if (-not $secondary.Contains($line)) { $secondary.Add($line) }
            } else {
                if (-not $primary.Contains($line)) { $primary.Add($line) }
            }
        }

        foreach ($p in $primary)   { if (-not $out.Contains($p)) { $out.Add($p) } }
        foreach ($s in $secondary) { if (-not $out.Contains($s)) { $out.Add($s) } }

        # --- 2) Fallback: GetDetails() ---------------------------------------
        if ($out.Count -eq 0) {
            try {
                $d = Format-Msg -Text ([string]$Obj.GetDetails())
                if ($d) { $out.Add($d) }
            } catch {}
        }

        # --- 3) Fallback: Info.Reason ----------------------------------------
        if ($out.Count -eq 0) {
            try {
                $d = Format-Msg -Text ([string]$Obj.Info.Reason)
                if ($d) { $out.Add($d) }
            } catch {}
        }

        # --- 4) Fallback: Description / Info.Description ----------------------
        if ($out.Count -eq 0) {
            try {
                $d = Format-Msg -Text ([string]$Obj.Description)
                if ($d) { $out.Add($d) }
            } catch {}
        }
        if ($out.Count -eq 0) {
            try {
                $d = Format-Msg -Text ([string]$Obj.Info.Description)
                if ($d) { $out.Add($d) }
            } catch {}
        }

        if ($out.Count -gt $Max) {
            $trim = $out.GetRange(0, $Max)
            $trim.Add("[... " + ($out.Count - $Max) + " weitere Meldungen im Veeam-Log]")
            return ,$trim
        }
        return ,$out
    }

    # Holt die letzten N Log-Eintraege (unabhaengig vom Status).
    # Wichtig fuer laufende/wartende Sessions ("Waiting for tape ...").
    function Get-LastLogRecords {
        param($Obj, [int]$Count = 3)
        $out = New-Object System.Collections.Generic.List[string]
        if ($null -eq $Obj) { return ,$out }
        try {
            $recs = @($Obj.Logger.GetLog().UpdatedRecords)
            $start = 0
            if ($recs.Count -gt $Count) { $start = $recs.Count - $Count }
            for ($i = $start; $i -lt $recs.Count; $i++) {
                $t = Format-Msg -Text ([string]$recs[$i].Title)
                if ($t -and -not $out.Contains($t)) { $out.Add($t) }
            }
        } catch {}
        return ,$out
    }

    # Schreibt eine Meldungsliste eingerueckt ins Log
    function Add-Messages {
        param($Messages, [string]$Prefix = "    > ")
        if ($null -eq $Messages) { return }
        foreach ($m in $Messages) {
            if ($m) { Add-Log ($Prefix + $m) }
        }
    }

    # Ermittelt den Status-String eines Objekts robust ueber mehrere Properties
    function Get-StatusString {
        param($Obj)
        $s = ""
        try { $s = [string]$Obj.Result } catch {}
        if ([string]::IsNullOrEmpty($s) -or $s -eq "None") { try { $s2 = [string]$Obj.State;  if ($s2) { $s = $s2 } } catch {} }
        if ([string]::IsNullOrEmpty($s)) { try { $s3 = [string]$Obj.Status; if ($s3) { $s = $s3 } } catch {} }
        if ([string]::IsNullOrEmpty($s)) { $s = "Unknown" }
        return $s
    }

    # --- MODULE LADEN ---
    if (-not (Get-Command Get-VBRJob -ErrorAction SilentlyContinue)) {
        try {
            if (Get-Module -ListAvailable -Name "Veeam.Backup.PowerShell") {
                Import-Module "Veeam.Backup.PowerShell" -DisableNameChecking -WarningAction SilentlyContinue -ErrorAction Stop | Out-Null
            }
        } catch {}

        if (-not (Get-Command Get-VBRJob -ErrorAction SilentlyContinue)) {
            try {
                if ((Get-PSSnapin -Name VeeamPSSnapin -ErrorAction SilentlyContinue) -eq $null) {
                    Add-PSSnapin VeeamPSSnapin -ErrorAction SilentlyContinue | Out-Null
                }
            } catch {}
        }
    }

    if (-not (Get-Command Get-VBRJob -ErrorAction SilentlyContinue)) {
        Add-Log "ERROR: Veeam Module/Snapins konnten nicht geladen werden."
        $res.Failed = $true
        $res.Log = $LogListFull -join "`n"
        return ($res | ConvertTo-Json -Depth 2)
    }

    # Die interne Core-Assembly wird erst durch einen echten VBR-Aufruf geladen.
    # Ohne das schlaegt [Veeam.Backup.Core.CBackupJob] fehl.
    try { Get-VBRLicenseAutoUpdateStatus -ErrorAction SilentlyContinue | Out-Null } catch {}

    # =========================================================================
    # 1. VM BACKUP JOBS
    # =========================================================================
    try {
        $latestSessions = Invoke-WithRetry -Command {
            Get-VBRBackupSession |
            Where-Object { $_.State -eq "Stopped" } |
            Sort-Object JobName, CreationTime -Descending |
            Group-Object JobName |
            ForEach-Object { $_.Group[0] } |
            Where-Object {
                $job = Get-VBRJob -Name $_.JobName -ErrorAction SilentlyContinue
                if ($job) { $job.IsScheduleEnabled -eq $true } else { $false }
            }
        }

        if ($latestSessions) {
            foreach ($session in $latestSessions) {
                $jobName   = $session.JobName
                $jobStatus = $session.Result
                $jobTime   = $session.CreationTime.ToString("dd.MM.yyyy HH:mm:ss")

                $HandledJobNames["$jobName"] = $true

                Add-Log "--------------------------------------"
                Add-Log "Backup-Job: $jobName | Status: $jobStatus | Zeitpunkt: $jobTime"

                # NEU: Klartext-Meldungen auf Job-Ebene
                if ("$jobStatus" -ne "Success") {
                    Add-Messages -Messages (Get-DetailMessages -Obj $session -Max 5) -Prefix "    > "
                }

                $tasks = $session.GetTaskSessions()
                foreach ($task in $tasks) {
                    $vmName   = $task.Name
                    $vmStatus = Get-StatusString -Obj $task

                    if ("$vmStatus" -eq "Success") {
                        # Erfolgs-Details duerfen bei Platzmangel wegfallen
                        Add-Log "    - VM: $vmName | Status: $vmStatus" -IsDetail
                    } else {
                        # NEU: Problem-VMs samt Ursache immer mitloggen
                        Add-Log "    - VM: $vmName | Status: $vmStatus"
                        Add-Messages -Messages (Get-DetailMessages -Obj $task -Max 3) -Prefix "        > "
                    }
                }

                if ($jobStatus -eq "Success") { $res.Success = $true }
                if ($jobStatus -eq "Warning") { $res.Warning = $true }
                if ($jobStatus -eq "Failed")  { $res.Failed  = $true }
            }
        } else {
            Add-Log "--------------------------------------"
            Add-Log "Keine aktiven (abgeschlossenen) VM-Backup-Jobs gefunden."
        }
    } catch {
        Add-Log "Fehler bei VM-Backup-Abfrage: $($_.Exception.Message)"
        $res.Failed = $true
    }

    # =========================================================================
    # 2. PLATFORM BACKUP JOBS
    #    Proxmox VE, Nutanix AHV, oVirt/RHV/OLVM, Scale Computing, Morpheus
    #
    #    WICHTIG: Diese Jobs liefert Get-VBRJob NICHT zurueck. Deshalb sind sie
    #    in Abschnitt 1 systematisch durch den Get-VBRJob-Filter gefallen.
    # =========================================================================
    try {
        # --- Job-Liste ueber die interne Core-Klasse (Name, Typ, Zeitplan) ----
        $platformJobs = @()
        try {
            $platformJobs = @([Veeam.Backup.Core.CBackupJob]::GetAll() | Where-Object {
                "$($_.TypeToString)" -match 'Proxmox|Nutanix|AHV|oVirt|RHV|OLVM|Scale Computing|Morpheus'
            })
        } catch {}

        $platformJobById = @{}
        foreach ($pj in $platformJobs) {
            try { $platformJobById["$($pj.Id)"] = $pj } catch {}
        }

        # --- Sessions holen: Weg A (offiziell, ab V12.3) ----------------------
        $platformSessions = @()
        $platformCandidates = @()
        try {
            $rawPlatform = @(Get-VBRSession -Type PlatformBackupJob -ErrorAction Stop)
            if ($rawPlatform.Count -gt 0) {
                $stopped = @($rawPlatform | Where-Object { "$($_.State)" -eq "Stopped" })
                if ($stopped.Count -eq 0) { $stopped = $rawPlatform }

                $platformCandidates = @(
                    $stopped |
                    Sort-Object CreationTime -Descending |
                    Group-Object JobId |
                    ForEach-Object { $_.Group[0] }
                )
            }
        } catch {}

        foreach ($cand in $platformCandidates) {
            $full = $null
            # Die "leichte" Session hat kein Logger/GetTaskSessions -> volle Session nachladen
            try { $full = Get-VBRBackupSession -Id $cand.Id -ErrorAction SilentlyContinue } catch {}
            if ($full) { $platformSessions += $full } else { $platformSessions += $cand }
        }

        # --- Weg B (Fallback ueber Core-Klasse), falls Weg A nichts lieferte ---
        if ($platformSessions.Count -eq 0 -and $platformJobs.Count -gt 0) {
            foreach ($pj in $platformJobs) {
                $s = $null
                try { $s = $pj.FindLastSession() } catch {}
                if (-not $s) {
                    try {
                        $s = @([Veeam.Backup.Core.CBackupSession]::GetByJob($pj.Id) | Sort-Object CreationTime -Descending)[0]
                    } catch {}
                }
                if ($s) { $platformSessions += $s }
            }
        }

        $reportedPlatformJobs = @{}

        foreach ($session in $platformSessions) {
            if ($null -eq $session) { continue }

            $pjob  = $null
            $jobId = ""
            try { $jobId = "$($session.JobId)" } catch {}
            if ($jobId -and $platformJobById.ContainsKey($jobId)) { $pjob = $platformJobById[$jobId] }

            $jobName = ""
            try { $jobName = [string]$session.JobName } catch {}
            if ([string]::IsNullOrEmpty($jobName) -and $pjob) { $jobName = [string]$pjob.Name }
            if ([string]::IsNullOrEmpty($jobName)) { $jobName = "<Unbekannt>" }

            # Doppelmeldung vermeiden, falls der Job doch in Abschnitt 1 auftauchte
            if ($HandledJobNames.ContainsKey("$jobName")) { continue }
            if ($reportedPlatformJobs.ContainsKey("$jobName")) { continue }

            # Nur aktivierte Jobs, sofern der Zeitplan ermittelbar ist
            if ($pjob -and ($pjob.IsScheduleEnabled -eq $false)) { continue }

            # Plattform-Bezeichnung ermitteln (z.B. "Proxmox VE")
            $platformName = ""
            try { if ($session.Platform) { $platformName = [string]$session.Platform.ToHumanReadable() } } catch {}
            if ([string]::IsNullOrEmpty($platformName) -and $pjob) {
                try { $platformName = [string]$pjob.TypeToString } catch {}
            }
            if ([string]::IsNullOrEmpty($platformName)) { $platformName = "Platform" }

            $jobStatus = Get-StatusString -Obj $session
            $jobTime   = "<unbekannt>"
            try { $jobTime = $session.CreationTime.ToString("dd.MM.yyyy HH:mm:ss") } catch {}

            Add-Log "--------------------------------------"
            Add-Log "$platformName-Job: $jobName | Status: $jobStatus | Zeitpunkt: $jobTime"
            $reportedPlatformJobs["$jobName"] = $true
            $HandledJobNames["$jobName"] = $true

            if ("$jobStatus" -ne "Success") {
                Add-Messages -Messages (Get-DetailMessages -Obj $session -Max 5) -Prefix "    > "
            }

            # --- Einzelne VMs des Platform-Jobs ------------------------------
            $pTasks = $null
            try { $pTasks = $session.GetTaskSessions() } catch {}
            if (-not $pTasks) {
                try { $pTasks = Get-VBRTaskSession -Session $session -ErrorAction SilentlyContinue } catch {}
            }

            foreach ($task in $pTasks) {
                $vmName   = ""
                try { $vmName = [string]$task.Name } catch {}
                if ([string]::IsNullOrEmpty($vmName)) { $vmName = "<Objekt>" }
                $vmStatus = Get-StatusString -Obj $task

                if ("$vmStatus" -eq "Success") {
                    Add-Log "    - VM: $vmName | Status: $vmStatus" -IsDetail
                } else {
                    Add-Log "    - VM: $vmName | Status: $vmStatus"
                    Add-Messages -Messages (Get-DetailMessages -Obj $task -Max 3) -Prefix "        > "
                }
            }

            if ($jobStatus -eq "Success") { $res.Success = $true }
            if ($jobStatus -eq "Warning") { $res.Warning = $true }
            if ($jobStatus -eq "Failed")  { $res.Failed  = $true }
        }

        # --- Konfigurierte Platform-Jobs ohne jede Session melden -------------
        foreach ($pj in $platformJobs) {
            $pName = ""
            try { $pName = [string]$pj.Name } catch {}
            if ([string]::IsNullOrEmpty($pName)) { continue }
            if ($reportedPlatformJobs.ContainsKey("$pName")) { continue }
            if ($HandledJobNames.ContainsKey("$pName")) { continue }
            if ($pj.IsScheduleEnabled -eq $false) { continue }

            $pType = "Platform"
            try { $pType = [string]$pj.TypeToString } catch {}

            Add-Log "--------------------------------------"
            Add-Log "$pType-Job: $pName | Status: Keine abgeschlossene Session gefunden."
        }

        if ($platformSessions.Count -eq 0 -and $platformJobs.Count -eq 0) {
            Add-Log "--------------------------------------"
            Add-Log "Keine Platform-Jobs (Proxmox VE / Nutanix AHV / oVirt) konfiguriert."
        }
    } catch {
        Add-Log "Fehler bei Platform-Job-Abfrage (Proxmox/Nutanix/oVirt): $($_.Exception.Message)"
    }

    # =========================================================================
    # 3. BACKUP COPY JOBS
    # =========================================================================
    try {
        if (Get-Command Get-VBRBackupCopyJob -ErrorAction SilentlyContinue) {
            $copyJobs = Invoke-WithRetry -Command {
                Get-VBRBackupCopyJob -ErrorAction SilentlyContinue | Where-Object {
                    $baseJob = Get-VBRJob -Name $_.Name -ErrorAction SilentlyContinue
                    if ($baseJob) { ($baseJob | Select-Object -First 1).IsScheduleEnabled -eq $true } else { $true }
                }
            }

            if ($copyJobs) {
                foreach ($cJob in $copyJobs) {
                    $cJobName = $cJob.Name
                    $lastCSession = $null

                    try {
                        $lastCSession = Invoke-WithRetry -Command { Get-VBRBackupSession -Job $cJob -ErrorAction SilentlyContinue | Where-Object { $_.State -eq "Stopped" } | Sort-Object CreationTime -Descending | Select-Object -First 1 }
                    } catch {}

                    if (-not $lastCSession) {
                        try {
                            $lastCSession = Invoke-WithRetry -Command { Get-VBRBackupSession -Name $cJobName -ErrorAction SilentlyContinue | Where-Object { $_.State -eq "Stopped" } | Sort-Object CreationTime -Descending | Select-Object -First 1 }
                        } catch {}
                    }

                    if ($lastCSession) {
                        $cStatus = $lastCSession.Result
                        $cTime   = $lastCSession.CreationTime.ToString("dd.MM.yyyy HH:mm:ss")

                        Add-Log "--------------------------------------"
                        Add-Log "Copy-Job: $cJobName | Status: $cStatus | Zeitpunkt: $cTime"

                        if ("$cStatus" -ne "Success") {
                            Add-Messages -Messages (Get-DetailMessages -Obj $lastCSession -Max 5) -Prefix "    > "
                        }

                        if ($cStatus -eq "Success") { $res.Success = $true }
                        if ($cStatus -eq "Warning") { $res.Warning = $true }
                        if ($cStatus -eq "Failed")  { $res.Failed  = $true }
                    } else {
                         $jobResult = $cJob.LastResult
                         if ($jobResult -and $jobResult -ne "None" -and $jobResult -ne "Working") {
                             Add-Log "--------------------------------------"
                             Add-Log "Copy-Job: $cJobName | Status: $jobResult (Job-Status) | Zeitpunkt: <Keine Session Log>"

                             if ($jobResult -eq "Success") { $res.Success = $true }
                             if ($jobResult -eq "Warning") { $res.Warning = $true }
                             if ($jobResult -eq "Failed")  { $res.Failed  = $true }
                         } else {
                             Add-Log "--------------------------------------"
                             Add-Log "Copy-Job: $cJobName | Status: Keine abgeschlossene Session gefunden."
                         }
                    }
                }
            } else {
                Add-Log "--------------------------------------"
                Add-Log "Keine aktiven Backup-Copy-Jobs konfiguriert."
            }
        }
    } catch {
        Add-Log "Fehler bei Copy-Job-Abfrage: $($_.Exception.Message)"
        $res.Failed = $true
    }

    # =========================================================================
    # 4. AGENT JOBS (Windows/Linux/Mac)
    # =========================================================================
    try {
        if (Get-Command Get-VBRComputerBackupJob -ErrorAction SilentlyContinue) {
            $agentJobs = Invoke-WithRetry -Command {
                Get-VBRComputerBackupJob -ErrorAction SilentlyContinue | Where-Object {
                    $baseJob = Get-VBRJob -Name $_.Name -ErrorAction SilentlyContinue
                    if ($baseJob) { ($baseJob | Select-Object -First 1).IsScheduleEnabled -eq $true } else { $true }
                }
            }

            if ($agentJobs) {
                foreach ($job in $agentJobs) {
                    $jobName = $job.Name
                    $lastSession = $null
                    try {
                        $lastSession = Invoke-WithRetry -Command { Get-VBRComputerBackupJobSession -Name $jobName -ErrorAction SilentlyContinue | Where-Object { $_.State -eq "Stopped" } | Sort-Object CreationTime -Descending | Select-Object -First 1 }
                    } catch {}

                    if ($lastSession) {
                        $agentStatus = $lastSession.Result
                        $agentTime   = $lastSession.CreationTime.ToString("dd.MM.yyyy HH:mm:ss")

                        Add-Log "--------------------------------------"
                        Add-Log "Agent-Job: $jobName | Status: $agentStatus | Zeitpunkt: $agentTime"

                        if ("$agentStatus" -ne "Success") {
                            Add-Messages -Messages (Get-DetailMessages -Obj $lastSession -Max 5) -Prefix "    > "

                            # Einzelne Rechner des Agent-Jobs mit Problem
                            $aTasks = $null
                            try { $aTasks = $lastSession.GetTaskSessions() } catch {}
                            if (-not $aTasks) {
                                try { $aTasks = Get-VBRTaskSession -Session $lastSession -ErrorAction SilentlyContinue } catch {}
                            }
                            foreach ($task in $aTasks) {
                                $tStatus = Get-StatusString -Obj $task
                                if ("$tStatus" -ne "Success") {
                                    Add-Log "    - Rechner: $($task.Name) | Status: $tStatus"
                                    Add-Messages -Messages (Get-DetailMessages -Obj $task -Max 3) -Prefix "        > "
                                }
                            }
                        }

                        if ($agentStatus -eq "Success") { $res.Success = $true }
                        if ($agentStatus -eq "Warning") { $res.Warning = $true }
                        if ($agentStatus -eq "Failed")  { $res.Failed  = $true }
                        if ($agentStatus -eq "None")    { $res.Failed  = $true }
                    } else {
                        Add-Log "--------------------------------------"
                        Add-Log "Agent-Job: $jobName | Status: Keine abgeschlossene Session gefunden."
                    }
                }
            } else {
                Add-Log "--------------------------------------"
                Add-Log "Keine aktiven Agent-Backup-Jobs gefunden."
            }
        }
    } catch {
        Add-Log "Fehler bei Agent-Backup-Abfrage: $($_.Exception.Message)"
    }

    # =========================================================================
    # 5. TAPE JOBS
    #    NEU: - Klartext-Ursachen (z.B. "No suitable media found")
    #         - Erkennung von Sessions die auf ein Band warten (WaitingTape)
    #         - Fallback auf den Job-Status, wenn keine Session gefunden wird
    # =========================================================================
    try {
        if (Get-Command Get-VBRTapeJob -ErrorAction SilentlyContinue) {
            $tapeJobs = Invoke-WithRetry -Command {
                Get-VBRTapeJob -ErrorAction SilentlyContinue | Where-Object {
                    $baseJob = Get-VBRJob -Name $_.Name -ErrorAction SilentlyContinue
                    if ($baseJob) {
                        ($baseJob | Select-Object -First 1).IsScheduleEnabled -eq $true
                    } else {
                        # Tape-Jobs tauchen in Get-VBRJob i.d.R. nicht auf -> eigenes Enabled-Flag
                        $_.Enabled -ne $false
                    }
                }
            }

            if ($tapeJobs) {
                foreach ($tapeJob in $tapeJobs) {
                    $tapeJobName = $tapeJob.Name

                    $allTapeSessions = @()
                    try {
                        $allTapeSessions = @(Invoke-WithRetry -Command { Get-VBRTapeBackupSession -Job $tapeJob -ErrorAction SilentlyContinue })
                    } catch {}
                    if ($allTapeSessions.Count -eq 0) {
                        try {
                            $allTapeSessions = @(Invoke-WithRetry -Command { Get-VBRTapeBackupSession -Name $tapeJobName -ErrorAction SilentlyContinue })
                        } catch {}
                    }

                    $latestSession = $null
                    try {
                        $latestSession = @($allTapeSessions | Where-Object { "$($_.State)" -eq "Stopped" } | Sort-Object CreationTime -Descending)[0]
                    } catch {}

                    # --- Laufende Session, die auf ein Band wartet -------------
                    $activeSession = $null
                    try {
                        $activeSession = @($allTapeSessions | Where-Object { "$($_.State)" -ne "Stopped" } | Sort-Object CreationTime -Descending)[0]
                    } catch {}

                    if ($null -ne $activeSession) {
                        $activeState = "$($activeSession.State)"
                        if ($activeState -match 'Wait|Idle|Pending') {
                            $aTime = "<unbekannt>"
                            try { $aTime = $activeSession.CreationTime.ToString("dd.MM.yyyy HH:mm:ss") } catch {}

                            Add-Log "--------------------------------------"
                            Add-Log "Tape-Job: $tapeJobName | Status: WARTET ($activeState) | Start: $aTime"

                            $waitMsgs = @(Get-DetailMessages -Obj $activeSession -Max 3)
                            if ($waitMsgs.Count -eq 0) {
                                $waitMsgs = @(Get-LastLogRecords -Obj $activeSession -Count 3)
                            }
                            Add-Messages -Messages $waitMsgs -Prefix "    > "
                            Add-Log "    > Hinweis: Job wartet auf Benutzeraktion (z.B. Band einlegen / Medium wechseln)."

                            $res.Warning = $true
                        }
                    }

                    if ($null -ne $latestSession) {
                        $tapeStatus = Get-StatusString -Obj $latestSession
                        $tapeTime   = "<unbekannt>"
                        try { $tapeTime = $latestSession.CreationTime.ToString("dd.MM.yyyy HH:mm:ss") } catch {}

                        Add-Log "--------------------------------------"
                        Add-Log "Tape-Job: $tapeJobName | Status: $tapeStatus | Zeitpunkt: $tapeTime"

                        # NEU: die eigentliche Ursache aus dem Session-Log
                        if ("$tapeStatus" -ne "Success") {
                            $tapeMsgs = @(Get-DetailMessages -Obj $latestSession -Max 6)

                            # Fallback: manche Wrapper-Objekte haben keinen Logger ->
                            # Session direkt ueber die Core-Klasse nachladen
                            if ($tapeMsgs.Count -eq 0) {
                                try {
                                    $coreSession = @([Veeam.Backup.Core.CBackupSession]::GetByJob($tapeJob.Id) | Sort-Object CreationTime -Descending)[0]
                                    if ($coreSession) { $tapeMsgs = @(Get-DetailMessages -Obj $coreSession -Max 6) }
                                } catch {}
                            }

                            Add-Messages -Messages $tapeMsgs -Prefix "    > "

                            # Einzelne Objekte des Tape-Jobs (Backups/Dateien)
                            $tTasks = $null
                            try { $tTasks = $latestSession.GetTaskSessions() } catch {}
                            if (-not $tTasks) {
                                try { $tTasks = Get-VBRTaskSession -Session $latestSession -ErrorAction SilentlyContinue } catch {}
                            }
                            foreach ($task in $tTasks) {
                                $tStatus = Get-StatusString -Obj $task
                                if ("$tStatus" -ne "Success") {
                                    Add-Log "    - Objekt: $($task.Name) | Status: $tStatus"
                                    Add-Messages -Messages (Get-DetailMessages -Obj $task -Max 3) -Prefix "        > "
                                }
                            }
                        }

                        if ($tapeStatus -eq "Success") { $res.Success = $true }
                        if ($tapeStatus -eq "Warning") { $res.Warning = $true }
                        if ($tapeStatus -eq "Failed")  { $res.Failed  = $true }
                        if ($tapeStatus -eq "None")    { $res.Failed  = $true }
                    }
                    elseif ($null -eq $activeSession) {
                        # NEU: bisher wurde hier gar nichts gemeldet (stille Luecke)
                        $tapeJobResult = ""
                        try { $tapeJobResult = [string]$tapeJob.LastResult } catch {}

                        Add-Log "--------------------------------------"
                        if ($tapeJobResult -and $tapeJobResult -ne "None") {
                            Add-Log "Tape-Job: $tapeJobName | Status: $tapeJobResult (Job-Status) | Zeitpunkt: <Keine Session Log>"
                            if ($tapeJobResult -eq "Success") { $res.Success = $true }
                            if ($tapeJobResult -eq "Warning") { $res.Warning = $true }
                            if ($tapeJobResult -eq "Failed")  { $res.Failed  = $true }
                        } else {
                            Add-Log "Tape-Job: $tapeJobName | Status: Keine abgeschlossene Session gefunden."
                        }
                    }
                }
            } else {
                Add-Log "--------------------------------------"
                Add-Log "Keine aktiven Tape-Jobs konfiguriert."
            }
        }
    } catch { Add-Log "Fehler bei Tape: $($_.Exception.Message)" }

    # =========================================================================
    # 6. M365 (VBO)
    # =========================================================================
    if (Get-Module -Name Veeam.Archiver.PowerShell -ListAvailable) {
        Import-Module Veeam.Archiver.PowerShell -DisableNameChecking -WarningAction SilentlyContinue -ErrorAction SilentlyContinue | Out-Null
        try {
            $m365Sessions = Invoke-WithRetry -Command {
                Get-VBOJobSession |
                Where-Object { $_.CreationTime -ge (Get-Date).AddDays(-1) -and $_.Status -ne "Running" } |
                Sort-Object CreationTime -Descending |
                Group-Object JobName |
                ForEach-Object { $_.Group[0] }
            }

            if ($m365Sessions) {
                foreach ($session in $m365Sessions) {
                    $jobName   = $session.JobName
                    $jobStatus = $session.Status
                    $jobTime   = $session.CreationTime.ToString("dd.MM.yyyy HH:mm:ss")

                    Add-Log "--------------------------------------"
                    Add-Log "M365-Job: $jobName | Status: $jobStatus | Zeitpunkt: $jobTime"

                    # NEU: Klartext aus dem VBO-Sessionlog
                    if ("$jobStatus" -ne "Success") {
                        $vboMsgs = New-Object System.Collections.Generic.List[string]
                        try {
                            $vboLog = @($session.Log | Where-Object { "$($_.Status)" -match 'Warning|Error|Failed' })
                            foreach ($l in $vboLog) {
                                $t = Format-Msg -Text ([string]$l.Title)
                                if ($t -and -not $vboMsgs.Contains($t)) { $vboMsgs.Add($t) }
                                if ($vboMsgs.Count -ge 5) { break }
                            }
                        } catch {}
                        Add-Messages -Messages $vboMsgs -Prefix "    > "
                    }

                    if ($jobStatus -eq "Success") { $res.Success = $true }
                    if ($jobStatus -eq "Warning") { $res.Warning = $true }
                    if ($jobStatus -eq "Failed")  { $res.Failed  = $true }
                }
            } else {
                Add-Log "--------------------------------------"
                Add-Log "Keine abgeschlossenen M365-Jobs in den letzten 24h gefunden."
            }
        } catch { Add-Log "Fehler bei M365: $($_.Exception.Message)" }
    } else {
        Add-Log "--------------------------------------"
        Add-Log "Veeam M365-Modul nicht installiert."
    }

    if ($LogListFull.Count -eq 0) {
        Add-Log "Keine Backup-Daten gefunden."
        $res.Failed = $true
    }

    # =========================================================================
    # LOG KOMPRIMIERUNG BEI BEDARF
    # Stufe 1: Volltext inkl. aller erfolgreichen VMs
    # Stufe 2: erfolgreiche VM-Zeilen raus, Fehlermeldungen bleiben erhalten
    # =========================================================================
    $fullLogText = $LogListFull -join "`n"

    if ($fullLogText.Length -gt 9000) {
        $LogListCompact.Add("--------------------------------------")
        $LogListCompact.Add("HINWEIS: Details erfolgreicher VMs ausgeblendet, da das NinjaOne Limit von 10.000 Zeichen ueberschritten wurde.")
        $res.Log = $LogListCompact -join "`n"
    } else {
        $res.Log = $fullLogText
    }

    return ($res | ConvertTo-Json -Depth 2)
'@

# -----------------------------------------------------------------------------
# EXECUTION CONTROLLER
# -----------------------------------------------------------------------------

$finalResult = $null
$debugLog = ""
$nativeCapable = $false

# Extrahiert das JSON-Objekt aus einer evtl. verrauschten Ausgabe
function Get-JsonPayload {
    param($RawOutput)
    if ($null -eq $RawOutput) { return $null }
    $text = ($RawOutput | Out-String)
    $startIdx = $text.IndexOf('{')
    $endIdx   = $text.LastIndexOf('}')
    if ($startIdx -ge 0 -and $endIdx -gt $startIdx) {
        return $text.Substring($startIdx, $endIdx - $startIdx + 1)
    }
    return $text
}

if ((Get-PSSnapin -Name VeeamPSSnapin -ErrorAction SilentlyContinue) -or (Get-Command Add-PSSnapin -ErrorAction SilentlyContinue)) {
    try {
        if (-not (Get-PSSnapin -Name VeeamPSSnapin -ErrorAction SilentlyContinue)) {
            Add-PSSnapin VeeamPSSnapin -ErrorAction Stop | Out-Null
        }
        if (Get-Command Get-VBRJob -ErrorAction SilentlyContinue) { $nativeCapable = $true }
    } catch {}
}

if (-not $nativeCapable) {
    try {
        if (Get-Module -ListAvailable -Name "Veeam.Backup.PowerShell") {
            Import-Module "Veeam.Backup.PowerShell" -DisableNameChecking -WarningAction SilentlyContinue -ErrorAction Stop | Out-Null
            if (Get-Command Get-VBRJob -ErrorAction SilentlyContinue) { $nativeCapable = $true }
        }
    } catch {}
}

if ($nativeCapable) {
    Write-Host "Modus: Native (Veeam SnapIn/Module erkannt). Fuehre Skript direkt aus."
    $sb = [scriptblock]::Create($VeeamCheckLogic)
    try {
        $jsonRaw  = & $sb
        $jsonText = Get-JsonPayload -RawOutput $jsonRaw
        $finalResult = $jsonText | ConvertFrom-Json
    } catch {
        $debugLog += "Fehler bei nativer Ausfuehrung: $($_.Exception.Message)`n"
    }
} else {
    Write-Host "Modus: Modern. Native Module fehlen/inkompatibel. Suche PowerShell 7..."

    $pwshPath = Get-Command pwsh -ErrorAction SilentlyContinue
    if (-not $pwshPath) {
        $paths = @("C:\Program Files\PowerShell\7\pwsh.exe", "C:\Program Files\PowerShell\7-preview\pwsh.exe")
        foreach ($p in $paths) { if (Test-Path $p) { $pwshPath = $p; break } }
    }

    if ($pwshPath) {
        Write-Host "PowerShell 7 gefunden: $pwshPath"

        $tempScriptPath = Join-Path $env:TEMP "VeeamCheck_$(New-Guid).ps1"
        Set-Content -Path $tempScriptPath -Value $VeeamCheckLogic -Encoding UTF8

        $jsonOutput = & $pwshPath -NoProfile -NonInteractive -File $tempScriptPath 2>$null
        Remove-Item -Path $tempScriptPath -Force -ErrorAction SilentlyContinue

        try {
            $jsonText = Get-JsonPayload -RawOutput $jsonOutput
            $finalResult = $jsonText | ConvertFrom-Json
        } catch {
            $debugLog += "Fehler bei PS7 Ausfuehrung (JSON Parse): $jsonOutput`n"
        }
    } else {
        $debugLog += "CRITICAL: Weder kompatibles Veeam Modul (V11/12) noch PowerShell 7 (V13 Req.) gefunden.`n"
        $debugLog += "Bitte installieren Sie PowerShell 7 Core, falls Sie Veeam 13 nutzen.`n"
    }
}

# -----------------------------------------------------------------------------
# OUTPUT AN NINJA
# -----------------------------------------------------------------------------

if ($finalResult) {
    $hasSuccess = [bool]$finalResult.Success
    $hasWarning = [bool]$finalResult.Warning
    $hasFailed  = [bool]$finalResult.Failed
    $logContent = $finalResult.Log
} else {
    $hasSuccess = $false
    $hasWarning = $false
    $hasFailed  = $true
    $logContent = $debugLog + "`nSkript fehlgeschlagen."
}

if ($logContent.Length -gt 9800) {
    $logContent = $logContent.Substring(0, 9700) + "`n... [WEITERE DATEN ABGESCHNITTEN (10.000 ZEICHEN LIMIT)]"
}

Ninja-Property-Set veeamlog $logContent
Ninja-Property-Set veeamsuccess ([int]$hasSuccess)
Ninja-Property-Set veeamwarning ([int]$hasWarning)
Ninja-Property-Set veeamfail    ([int]$hasFailed)

Write-Host "backupsuccess:" $hasSuccess
Write-Host "backupwarning:" $hasWarning
Write-Host "backupfailed:"  $hasFailed

Write-Host "`nLOG:`n$logContent"
