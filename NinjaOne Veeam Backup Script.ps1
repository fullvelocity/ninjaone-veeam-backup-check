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
#    HINWEIS: Diese vier Werte werden aktuell nur eingelesen, aber nirgends
#    ausgewertet - die Felder werden am Ende ohnehin komplett neu geschrieben.
#    Sie stehen hier als Platzhalter, falls spaeter ein Vergleich mit dem
#    Vorlauf gebraucht wird (z.B. "Status hat sich seit gestern geaendert").
#    Wer das nicht braucht, kann die vier Zeilen ersatzlos loeschen.
$successProp = Ninja-Property-Get veeamsuccess
$warningProp = Ninja-Property-Get veeamwarning
$failProp    = Ninja-Property-Get veeamfail
$logProp     = Ninja-Property-Get veeamlog

# -----------------------------------------------------------------------------
# DER CHECK-CODE
#
# Warum steckt die gesamte Logik in einem Here-String statt direkt im Script?
# Veeam V13 laeuft nur noch unter PowerShell 7. NinjaOne startet Scripts aber
# in der Regel mit Windows PowerShell 5.1. Der Block muss deshalb wahlweise
#   a) direkt im aktuellen Prozess ausgefuehrt werden (V10-V12), oder
#   b) in eine temporaere .ps1 geschrieben und mit pwsh.exe gestartet werden (V13).
# Das entscheidet der EXECUTION CONTROLLER weiter unten.
#
# Der Here-String ist EINFACH quotiert (@'...'@). Damit bleiben alle $-Variablen
# im Block unveraendert stehen und werden erst beim Ausfuehren aufgeloest.
# Wichtig: das abschliessende '@ muss in Spalte 1 stehen, sonst bricht das Parsing.
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

    # Jobnamen die bereits gemeldet wurden. Verhindert, dass derselbe Job
    # zweimal im Log landet - relevant seit Abschnitt 2 (Platform-Jobs), falls
    # ein Job wider Erwarten doch schon in Abschnitt 1 aufgetaucht ist.
    $HandledJobNames = @{}

    # Zwei Log-Listen fuer das 10.000-Zeichen-Limit von NinjaOne:
    #   $LogListFull    = alles, inkl. jeder erfolgreich gesicherten VM
    #   $LogListCompact = dasselbe OHNE die Zeilen erfolgreicher VMs
    # Wird der Volltext zu lang, faellt automatisch auf die Kurzfassung zurueck.
    # Deshalb gilt: -IsDetail NUR fuer Zeilen, die man verlieren darf.
    # Fehlermeldungen niemals mit -IsDetail loggen, sonst fehlt beim grossen
    # Kunden genau die Information, wegen der man ins Log schaut.
    function Add-Log {
        param([string]$msg, [switch]$IsDetail)
        $LogListFull.Add($msg)
        if (-not $IsDetail) { $LogListCompact.Add($msg) }
    }

    # Die Veeam-Datenbank antwortet unter Last gelegentlich nicht sofort
    # (Timeouts waehrend laufender Jobs). Deshalb kritische Abfragen wiederholen,
    # statt den kompletten Check mit einem Fehler abzubrechen.
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

    # Bereinigt eine Veeam-Meldung fuer die Ausgabe im NinjaOne-Feld.
    # Veeam liefert Log-Texte teilweise als HTML-Fragmente ("<b>", "<br>",
    # "&quot;") und mit eingebetteten Zeilenumbruechen. Beides wuerde das Log
    # unleserlich machen und unnoetig Zeichen vom 10.000er-Budget fressen.
    # Ergebnis: eine einzeilige, entschaerfte Meldung, hart auf $MaxLen gekappt.
    # MaxLen 400: Veeam-Meldungen sind lang. Der Reinigungshinweis eines
    # Bandlaufwerks braucht allein rund 250 Zeichen, davor steht noch der Titel.
    # Bei 260 waere genau die Handlungsanweisung abgeschnitten worden.
    function Format-Msg {
        param([string]$Text, [int]$MaxLen = 400)
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

    # Erkennt nichtssagende Standard-Abschlussmeldungen.
    # Veeam schreibt bei jedem Job mit Problem eine Zeile wie
    # "Job finished with warning at 03:00" ins Log. Das ist genau die Information,
    # die man ohnehin schon aus dem Status hat. Solche Zeilen werden nicht
    # verworfen (sie koennen Zeitangaben enthalten), aber nach hinten sortiert,
    # damit bei begrenzter Meldungsanzahl die echte Ursache oben steht.
    function Test-IsNoiseMsg {
        param([string]$Text)
        if ([string]::IsNullOrEmpty($Text)) { return $true }
        return ($Text -match '(finished with|completed with|Processing finished|Job finished|abgeschlossen mit)')
    }

    # KERNSTUECK: holt die eigentlichen Warn-/Fehlermeldungen im Klartext.
    #
    # Das Problem: $session.Result liefert nur "Warning" oder "Failed". Die
    # Ursache ("No suitable media found", "Tape is not available") steht
    # ausschliesslich im Session-Log. Die offiziellen Cmdlets geben davon
    # bestenfalls die oberste Meldung zurueck.
    #
    # Deshalb hier eine Kaskade von vier Quellen, von der besten zur schlechtesten.
    # Sobald eine etwas liefert, wird abgebrochen:
    #   1) Logger.GetLog().GetAttentionRecords()  -> alle Warn-/Fehler-Records
    #      Fallback: UpdatedRecords manuell auf EWarning/EFailed filtern,
    #      da GetAttentionRecords() nicht in jeder Veeam-Version existiert.
    #   2) GetDetails()      -> Freitext-Begruendung der Session
    #   3) Info.Reason       -> bei abgebrochenen Sessions gefuellt
    #   4) Description       -> letzter Strohhalm
    #
    # Funktioniert mit Session- UND TaskSession-Objekten, beide haben einen Logger.
    #
    # ACHTUNG: Alle Zugriffe sind in try/catch gekapselt, weil Logger, GetDetails()
    # und Info.Reason INTERNE Veeam-Members sind. Sie sind nicht dokumentiert und
    # koennen mit einem Update verschwinden. Faellt das aus, bleibt der Job-Status
    # korrekt, es fehlen nur die Detailzeilen.
    #
    # Rueckgabe ist bewusst "return ,$out" (mit fuehrendem Komma): ohne das
    # wuerde PowerShell die Liste in Einzelobjekte zerlegen und bei genau einer
    # Meldung einen String statt einer Liste zurueckgeben - .Count waere dann
    # die Stringlaenge statt 1.
    function Get-DetailMessages {
        param($Obj, [int]$Max = 5)

        # Zwei Toepfe: echte Ursachen zuerst, Standardfloskeln danach
        $primary   = New-Object System.Collections.Generic.List[string]
        $secondary = New-Object System.Collections.Generic.List[string]
        $out       = New-Object System.Collections.Generic.List[string]

        if ($null -eq $Obj) { return ,$out }

        # --- 1a) Interner Logger: Core-Sessions und TaskSessions --------------
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

        # --- 1b) .Log-Property: Tape-Sessions und M365 ------------------------
        # Get-VBRTapeBackupSession liefert ein VBRTapeBackupSession-Objekt, das
        # KEIN Logger-Objekt hat, sondern eine fertige Record-Liste unter .Log
        # (laut Veeam-Doku: Progress, RunManually, Log, Initiator, ...).
        # Genau daran lag es, dass beim Tape-Job nur "Warning" ohne Text ankam:
        # 1a lief ins Leere und es gab keine zweite Quelle.
        if (-not $records) {
            try {
                $records = @($Obj.Log | Where-Object {
                    $_ -and $null -ne $_.Title -and "$($_.Status)" -match 'Warning|Failed|Error'
                })
            } catch {}
        }

        # Title = Kurzmeldung, Description = Zusatzinfo. Beide zusammengefuehrt,
        # weil die eigentliche Ursache mal im einen, mal im anderen Feld steht.
        # Dedup ueber .Contains(), da bei mehreren VMs oft dieselbe Warnung faellt.
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

        # Reihenfolge: echte Ursachen zuerst, Standardfloskeln ans Ende.
        # Wird spaeter auf $Max gekuerzt - so ueberlebt die wichtige Meldung.
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

    # Holt die letzten N Log-Eintraege UNABHAENGIG vom Status.
    #
    # Gebraucht fuer Sessions, die noch laufen bzw. haengen: dort gibt es oft
    # noch gar keinen Warn-Record, weil Veeam den Job nicht als fehlgeschlagen
    # sieht - er wartet ja nur. Die entscheidende Zeile ("Waiting for tape
    # 'LTO-004' to be inserted") ist dann ein ganz normaler Info-Eintrag.
    # Get-DetailMessages wuerde hier nichts finden, deshalb dieser Griff auf
    # die letzten Zeilen des laufenden Logs.
    function Get-LastLogRecords {
        param($Obj, [int]$Count = 3)
        $out = New-Object System.Collections.Generic.List[string]
        if ($null -eq $Obj) { return ,$out }

        # Beide Quellen wie in Get-DetailMessages: Logger (Core) und .Log (Tape/M365)
        $recs = @()
        try { $recs = @($Obj.Logger.GetLog().UpdatedRecords) } catch {}
        if ($recs.Count -eq 0) {
            try { $recs = @($Obj.Log | Where-Object { $_ -and $null -ne $_.Title }) } catch {}
        }

        try {
            $start = 0
            if ($recs.Count -gt $Count) { $start = $recs.Count - $Count }
            for ($i = $start; $i -lt $recs.Count; $i++) {
                $t = Format-Msg -Text ([string]$recs[$i].Title)
                if ($t -and -not $out.Contains($t)) { $out.Add($t) }
            }
        } catch {}
        return ,$out
    }

    # Holt zu einer "leichten" Session das vollstaendige Core-Objekt.
    #
    # Mehrere Cmdlets (Get-VBRTapeBackupSession, Get-VBRSession) geben Wrapper
    # zurueck, denen der Logger fehlt. Ueber die Session-Id laesst sich dieselbe
    # Session als vollwertiges Objekt nachladen - der Weg, der beim Proxmox-
    # Abschnitt nachweislich funktioniert. Findet sich nichts Besseres, kommt
    # das Original unveraendert zurueck.
    function Resolve-FullSession {
        param($Session)
        if ($null -eq $Session) { return $null }

        # Hat das Objekt bereits einen brauchbaren Logger? Dann nichts tun.
        try { if ($Session.Logger -and $Session.Logger.GetLog()) { return $Session } } catch {}

        $sid = $null
        try { $sid = $Session.Id } catch {}
        if (-not $sid) { return $Session }

        try {
            $f = Get-VBRBackupSession -Id $sid -ErrorAction SilentlyContinue
            if ($f) { return $f }
        } catch {}
        try {
            $f = [Veeam.Backup.Core.CBackupSession]::Get($sid)
            if ($f) { return $f }
        } catch {}

        return $Session
    }

    # Erkennt interne Kind-Jobs, die Veeam pro VM bzw. pro Copy-Ziel anlegt.
    #
    # [Veeam.Backup.Core.CBackupJob]::GetAll() liefert auch diese Jobs mit,
    # obwohl sie keine eigenstaendigen Jobs sind: sie laufen ausschliesslich als
    # Teil ihres Elternjobs und haben deshalb NIE eine eigene abgeschlossene
    # Session. Ohne diesen Filter flutet das Log mit Zeilen wie
    # "GL-DC-01 Backup | Keine abgeschlossene Session gefunden" - im Praxistest
    # waren das 22 von 24 Jobs.
    #
    # Die Marker sind am Diagnoselauf auf VBR 13.0.2 verifiziert:
    #   - die 22 Pro-VM-Jobs (TypeToString "Proxmox Agent Backup") tragen alle
    #     dieselbe ParentJobId und IsChildJob = True
    #   - die 2 Copy-Kindjobs heissen "Elternjob\Kindjob", haben aber KEINE
    #     ParentJobId - deshalb ist die Namenspruefung noetig
    #   - die 2 echten Jobs (TypeToString "Proxmox Backup") tragen keinen Marker
    function Test-IsChildJob {
        param($Job)
        if ($null -eq $Job) { return $false }

        # a) Kinder von Copy-Jobs heissen "Elternjob\Kindjob"
        try { if ("$($Job.Name)" -match '\\') { return $true } } catch {}

        # b) Gesetzte ParentJobId (leere GUID = kein Elternteil)
        try {
            $parentId = "$($Job.ParentJobId)"
            if ($parentId -and $parentId -ne "00000000-0000-0000-0000-000000000000") { return $true }
        } catch {}

        # c) Explizite Kind-Flags. Welches davon existiert, haengt an der
        #    Version - auf V13 ist es IsChildJob, IsChild ist dort $null.
        try { if ($Job.IsChildJob       -eq $true) { return $true } } catch {}
        try { if ($Job.IsChildWorkerJob -eq $true) { return $true } } catch {}
        try { if ($Job.HasParent        -eq $true) { return $true } } catch {}
        try { if ($Job.IsChild          -eq $true) { return $true } } catch {}

        return $false
    }

    # Schreibt eine Meldungsliste eingerueckt ins Log
    function Add-Messages {
        param($Messages, [string]$Prefix = "    > ")
        if ($null -eq $Messages) { return }
        foreach ($m in $Messages) {
            if ($m) { Add-Log ($Prefix + $m) }
        }
    }

    # Ermittelt den Status eines Objekts robust ueber mehrere Properties.
    #
    # Je nach Job-Typ und Veeam-Version steckt das Ergebnis in .Result, .State
    # oder .Status - und bei Platform-Jobs (Proxmox) ist .Result waehrend des
    # Laufs "None", waehrend .State bereits aussagekraeftig ist. Reihenfolge:
    # Result -> State -> Status -> "Unknown".
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
    # Zwei Wege, je nach Veeam-Generation:
    #   V11+ : Modul "Veeam.Backup.PowerShell"
    #   V10  : altes PSSnapin "VeeamPSSnapin" (gibt es unter PS7 nicht mehr)
    # Ist Get-VBRJob danach immer noch nicht da, laeuft das Script auf einer
    # Maschine ohne Veeam-Konsole und bricht sauber mit Failed ab.
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

    # Dummy-Aufruf mit Zweck: das Laden des Moduls allein reicht nicht, die
    # Assembly Veeam.Backup.Core wird erst durch einen echten Cmdlet-Aufruf in
    # die AppDomain geladen. Ohne das wirft [Veeam.Backup.Core.CBackupJob] in
    # Abschnitt 2 einen "type not found"-Fehler und die Proxmox-Erkennung faellt
    # still aus. Das Cmdlet selbst ist beliebig, es muss nur harmlos sein.
    try { Get-VBRLicenseAutoUpdateStatus -ErrorAction SilentlyContinue | Out-Null } catch {}

    # =========================================================================
    # 1. VM BACKUP JOBS
    # =========================================================================
    try {
        # Ablauf der Pipeline:
        #   1. alle Sessions holen, nur abgeschlossene ("Stopped") behalten -
        #      laufende Jobs haben noch kein Ergebnis und wuerden falsch alarmieren
        #   2. nach Job gruppieren und je Job nur die neueste Session nehmen
        #   3. Jobs ohne aktiven Zeitplan aussortieren
        #
        # ACHTUNG - genau hier fallen Proxmox-Jobs raus: Get-VBRJob kennt sie
        # nicht, $job ist $null, der Filter liefert $false. Das ist kein Bug in
        # dieser Zeile, sondern eine Veeam-Einschraenkung. Behandelt wird das
        # in Abschnitt 2.
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

                # Job als erledigt vormerken, damit Abschnitt 2 ihn nicht erneut meldet
                $HandledJobNames["$jobName"] = $true

                Add-Log "--------------------------------------"
                Add-Log "Backup-Job: $jobName | Status: $jobStatus | Zeitpunkt: $jobTime"

                # Klartext-Ursache auf Job-Ebene - nur wenn es etwas zu erklaeren gibt.
                # Bei "Success" spart man sich den teuren Zugriff auf das Session-Log.
                if ("$jobStatus" -ne "Success") {
                    Add-Messages -Messages (Get-DetailMessages -Obj $session -Max 5) -Prefix "    > "
                }

                # Einzelne VMs des Jobs. Ein Job kann "Success" sein und trotzdem
                # eine VM mit Warnung enthalten - deshalb wird immer durchiteriert.
                $tasks = $session.GetTaskSessions()
                foreach ($task in $tasks) {
                    $vmName   = $task.Name
                    $vmStatus = Get-StatusString -Obj $task

                    if ("$vmStatus" -eq "Success") {
                        # -IsDetail: darf bei Platzmangel wegfallen (siehe Add-Log)
                        Add-Log "    - VM: $vmName | Status: $vmStatus" -IsDetail
                    } else {
                        # Problem-VMs OHNE -IsDetail, damit sie samt Ursache
                        # auch in der gekuerzten Fassung erhalten bleiben
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
    #    DAS PROBLEM
    #    Veeam liefert diese Jobs bewusst NICHT ueber Get-VBRJob zurueck (offiziell
    #    bestaetigt im Veeam-Forum: "Currently, PowerShell is not available for
    #    Proxmox workloads"). Jedes Script, das seine Sessions ueber
    #    "Get-VBRJob -Name" gegenprueft, wirft sie damit unbemerkt weg. Genau
    #    deshalb kamen bei Proxmox-Kunden bisher gar keine Daten an.
    #
    #    DIE LOESUNG - zwei Wege, A zuerst, B als Netz
    #    Weg A: Get-VBRSession -Type PlatformBackupJob      (offiziell, ab V12.3)
    #    Weg B: [Veeam.Backup.Core.CBackupJob]::GetAll()    (intern, aeltere Staende)
    #
    #    Weg B ist von Veeam ausdruecklich NICHT supportet und nur fuer Reporting
    #    freigegeben - niemals Jobs darueber starten/stoppen/loeschen. Sobald
    #    Veeam offizielle Proxmox-Cmdlets liefert, kann Weg B ersatzlos raus.
    # =========================================================================
    try {
        # --- Job-Liste ueber die interne Core-Klasse --------------------------
        # Liefert Name, Typ und Zeitplan. Wird fuer beides gebraucht:
        #   - als Datenquelle in Weg B
        #   - in Weg A nur ergaenzend, um IsScheduleEnabled pruefen zu koennen
        #     (die Session allein sagt nichts darueber aus, ob der Job noch aktiv ist)
        # Schlaegt der Zugriff fehl, bleibt die Liste leer und Weg A traegt allein.
        # Test-IsChildJob filtert die internen Pro-VM- und Copy-Kindjobs weg,
        # die GetAll() ungefragt mitliefert (siehe Funktionskommentar oben).
        $platformJobs = @()
        try {
            $platformJobs = @([Veeam.Backup.Core.CBackupJob]::GetAll() | Where-Object {
                "$($_.TypeToString)" -match 'Proxmox|Nutanix|AHV|oVirt|RHV|OLVM|Scale Computing|Morpheus'
            } | Where-Object {
                -not (Test-IsChildJob -Job $_)
            })
        } catch {}

        # Nachschlagetabelle JobId -> Job, um spaeter je Session den Job zu finden
        $platformJobById = @{}
        foreach ($pj in $platformJobs) {
            try { $platformJobById["$($pj.Id)"] = $pj } catch {}
        }

        # --- Sessions holen: Weg A (offiziell, ab V12.3) ----------------------
        # Get-VBRSession liefert nur "leichte" Session-Objekte (Id, JobId, Status,
        # Zeit) - schnell, aber ohne Logger und ohne GetTaskSessions(). Deshalb
        # hier erst die Kandidaten bestimmen und weiter unten nur diese wenigen
        # als vollstaendige Session nachladen. Andersherum (alles nachladen) waere
        # auf Servern mit langer Historie spuerbar langsam.
        #
        # -ErrorAction Stop ist Absicht: kennt die installierte Version den Typ
        # "PlatformBackupJob" noch nicht, gibt es einen Bindungsfehler, den der
        # catch abfaengt - dann uebernimmt Weg B.
        $platformSessions = @()
        $platformCandidates = @()
        try {
            $rawPlatform = @(Get-VBRSession -Type PlatformBackupJob -ErrorAction Stop)
            if ($rawPlatform.Count -gt 0) {
                # Normalfall: nur abgeschlossene Sessions bewerten.
                # Sonderfall: laeuft der allererste Job gerade noch, gibt es keine
                # einzige "Stopped"-Session - dann lieber den laufenden Stand
                # melden als gar nichts.
                $stopped = @($rawPlatform | Where-Object { "$($_.State)" -eq "Stopped" })
                if ($stopped.Count -eq 0) { $stopped = $rawPlatform }

                # Je Job nur die neueste Session behalten.
                # Gruppiert wird ueber JobId statt JobName, weil umbenannte Jobs
                # sonst als zwei getrennte Jobs erscheinen wuerden.
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
        # Greift bei Versionen ohne "-Type PlatformBackupJob". Hier kommt man vom
        # Job zur Session statt umgekehrt: FindLastSession() ist der direkte Weg,
        # GetByJob() der Ersatz, falls die Methode nicht existiert.
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

        # Merkt sich, welche Platform-Jobs schon eine Zeile bekommen haben.
        # Wird unten gebraucht, um Jobs OHNE Session nachzumelden.
        $reportedPlatformJobs = @{}

        foreach ($session in $platformSessions) {
            if ($null -eq $session) { continue }

            # Passenden Job zur Session suchen (fuer Name, Typ und Zeitplan).
            # Bleibt $pjob leer, arbeiten wir nur mit dem, was die Session hergibt.
            $pjob  = $null
            $jobId = ""
            try { $jobId = "$($session.JobId)" } catch {}
            if ($jobId -and $platformJobById.ContainsKey($jobId)) { $pjob = $platformJobById[$jobId] }

            # Jobname: bevorzugt aus der Session, sonst aus dem Job-Objekt.
            # Je nach Weg (A oder B) ist mal das eine, mal das andere gefuellt.
            $jobName = ""
            try { $jobName = [string]$session.JobName } catch {}
            if ([string]::IsNullOrEmpty($jobName) -and $pjob) { $jobName = [string]$pjob.Name }
            if ([string]::IsNullOrEmpty($jobName)) { $jobName = "<Unbekannt>" }

            # Doppelmeldung vermeiden - einmal gegen Abschnitt 1, einmal gegen
            # doppelte Sessions innerhalb dieses Abschnitts
            if ($HandledJobNames.ContainsKey("$jobName")) { continue }
            if ($reportedPlatformJobs.ContainsKey("$jobName")) { continue }

            # Deaktivierte Jobs ueberspringen - aber NUR wenn wir den Zeitplan
            # wirklich kennen. Ohne $pjob lieber melden als stillschweigend
            # verschlucken; ein Job zu viel im Log ist harmloser als ein
            # uebersehener Ausfall.
            if ($pjob -and ($pjob.IsScheduleEnabled -eq $false)) { continue }

            # Plattform-Bezeichnung fuer die Log-Zeile, z.B. "Proxmox VE-Job: ...".
            # ToHumanReadable() gibt es nur auf vollstaendigen Sessions (Weg A),
            # deshalb TypeToString aus dem Job-Objekt als Ersatz (Weg B).
            $platformName = ""
            try { if ($session.Platform) { $platformName = [string]$session.Platform.ToHumanReadable() } } catch {}
            if ([string]::IsNullOrEmpty($platformName) -and $pjob) {
                try { $platformName = [string]$pjob.TypeToString } catch {}
            }
            # Trim, weil TypeToString teils mit Leerzeichen endet - sonst steht
            # im Log "Proxmox -Job:" statt "Proxmox VE-Job:"
            $platformName = "$platformName".Trim()
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
            # Zwei Wege, weil es vom Session-Typ abhaengt: vollstaendige Core-
            # Sessions koennen GetTaskSessions(), die leichten aus Get-VBRSession
            # nicht - dort muss das Cmdlet Get-VBRTaskSession ran.
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
        # Wichtig fuer den haeufigsten Praxisfall: Proxmox-Job frisch angelegt,
        # Zeitplan falsch gesetzt, noch nie gelaufen. Ohne diese Schleife stuende
        # dazu nichts im Log und der Job faellt schlicht nicht auf.
        #
        # Deckel bei 5 Zeilen: Kindjobs sind zwar oben herausgefiltert, aber falls
        # eine Veeam-Version die Erkennung umgeht, soll das Log nicht zulaufen und
        # dabei die wichtigen Fehlermeldungen aus dem Zeichenlimit draengen.
        # Die Restanzahl wird ausgewiesen, damit nichts stillschweigend verschwindet.
        $noSessionTotal = 0
        $noSessionShown = 0
        foreach ($pj in $platformJobs) {
            $pName = ""
            try { $pName = [string]$pj.Name } catch {}
            if ([string]::IsNullOrEmpty($pName)) { continue }
            if ($reportedPlatformJobs.ContainsKey("$pName")) { continue }
            if ($HandledJobNames.ContainsKey("$pName")) { continue }
            if ($pj.IsScheduleEnabled -eq $false) { continue }

            $noSessionTotal++
            if ($noSessionShown -ge 5) { continue }

            $pType = "Platform"
            try { $pType = "$($pj.TypeToString)".Trim() } catch {}
            if ([string]::IsNullOrEmpty($pType)) { $pType = "Platform" }

            Add-Log "--------------------------------------"
            Add-Log "$pType-Job: $pName | Status: Keine abgeschlossene Session gefunden."
            $noSessionShown++
        }
        if ($noSessionTotal -gt $noSessionShown) {
            Add-Log "[... $($noSessionTotal - $noSessionShown) weitere Platform-Jobs ohne abgeschlossene Session]"
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

                    # Dritter Versuch ueber die Core-Klasse.
                    # Noetig bei Copy-Jobs auf Plattform-Basis (Proxmox): dort
                    # finden beide Cmdlets nichts, weil die Sessions an den
                    # internen Kindjobs "Elternjob\Kindjob" haengen. Ohne das
                    # meldete der Job nur "Success (Job-Status)" ohne Zeitpunkt -
                    # und im Warnfall entsprechend ohne jede Ursache.
                    if (-not $lastCSession) {
                        try {
                            $coreCopyJob = @([Veeam.Backup.Core.CBackupJob]::GetAll() |
                                             Where-Object { "$($_.Name)" -eq "$cJobName" })[0]
                            if ($coreCopyJob) { $lastCSession = $coreCopyJob.FindLastSession() }
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
    #    Der Abschnitt, der am meisten dazugewonnen hat:
    #      - Klartext-Ursachen statt nur "Warning"
    #        (z.B. "No suitable media found", "Tape is not available")
    #      - Erkennung von Sessions, die auf ein Band warten (WaitingTape).
    #        Diese erreichen NIE den Zustand "Stopped" und waren deshalb bisher
    #        voellig unsichtbar - der haeufigste stille Ausfall bei Bandsicherung.
    #      - Fallback auf den Job-Status, wenn gar keine Session existiert.
    #        Vorher wurde in dem Fall ueberhaupt nichts geloggt.
    # =========================================================================
    try {
        if (Get-Command Get-VBRTapeJob -ErrorAction SilentlyContinue) {
            $tapeJobs = Invoke-WithRetry -Command {
                Get-VBRTapeJob -ErrorAction SilentlyContinue | Where-Object {
                    $baseJob = Get-VBRJob -Name $_.Name -ErrorAction SilentlyContinue
                    if ($baseJob) {
                        ($baseJob | Select-Object -First 1).IsScheduleEnabled -eq $true
                    } else {
                        # Regelfall: Tape-Jobs kennt Get-VBRJob nicht (wie bei
                        # Proxmox). Dann das Enabled-Flag des Tape-Jobs selbst
                        # nehmen, statt pauschal jeden Job durchzulassen -
                        # sonst alarmieren dauerhaft deaktivierte Jobs mit.
                        $_.Enabled -ne $false
                    }
                }
            }

            if ($tapeJobs) {
                foreach ($tapeJob in $tapeJobs) {
                    $tapeJobName = $tapeJob.Name

                    # Sessions einmal komplett holen, danach zweimal auswerten
                    # (abgeschlossen / noch laufend). -Job schlaegt je nach
                    # Version fehl, deshalb -Name als zweiter Versuch.
                    $allTapeSessions = @()
                    try {
                        $allTapeSessions = @(Invoke-WithRetry -Command { Get-VBRTapeBackupSession -Job $tapeJob -ErrorAction SilentlyContinue })
                    } catch {}
                    if ($allTapeSessions.Count -eq 0) {
                        try {
                            $allTapeSessions = @(Invoke-WithRetry -Command { Get-VBRTapeBackupSession -Name $tapeJobName -ErrorAction SilentlyContinue })
                        } catch {}
                    }

                    # Auswertung 1: letzter abgeschlossener Lauf -> normaler Status
                    $latestSession = $null
                    try {
                        $latestSession = @($allTapeSessions | Where-Object { "$($_.State)" -eq "Stopped" } | Sort-Object CreationTime -Descending)[0]
                    } catch {}

                    # Auswertung 2: laeuft gerade noch etwas?
                    $activeSession = $null
                    try {
                        $activeSession = @($allTapeSessions | Where-Object { "$($_.State)" -ne "Stopped" } | Sort-Object CreationTime -Descending)[0]
                    } catch {}

                    if ($null -ne $activeSession) {
                        # Nur haengende Zustaende melden, kein normal laufender Job.
                        # Der Zustandsname unterscheidet sich je nach Version
                        # (WaitingTape, WaitingRepository, Idle, Pending), deshalb
                        # ein Muster statt eines festen Vergleichs.
                        $activeState = "$($activeSession.State)"
                        if ($activeState -match 'Wait|Idle|Pending') {
                            $aTime = "<unbekannt>"
                            try { $aTime = $activeSession.CreationTime.ToString("dd.MM.yyyy HH:mm:ss") } catch {}

                            Add-Log "--------------------------------------"
                            Add-Log "Tape-Job: $tapeJobName | Status: WARTET ($activeState) | Start: $aTime"

                            # Ein wartender Job hat meist noch keinen Warn-Record.
                            # Deshalb: erst regulaer versuchen, sonst die letzten
                            # Log-Zeilen nehmen - dort steht die Aufforderung
                            # ("Waiting for tape 'LTO-004' to be inserted").
                            $waitMsgs = @(Get-DetailMessages -Obj $activeSession -Max 3)
                            if ($waitMsgs.Count -eq 0) {
                                $waitMsgs = @(Get-LastLogRecords -Obj $activeSession -Count 3)
                            }
                            Add-Messages -Messages $waitMsgs -Prefix "    > "
                            Add-Log "    > Hinweis: Job wartet auf Benutzeraktion (z.B. Band einlegen / Medium wechseln)."

                            # Bewusst nur Warning, nicht Failed: der Job ist nicht
                            # fehlgeschlagen, er braucht Handarbeit. Wer das haerter
                            # bewerten will, setzt hier zusaetzlich $res.Failed.
                            $res.Warning = $true
                        }
                    }

                    if ($null -ne $latestSession) {
                        $tapeStatus = Get-StatusString -Obj $latestSession
                        $tapeTime   = "<unbekannt>"
                        try { $tapeTime = $latestSession.CreationTime.ToString("dd.MM.yyyy HH:mm:ss") } catch {}

                        Add-Log "--------------------------------------"
                        Add-Log "Tape-Job: $tapeJobName | Status: $tapeStatus | Zeitpunkt: $tapeTime"

                        # HIER kommt die eigentliche Ursache her - genau das, was
                        # vorher fehlte. Max 6 statt 3, weil ein Tape-Job gern
                        # mehrere zusammenhaengende Meldungen produziert
                        # (Medium fehlt -> Pool leer -> Job abgebrochen).
                        if ("$tapeStatus" -ne "Success") {

                            # Reihenfolge der Versuche, vom Wahrscheinlichsten abwaerts:
                            #   1. Wrapper-Objekt direkt (.Log-Property)
                            #   2. dieselbe Session als volles Core-Objekt (.Logger)
                            #   3. letzte Session des Jobs ueber die Core-Klasse
                            #   4. letzte Log-Zeilen ohne Statusfilter
                            # Ein einzelner Weg reicht nicht: welches Objekt
                            # Get-VBRTapeBackupSession zurueckgibt, haengt an der
                            # Veeam-Version.
                            $tapeMsgs = @(Get-DetailMessages -Obj $latestSession -Max 6)

                            if ($tapeMsgs.Count -eq 0) {
                                $fullTapeSession = Resolve-FullSession -Session $latestSession
                                if ($fullTapeSession) {
                                    $tapeMsgs = @(Get-DetailMessages -Obj $fullTapeSession -Max 6)
                                }
                            }

                            if ($tapeMsgs.Count -eq 0) {
                                try {
                                    $coreSession = @([Veeam.Backup.Core.CBackupSession]::GetByJob($tapeJob.Id) | Sort-Object CreationTime -Descending)[0]
                                    if ($coreSession) { $tapeMsgs = @(Get-DetailMessages -Obj $coreSession -Max 6) }
                                } catch {}
                            }

                            # Letzter Ausweg: die letzten Log-Zeilen ungefiltert.
                            # Besser eine ungenaue Spur als eine leere Warnung.
                            if ($tapeMsgs.Count -eq 0) {
                                $tapeMsgs = @(Get-LastLogRecords -Obj $latestSession -Count 4)
                            }

                            if ($tapeMsgs.Count -eq 0) {
                                Add-Log "    > (Keine Detailmeldung im Session-Log auslesbar - Ursache bitte in der Veeam-Konsole pruefen.)"
                            } else {
                                Add-Messages -Messages $tapeMsgs -Prefix "    > "
                            }

                            # Einzelne Objekte des Tape-Jobs (Backups/Dateien).
                            # Oft steckt die Ursache hier statt auf Job-Ebene:
                            # "Backup XY konnte nicht auf Band geschrieben werden".
                            $tTasks = $null
                            try { $tTasks = $latestSession.GetTaskSessions() } catch {}
                            if (-not $tTasks) {
                                try { $tTasks = Get-VBRTaskSession -Session $latestSession -ErrorAction SilentlyContinue } catch {}
                            }
                            if (-not $tTasks) {
                                $fts = Resolve-FullSession -Session $latestSession
                                try { $tTasks = $fts.GetTaskSessions() } catch {}
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
                        # Weder abgeschlossene noch laufende Session.
                        # Diese Konstellation blieb frueher voellig unkommentiert -
                        # der Tape-Job fehlte dann einfach im Log. Jetzt wird
                        # ersatzweise der am Job hinterlegte LastResult gemeldet.
                        # Die elseif-Bedingung verhindert eine zweite Zeile, wenn
                        # oben bereits "WARTET" ausgegeben wurde.
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

                    # Klartext aus dem VBO-Sessionlog.
                    # M365 hat eine eigene API: kein Logger-Objekt, sondern eine
                    # fertige .Log-Liste an der Session. Deshalb hier von Hand
                    # statt ueber Get-DetailMessages.
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

    # Kein einziger Abschnitt hat etwas geliefert -> als Fehler werten.
    # Sonst meldet das Script "alles gut", obwohl es nichts geprueft hat.
    if ($LogListFull.Count -eq 0) {
        Add-Log "Keine Backup-Daten gefunden."
        $res.Failed = $true
    }

    # =========================================================================
    # LOG KOMPRIMIERUNG BEI BEDARF
    #
    # NinjaOne-Felder fassen 10.000 Zeichen. Zwei Stufen:
    #   Stufe 1: Volltext inkl. jeder erfolgreich gesicherten VM
    #   Stufe 2: erfolgreiche VM-Zeilen raus - Job-Header und ALLE
    #            Fehlermeldungen bleiben erhalten
    #
    # Die Schwelle liegt bei 9000, nicht 10.000: der Controller haengt unten
    # noch einen Hinweistext an, und JSON-Escaping kann die Laenge zusaetzlich
    # veraendern. Der Puffer verhindert, dass mitten in einer Meldung gekappt wird.
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
#
# Entscheidet, WIE der Logik-Block oben ausgefuehrt wird:
#   1. Laesst sich Veeam im aktuellen Prozess ansprechen (SnapIn oder Modul)?
#      -> "Native": Block direkt als Scriptblock ausfuehren. Schnellster Weg.
#   2. Sonst -> "Modern": Veeam V13 braucht PowerShell 7, NinjaOne startet aber
#      meist Windows PowerShell 5.1. Der Block wird in eine temporaere .ps1
#      geschrieben und mit pwsh.exe gestartet.
# Beide Wege liefern dasselbe JSON zurueck, die Auswertung darunter ist identisch.
# -----------------------------------------------------------------------------

$finalResult = $null
$debugLog = ""
$nativeCapable = $false

# Schneidet das JSON aus einer moeglicherweise verrauschten Ausgabe heraus.
#
# Hintergrund: Der Logik-Block soll genau einen JSON-String zurueckgeben. In der
# Praxis kann sich aber Fremdausgabe dazwischenmogeln (Veeam-Modul-Banner,
# Warnungen, Fortschrittstext von pwsh). ConvertFrom-Json bricht dann mit
# "Invalid JSON primitive" ab - und das Script meldete komplett "fehlgeschlagen",
# obwohl die Daten da waren. Deshalb: alles von der ersten '{' bis zur letzten
# '}' nehmen und den Rest verwerfen.
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

        # Umweg ueber eine Datei, weil sich der Block als -Command-Argument an
        # der Zitierung zerlegen wuerde. GUID im Namen, damit parallele Laeufe
        # sich nicht gegenseitig die Datei ueberschreiben.
        $tempScriptPath = Join-Path $env:TEMP "VeeamCheck_$(New-Guid).ps1"
        Set-Content -Path $tempScriptPath -Value $VeeamCheckLogic -Encoding UTF8

        # -NoProfile:      Profilskripte koennten Text ausgeben und das JSON stoeren
        # -NonInteractive: nie auf eine Eingabe warten, das Script laeuft unbeaufsichtigt
        # 2>$null:         stderr verwerfen, damit nur das JSON uebrig bleibt
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
#
# Die drei Flags sind bewusst NICHT exklusiv: ein Server kann gleichzeitig
# erfolgreiche und fehlgeschlagene Jobs haben. In NinjaOne deshalb zuerst auf
# veeamfail pruefen, dann auf veeamwarning - sonst uebertoent ein einzelner
# erfolgreicher Job den Ausfall daneben.
#
# ACHTUNG - laeuft das Script nicht (Maschine aus, Agent offline), behalten die
# Felder ihre ALTEN Werte. Ein Ausfall sieht dann im Dashboard weiter gruen aus.
# Dagegen hilft nur eine separate NinjaOne-Condition auf "Device Offline" bzw.
# "Last Contact", das Script selbst kann das nicht abfangen.
# -----------------------------------------------------------------------------

if ($finalResult) {
    $hasSuccess = [bool]$finalResult.Success
    $hasWarning = [bool]$finalResult.Warning
    $hasFailed  = [bool]$finalResult.Failed
    $logContent = $finalResult.Log
} else {
    # Kein verwertbares Ergebnis -> bewusst als Fehler melden, nicht als "unbekannt".
    # Ein stiller Nicht-Lauf ist gefaehrlicher als ein Fehlalarm.
    $hasSuccess = $false
    $hasWarning = $false
    $hasFailed  = $true
    $logContent = $debugLog + "`nSkript fehlgeschlagen."
}

# Letzte Reissleine vor dem 10.000-Zeichen-Limit des Feldes. Greift nur, wenn
# schon die Kurzfassung aus dem Logik-Block zu lang war (sehr viele Jobs mit
# Fehlermeldungen). Der Hinweis macht sichtbar, dass etwas fehlt.
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
