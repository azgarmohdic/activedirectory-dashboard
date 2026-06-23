<#
.SYNOPSIS
    Tracks the full interactive session lifecycle - Logon, Disconnect,
    Reconnect, and Logoff - in one chronological text log.

.DESCRIPTION
    Each lifecycle stage is read from exactly one authoritative event
    source, so nothing gets double-logged across two different logs
    describing the same moment:

        Logon      - Security log, Event 4624 (as before)
        Logoff     - Security log, Event 4634 - a real, explicit Sign Out
                     only. Filtered the same way as logons (skips the
                     DWM/UMFD session-plumbing accounts, SYSTEM/service
                     principals, and non-interactive logon types).
        Disconnect - Microsoft-Windows-TerminalServices-LocalSessionManager
                     /Operational, Event 24. Fires when an RDP session is
                     left running in the background (window closed,
                     network drop) WITHOUT signing out - this is the case
                     that never produces a 4634, since the session is
                     still alive on the server.
        Reconnect  - same TS-LSM log, Event 25. Fires when a client
                     reattaches to that still-running disconnected session.

    Run via the Scheduled Task set up by Register-LogonAuditTask.ps1,
    which triggers on all four event IDs across both logs.

    Log line format (tab-separated):
        Timestamp  Action  User  Type  Workstation=..  SourceIP=..
    Disconnect/Reconnect rows leave Workstation/SourceIP blank - the RDS
    session log does not carry those fields the way 4624 does.

    MONTHLY ROTATION: rather than one ever-growing file, each line is
    written to C:\Logs\LogonAudit-MM-yyyy.log based on that EVENT's own
    timestamp (not "now") - e.g. everything from June 2026 lands in
    LogonAudit-06-2026.log, July in LogonAudit-07-2026.log, and so on. A
    new file appears automatically the first time an event from a new
    month is logged; nothing needs to be scheduled separately to create
    it. Using the event's own timestamp (rather than the time the script
    happens to run) means a small backlog processed just after midnight
    on the 1st still lands in the correct prior month's file instead of
    bleeding into the new one.

    PREREQUISITES:
      - "Audit Logon Events" Success auditing enabled (for 4624/4634):
            auditpol /set /subcategory:"Logon" /success:enable
      - The TS-LSM Operational log enabled (Register-LogonAuditTask.ps1
        does this for you, but to check manually):
            wevtutil gl "Microsoft-Windows-TerminalServices-LocalSessionManager/Operational"

    RELIABILITY NOTES (added in this revision):
      - BACKLOG-SAFE FETCHING: each log is read through Get-NewAuditEvents,
        which pages the fetch size up automatically if more matching events
        have accumulated since the last run than a single page would
        capture (e.g. the task was delayed, or a login storm produced a
        burst of events). A single capped fetch can silently drop the
        oldest events in a large gap; paging guarantees nothing in the gap
        is skipped, up to a 20,000-event-per-run ceiling -- an extreme
        backlog beyond that closes over a few runs instead of one run
        trying to scan forever.
      - VISIBLE FAILURES: a genuine fetch error (the TS-LSM log disabled,
        access denied, etc.) is written to C:\Logs\LogonAudit.diagnostics.log
        instead of being silently ignored. Only the normal "no matching
        events" case is treated as a non-error.
      - DUPLICATE-EVENT FIX: the de-dupe check for the UAC linked
        full-token/filtered-token pair now allows a small time tolerance
        ($DupWindowSec, default 2s) instead of requiring an exact
        rounded-to-the-second timestamp match. The two halves of that pair
        can straddle a wall-clock second boundary, which let some pairs
        slip through as two separate log lines under the old exact-match
        check.
#>

# ---- EDIT THESE PATHS IF NEEDED ----
$LogDir       = 'C:\Logs'
$LogPrefix    = 'LogonAudit'                              # files are named <LogPrefix>-MM-yyyy.log
$StateFile    = 'C:\Logs\LogonAudit.state.json'
$LegacyMarker = 'C:\Logs\LogonAudit.lastid'                # only used once, to migrate from the earlier (single-file) version of this script
$DiagFile     = 'C:\Logs\LogonAudit.diagnostics.log'       # genuine fetch errors land here, separate from the audit trail itself, so a failure is never silent
$DupWindowSec = 2                                          # tolerance window (seconds) for collapsing the UAC linked full/filtered-token duplicate 4624/4634 pair
# -------------------------------------

if (-not (Test-Path $LogDir)) { New-Item -Path $LogDir -ItemType Directory -Force | Out-Null }

function Get-MonthlyLogPath {
    param([datetime]$Date)
    Join-Path $LogDir ("{0}-{1:MM-yyyy}.log" -f $LogPrefix, $Date)
}

function Write-Diag {
    param([string]$Message)
    $line = "[{0:yyyy-MM-dd HH:mm:ss}] {1}" -f (Get-Date), $Message
    try { Add-Content -Path $DiagFile -Value $line } catch { }
}

# Fetches every new event newer than $LastId from the given log, paging the fetch
# size up automatically if the gap since the last successful run is bigger than a
# single page -- so a delayed or missed task run can never silently truncate the
# backlog the way one capped Get-WinEvent call would. Genuine fetch errors (log
# missing, access denied, etc.) are written to the diagnostics file instead of
# being swallowed -- only the normal "no matching events" case is ignored.
function Get-NewAuditEvents {
    param(
        [string]   $LogName,
        [object[]] $Ids,
        [int64]    $LastId
    )

    if ($LastId -le 0) {
        # No prior checkpoint (fresh install, or first run after this log was newly
        # added) -- intentionally start from recent activity only, not the entire
        # historical log.
        try {
            return @(Get-WinEvent -FilterHashtable @{ LogName = $LogName; Id = $Ids } -MaxEvents 500 -ErrorAction Stop)
        } catch {
            if ($_.Exception.Message -notmatch 'No events were found') { Write-Diag "ERROR fetching $LogName (baseline): $($_.Exception.Message)" }
            return @()
        }
    }

    $fetchSize = 500     # matches the original script's steady-state cost -- only grows if a backlog is detected
    $maxFetch  = 20000   # hard ceiling per run -- even an extreme multi-week gap closes
                          # over a few runs instead of one run trying to scan forever
    while ($true) {
        try {
            $batch = @(Get-WinEvent -FilterHashtable @{ LogName = $LogName; Id = $Ids } -MaxEvents $fetchSize -ErrorAction Stop)
        } catch {
            if ($_.Exception.Message -notmatch 'No events were found') { Write-Diag "ERROR fetching $LogName : $($_.Exception.Message)" }
            return @()
        }

        $newOnes          = @($batch | Where-Object { $_.RecordId -gt $LastId })
        $gapFullyCaptured = ($batch.Count -lt $fetchSize) -or ($newOnes.Count -lt $batch.Count)

        if ($gapFullyCaptured -or $fetchSize -ge $maxFetch) {
            if (-not $gapFullyCaptured) {
                Write-Diag "WARNING: $LogName backlog exceeded $maxFetch events in one run -- captured the most recent $maxFetch, remainder will be picked up on the next run(s)."
            }
            return $newOnes
        }
        $fetchSize = [Math]::Min($fetchSize * 2, $maxFetch)
    }
}

$LogonTypeMap = @{
    2  = 'Interactive (console)'
    7  = 'Unlock'
    10 = 'RemoteInteractive (RDP)'
    11 = 'CachedInteractive'
}

# ---- load persisted state (last RecordId seen per log, + dedupe key/time) ----
$state = @{ SecurityId = 0; TsLsmId = 0; LastKey = ''; LastTime = '' }
if (Test-Path $StateFile) {
    try {
        $loaded = Get-Content $StateFile -Raw | ConvertFrom-Json
        $state.SecurityId = [int64]$loaded.SecurityId
        $state.TsLsmId    = [int64]$loaded.TsLsmId
        $state.LastKey    = [string]$loaded.LastKey
        $state.LastTime   = [string]$loaded.LastTime
    } catch { }
}
elseif (Test-Path $LegacyMarker) {
    # One-time migration from the earlier (logon-only) version of this
    # script, so upgrading doesn't re-log every historical logon as if it
    # were new.
    $raw = @(Get-Content $LegacyMarker -ErrorAction SilentlyContinue)
    if ($raw.Count -ge 1 -and $raw[0] -match '^\d+$') { $state.SecurityId = [int64]$raw[0] }
    if ($raw.Count -ge 2) { $state.LastKey = $raw[1] }
}

$rows = New-Object System.Collections.Generic.List[object]

# ---- Logon (4624) + Logoff (4634) - Security log ----
$secEvents = @(Get-NewAuditEvents -LogName 'Security' -Ids 4624,4634 -LastId $state.SecurityId)

foreach ($evt in $secEvents) {
    $xml  = [xml]$evt.ToXml()
    $data = @{}
    foreach ($n in $xml.Event.EventData.Data) { $data[$n.Name] = $n.'#text' }

    $logonType = 0
    [void][int]::TryParse($data['LogonType'], [ref]$logonType)
    if (-not $LogonTypeMap.ContainsKey($logonType)) { continue }

    $user = $data['TargetUserName']
    if ($user -match '\$$' -or
        $user -in @('ANONYMOUS LOGON', 'SYSTEM', 'LOCAL SERVICE', 'NETWORK SERVICE') -or
        $user -match '^(DWM|UMFD)-\d+$') { continue }

    $rows.Add([ordered]@{
        Time        = $evt.TimeCreated
        Action      = if ($evt.Id -eq 4624) { 'Logon' } else { 'Logoff' }
        User        = "$($data['TargetDomainName'])\$user"
        Type        = $LogonTypeMap[$logonType]
        Workstation = $data['WorkstationName']
        SourceIP    = $data['IpAddress']
    })
}
if ($secEvents.Count -gt 0) { $state.SecurityId = ($secEvents | Sort-Object RecordId | Select-Object -Last 1).RecordId }

# ---- Disconnect (24) + Reconnect (25) - RDS session log ----
$tsEvents = @(Get-NewAuditEvents -LogName 'Microsoft-Windows-TerminalServices-LocalSessionManager/Operational' -Ids 24,25 -LastId $state.TsLsmId)

foreach ($evt in $tsEvents) {
    $xml  = [xml]$evt.ToXml()
    $data = @{}
    foreach ($n in $xml.Event.EventData.Data) { $data[$n.Name] = $n.'#text' }

    $user = $data['User']
    if (-not $user) { continue }

    $rows.Add([ordered]@{
        Time        = $evt.TimeCreated
        Action      = if ($evt.Id -eq 24) { 'Disconnect' } else { 'Reconnect' }
        User        = $user
        Type        = 'RemoteInteractive (RDP)'
        Workstation = ''
        SourceIP    = ''
    })
}
if ($tsEvents.Count -gt 0) { $state.TsLsmId = ($tsEvents | Sort-Object RecordId | Select-Object -Last 1).RecordId }

if ($rows.Count -gt 0) {
    # Written in true chronological order regardless of which log an
    # event came from, so Logon/Disconnect/Reconnect/Logoff read as one
    # timeline rather than batched by source.
    $lastKey  = $state.LastKey
    $lastTime = $null
    if ($state.LastTime) {
        try { $lastTime = [datetime]::Parse($state.LastTime, [System.Globalization.CultureInfo]::InvariantCulture) } catch { $lastTime = $null }
    }

    foreach ($r in ($rows | Sort-Object Time)) {
        $timeStr = $r.Time.ToString('yyyy-MM-dd HH:mm:ss')

        # Collapses the duplicate event Windows itself generates for one
        # physical logon/logoff when UAC is on (a linked full-token +
        # filtered-token pair, same user/action, different LogonId). The two
        # halves of that pair land within a second or two of each other but
        # can straddle a wall-clock second boundary, so this checks a small
        # time tolerance rather than requiring an exact timestamp match.
        $key = "$($r.Action)|$($r.User)|$($r.Type)|$($r.Workstation)|$($r.SourceIP)"
        $isDuplicate = $false
        if ($key -eq $lastKey -and $lastTime) {
            $diffSec = [Math]::Abs(($r.Time - $lastTime).TotalSeconds)
            if ($diffSec -le $DupWindowSec) { $isDuplicate = $true }
        }
        if ($isDuplicate) { continue }
        $lastKey  = $key
        $lastTime = $r.Time

        # Routed by the EVENT's own month, not the current month - so a
        # late-30th/31st backlog processed just after midnight still lands
        # in the correct prior month's file.
        $targetLogFile = Get-MonthlyLogPath -Date $r.Time
        $line = "$timeStr`t$($r.Action)`t$($r.User)`t$($r.Type)`tWorkstation=$($r.Workstation)`tSourceIP=$($r.SourceIP)"
        Add-Content -Path $targetLogFile -Value $line
    }
    $state.LastKey  = $lastKey
    $state.LastTime = if ($lastTime) { $lastTime.ToString('o') } else { '' }
}

($state | ConvertTo-Json -Compress) | Set-Content -Path $StateFile
