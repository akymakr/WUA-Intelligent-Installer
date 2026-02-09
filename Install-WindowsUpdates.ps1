<#
.SYNOPSIS
    Windows Update Installer with Intelligent Restart Detection
    
.DESCRIPTION
    This script performs Windows Server updates with controlled restart logic.
    
    Process:
    1. Scan for available Windows updates
    2. Download all available updates
    3. Install all downloaded updates
    4. Check if restart is actually needed (intelligent detection)
    5. Schedule restart ONLY if necessary (with delay for graceful shutdown)
    6. Log everything for audit purposes
    
.PARAMETER LogPath
    Where to store log files
    Default: C:\Windows\Logs\Updates
    
.PARAMETER RestartDelayMinutes
    How many minutes to wait before restart (gives applications time to shut down gracefully)
    Default: 120 (2 hours)

.PARAMETER UpdateMode
    Controls which updates to install.
    - Latest: install all available updates (default)
    - Delayed: install updates only after they have been available for a number of days (see DeferDays)
    - SecurityOnly: install only security/critical/definition updates when identifiable

.PARAMETER DeferDays
    Only applies when UpdateMode=Delayed.
    Updates with LastDeploymentChangeTime newer than (Now - DeferDays) will be skipped.
    Default: 14

.PARAMETER StrictDefer
    Only applies when UpdateMode=Delayed.
    If an update does not expose LastDeploymentChangeTime, StrictDefer will skip it instead of installing it.

.PARAMETER ScanOnly
    Only scan and list updates (after filtering). Does not download/install and will not schedule a reboot.

.PARAMETER NoRestart
    Do not schedule a reboot even if restart is detected as required.
    
.EXAMPLE
    .\Install-WindowsUpates.ps1
    
.EXAMPLE
    .\Install-WindowsUpates.ps1 -RestartDelayMinutes 120

.EXAMPLE
    # Only install important security updates
    .\Install-WindowsUpates.ps1 -UpdateMode SecurityOnly

.EXAMPLE
    # Delay updates: only install updates that have been available for 21+ days
    .\Install-WindowsUpates.ps1 -UpdateMode Delayed -DeferDays 21

.EXAMPLE
    # Just check what would be installed (no download/install/reboot)
    .\Install-WindowsUpates.ps1 -UpdateMode SecurityOnly -ScanOnly
    
.AUTHOR
    Infrastructure Team
    
.VERSION
    1.0 - Clear, simple, easy to understand
#>

[CmdletBinding()]
param(
    [string]$LogPath = 'C:\Windows\Logs\Updates',
    [int]$RestartDelayMinutes = 120,

    # When multiple update chains exist (e.g. SSU first, then cumulative),
    # a single scan/install pass may leave additional updates available.
    # MaxPasses allows a controlled re-scan/install loop in the same execution.
    [ValidateRange(1, 5)]
    [int]$MaxPasses = 2,

    # Poll interval used for async download/install progress.
    [ValidateRange(1, 60)]
    [int]$ProgressPollSeconds = 5,

    [ValidateSet('Latest', 'Delayed', 'SecurityOnly')]
    [string]$UpdateMode = 'Delayed',

    [ValidateRange(0, 365)]
    [int]$DeferDays = 28,

    [switch]$StrictDefer,
    [switch]$ScanOnly,
    [switch]$NoRestart,

    # Use shutdown.exe /f when scheduling restart (more reliable for automation,
    # but may forcibly close apps/services).
    [switch]$ForceRestart
)

#region INITIALIZE - Set up logging and variables
# ============================================================================
# This section prepares the script to run:
# - Creates log file in specified directory
# - Initializes tracking variables
# ============================================================================

# Set error handling - continue on errors (don't crash)
$ErrorActionPreference = 'Continue'

# Require elevation (Windows Update COM + reboot scheduling need admin rights)
$IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator
)
if (-not $IsAdmin) {
    Write-Host "[ERROR] This script must be run as Administrator." -ForegroundColor Red
    exit 1
}

# Track when script started (used for final duration report)
$ScriptStartTime = Get-Date

# Track what happened during execution (used in logging summary)
$UpdateCount = 0
$RestartScheduled = $false
$RestartAlreadyScheduled = $false
$RestartScheduledFor = $null

# Allow disabling restart scheduling (useful for reporting-only runs)
$AllowRestartScheduling = (-not $NoRestart) -and (-not $ScanOnly)

# Track overall script outcome (so we can always reach reboot checks + summary)
$ScriptExitCode = 0
$UpdateFlowFailed = $false
$UpdateFlowSkipped = $false

# Variables that may or may not be set depending on failures
$UpdateSession = $null
$SearchResult = $null
$DownloadCollection = $null
$InstallCollection = $null
$InstallResult = $null

# Create log file if directory doesn't exist
if (-not (Test-Path -Path $LogPath)) {
    New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
}

# Generate unique log filename with timestamp
$LogFilename = "UpdateInstall_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$LogFile = Join-Path -Path $LogPath -ChildPath $LogFilename

# Marker file used to remember that THIS script scheduled a reboot
$RestartMarkerPath = Join-Path -Path $LogPath -ChildPath 'RestartScheduled.json'

# Initialize log file with header information
Add-Content -Path $LogFile -Value "================================================================================" -Encoding UTF8
Add-Content -Path $LogFile -Value "Windows Update Installation Log" -Encoding UTF8
Add-Content -Path $LogFile -Value "================================================================================" -Encoding UTF8
Add-Content -Path $LogFile -Value "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Encoding UTF8
Add-Content -Path $LogFile -Value "Computer: $env:COMPUTERNAME" -Encoding UTF8
Add-Content -Path $LogFile -Value "User: $env:USERNAME" -Encoding UTF8
Add-Content -Path $LogFile -Value "PowerShell: $($PSVersionTable.PSVersion)" -Encoding UTF8
Add-Content -Path $LogFile -Value "MaxPasses: $MaxPasses" -Encoding UTF8
Add-Content -Path $LogFile -Value "ProgressPollSeconds: $ProgressPollSeconds" -Encoding UTF8
Add-Content -Path $LogFile -Value "UpdateMode: $UpdateMode" -Encoding UTF8
Add-Content -Path $LogFile -Value "DeferDays: $DeferDays (only for UpdateMode=Delayed)" -Encoding UTF8
Add-Content -Path $LogFile -Value "StrictDefer: $StrictDefer (only for UpdateMode=Delayed)" -Encoding UTF8
Add-Content -Path $LogFile -Value "ScanOnly: $ScanOnly" -Encoding UTF8
Add-Content -Path $LogFile -Value "NoRestart: $NoRestart" -Encoding UTF8
Add-Content -Path $LogFile -Value "ForceRestart: $ForceRestart" -Encoding UTF8
Add-Content -Path $LogFile -Value "" -Encoding UTF8

#endregion

#region HELPER FUNCTIONS - Write to log and console
# ============================================================================
# These functions handle all logging to make the main code cleaner
# ============================================================================

# Function: Write a message to both log file and console with timestamp
function Write-LogMessage {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Message,
        
        [ValidateSet('INFO', 'WARNING', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO'
    )

    # Allow callers to write a blank separator line (useful in logs and avoids breaking scheduled runs)
    if ([string]::IsNullOrEmpty($Message)) {
        Add-Content -Path $LogFile -Value "" -Encoding UTF8
        Write-Host ""
        return
    }
    
    # Create formatted log entry with timestamp
    $Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $LogEntry = "[$Timestamp] [$Level] $Message"
    
    # Write to log file (permanent record)
    Add-Content -Path $LogFile -Value $LogEntry -Encoding UTF8
    
    # Write to console with color (real-time visibility)
    switch ($Level) {
        'INFO'    { Write-Host $LogEntry -ForegroundColor Cyan }
        'SUCCESS' { Write-Host $LogEntry -ForegroundColor Green }
        'WARNING' { Write-Host $LogEntry -ForegroundColor Yellow }
        'ERROR'   { Write-Host $LogEntry -ForegroundColor Red }
    }
}

# Function: Log a section header (makes log easier to read)
function Write-LogSection {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Title
    )
    
    $Line = "=" * 80
    Add-Content -Path $LogFile -Value "" -Encoding UTF8
    Add-Content -Path $LogFile -Value $Line -Encoding UTF8
    Add-Content -Path $LogFile -Value ">>> $Title" -Encoding UTF8
    Add-Content -Path $LogFile -Value $Line -Encoding UTF8
    Write-Host ""
    Write-Host ">>> $Title" -ForegroundColor White -BackgroundColor DarkCyan
}

function Write-LogVerboseMessage {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Message,

        [ValidateSet('INFO', 'WARNING', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO'
    )

    if ($VerbosePreference -ne 'SilentlyContinue') {
        Write-LogMessage -Message $Message -Level $Level
    }
}

function Get-SystemLastBootTime {
    try {
        return (Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop).LastBootUpTime
    }
    catch {
        return $null
    }
}

function Get-RestartMarkerState {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $state = [pscustomobject]@{
        Exists       = $false
        ScheduledFor = $null
        ScheduledAt  = $null
        Reason       = $null
    }

    if (-not (Test-Path -Path $Path)) {
        return $state
    }

    try {
        $raw = Get-Content -Path $Path -Raw -ErrorAction Stop
        $data = $raw | ConvertFrom-Json -ErrorAction Stop
        $state.Exists = $true
        $state.ScheduledFor = if ($data.ScheduledFor) { [datetime]$data.ScheduledFor } else { $null }
        $state.ScheduledAt = if ($data.ScheduledAt) { [datetime]$data.ScheduledAt } else { $null }
        $state.Reason = $data.Reason
    }
    catch {
        # If marker is corrupted, ignore it (do not block updates/restarts)
        $state.Exists = $true
        $state.Reason = "Marker file exists but could not be parsed: $_"
    }

    # Best-effort cleanup: if the system has booted after ScheduledAt, assume reboot happened
    $lastBoot = Get-SystemLastBootTime
    if ($null -ne $lastBoot -and $null -ne $state.ScheduledAt -and $lastBoot -gt $state.ScheduledAt) {
        try { Remove-Item -Path $Path -Force -ErrorAction SilentlyContinue } catch {}
        $state.Exists = $false
        $state.ScheduledFor = $null
        $state.ScheduledAt = $null
        $state.Reason = $null
    }

    # If the marker says a reboot SHOULD have happened already (ScheduledFor in the past)
    # but we haven't actually rebooted (LastBootUpTime not after ScheduledAt), treat the marker as stale.
    # This prevents re-runs from being stuck in "already scheduled" forever when shutdown was canceled.
    if ($state.Exists -and $null -ne $state.ScheduledFor) {
        $staleAfter = $state.ScheduledFor.AddMinutes(30)
        if ((Get-Date) -gt $staleAfter) {
            $hasRebooted = ($null -ne $lastBoot -and $null -ne $state.ScheduledAt -and $lastBoot -gt $state.ScheduledAt)
            if (-not $hasRebooted) {
                try { Remove-Item -Path $Path -Force -ErrorAction SilentlyContinue } catch {}
                $state.Exists = $false
                $state.ScheduledFor = $null
                $state.ScheduledAt = $null
                $state.Reason = $null
            }
        }
    }

    return $state
}

function Set-RestartMarker {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [datetime]$ScheduledFor,
        [Parameter(Mandatory = $true)]
        [string]$Reason
    )

    $payload = [pscustomobject]@{
        ScheduledAt  = (Get-Date).ToString('o')
        ScheduledFor = $ScheduledFor.ToString('o')
        Reason       = $Reason
        Computer     = $env:COMPUTERNAME
    }

    try {
        $payload | ConvertTo-Json -Depth 4 | Set-Content -Path $Path -Encoding UTF8
    }
    catch {
        Write-LogMessage "Could not write restart marker: $_" -Level WARNING
    }
}

function Test-InteractiveSession {
    try {
        # SessionId 0 is typically services / non-interactive.
        $sessionId = (Get-Process -Id $PID -ErrorAction Stop).SessionId
        return ([Environment]::UserInteractive -and $sessionId -ne 0)
    }
    catch {
        return $false
    }
}

function Get-UpdateCategoryNames {
    param(
        [Parameter(Mandatory = $true)]
        $Update
    )

    $names = @()
    try {
        foreach ($c in $Update.Categories) {
            if ($null -ne $c -and -not [string]::IsNullOrWhiteSpace($c.Name)) {
                $names += [string]$c.Name
            }
        }
    }
    catch {
        # Best-effort only
    }

    return ($names | Sort-Object -Unique)
}

function Test-UpdateSelected {
    param(
        [Parameter(Mandatory = $true)]
        $Update,
        [Parameter(Mandatory = $true)]
        [ValidateSet('Latest', 'Delayed', 'SecurityOnly')]
        [string]$Mode,
        [Parameter(Mandatory = $true)]
        [int]$DeferDays,
        [switch]$StrictDefer
    )

    if ($Mode -eq 'Latest') {
        return $true
    }

    $categoryNames = Get-UpdateCategoryNames -Update $Update

    if ($Mode -eq 'SecurityOnly') {
        $msrcSeverity = $null
        try { $msrcSeverity = $Update.MsrcSeverity } catch { $msrcSeverity = $null }

        return (
            ($categoryNames -contains 'Security Updates') -or
            ($categoryNames -contains 'Critical Updates') -or
            ($categoryNames -contains 'Definition Updates') -or
            (-not [string]::IsNullOrWhiteSpace($msrcSeverity))
        )
    }

    # Mode = Delayed
    if ($DeferDays -le 0) {
        return $true
    }

    $deploymentTime = $null
    try { $deploymentTime = [datetime]$Update.LastDeploymentChangeTime } catch { $deploymentTime = $null }

    if ($null -eq $deploymentTime) {
        # Some update types may not expose this property consistently.
        if ($StrictDefer) {
            return $false
        }
        return $true
    }

    $cutoff = (Get-Date).AddDays(-1 * $DeferDays)
    return ($deploymentTime -le $cutoff)
}

function Invoke-WuaScan {
    param(
        [Parameter(Mandatory = $true)]
        $UpdateSession,
        [Parameter(Mandatory = $true)]
        [ValidateSet('Latest', 'Delayed', 'SecurityOnly')]
        [string]$UpdateMode,
        [Parameter(Mandatory = $true)]
        [int]$DeferDays,
        [switch]$StrictDefer
    )

    $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
    $SearchResult = $UpdateSearcher.Search('IsInstalled=0')

    $SelectedUpdates = @()
    $SkippedCount = 0

    foreach ($Update in $SearchResult.Updates) {
        $include = Test-UpdateSelected -Update $Update -Mode $UpdateMode -DeferDays $DeferDays -StrictDefer:$StrictDefer
        if ($include) {
            $SelectedUpdates += $Update
        }
        else {
            $SkippedCount++
        }
    }

    return [pscustomobject]@{
        SearchResult    = $SearchResult
        SelectedUpdates = $SelectedUpdates
        SkippedCount    = $SkippedCount
    }
}

function Invoke-WuaDownload {
    param(
        [Parameter(Mandatory = $true)]
        $UpdateSession,
        [Parameter(Mandatory = $true)]
        [object[]]$Updates,
        [Parameter(Mandatory = $true)]
        [int]$PollSeconds
    )

    $DownloadCollection = New-Object -ComObject 'Microsoft.Update.UpdateColl'

    foreach ($Update in $Updates) {
        try {
            if (-not $Update.EulaAccepted) {
                $Update.AcceptEula()
            }
        }
        catch {
            Write-LogMessage "Could not accept EULA for update: $($Update.Title). Error: $_" -Level WARNING
        }

        [void]$DownloadCollection.Add($Update)
    }

    $UpdateDownloader = $UpdateSession.CreateUpdateDownloader()
    $UpdateDownloader.Updates = $DownloadCollection

    Write-LogMessage "Starting download of $($DownloadCollection.Count) update(s)..." -Level INFO

    $downloadResult = $null
    $usedAsync = $false

    try {
        $job = $UpdateDownloader.BeginDownload($null, $null)
        $usedAsync = $true

        $nextLogPercent = 0
        $isInteractive = Test-InteractiveSession
        while (-not $job.IsCompleted) {
            $progress = $null
            try { $progress = $job.GetProgress() } catch { $progress = $null }

            $pct = $null
            $statusText = $null
            if ($null -ne $progress) {
                try { $pct = [int]$progress.PercentComplete } catch { $pct = $null }
                try {
                    $idx = [int]$progress.CurrentUpdateIndex
                    $current = $DownloadCollection.Item($idx)
                    if ($null -ne $current) {
                        $displayIdx = $idx + 1
                        $statusText = "[$displayIdx/$($DownloadCollection.Count)] $($current.Title)"
                    }
                }
                catch {
                    $statusText = $null
                }
            }

            if ($isInteractive -and $null -ne $pct) {
                $st = if ($statusText) { "$pct% - $statusText" } else { "$pct%" }
                Write-Progress -Activity 'Downloading Windows Updates' -Status $st -PercentComplete $pct
            }

            if ($null -ne $pct -and $pct -ge $nextLogPercent) {
                $logSt = if ($statusText) { "$pct% - $statusText" } else { "$pct%" }
                Write-LogMessage "Download progress: $logSt" -Level INFO
                $nextLogPercent = [Math]::Min(100, ($nextLogPercent + 10))
            }

            Start-Sleep -Seconds $PollSeconds
        }

        if ($isInteractive) {
            Write-Progress -Activity 'Downloading Windows Updates' -Completed
        }

        $downloadResult = $UpdateDownloader.EndDownload($job)
    }
    catch {
        Write-LogMessage "Async download not available/failed; falling back to synchronous Download(). Details: $_" -Level WARNING
        $downloadResult = $UpdateDownloader.Download()
    }

    return [pscustomobject]@{
        DownloadCollection = $DownloadCollection
        DownloadResult     = $downloadResult
        UsedAsync          = $usedAsync
    }
}

function Invoke-WuaInstall {
    param(
        [Parameter(Mandatory = $true)]
        $UpdateSession,
        [Parameter(Mandatory = $true)]
        $DownloadCollection,
        [Parameter(Mandatory = $true)]
        [int]$PollSeconds
    )

    $InstallCollection = New-Object -ComObject 'Microsoft.Update.UpdateColl'
    foreach ($Update in $DownloadCollection) {
        if ($Update.IsDownloaded) {
            [void]$InstallCollection.Add($Update)
        }
    }

    if ($InstallCollection.Count -eq 0) {
        throw "No downloaded updates to install"
    }

    $UpdateInstaller = $UpdateSession.CreateUpdateInstaller()
    $UpdateInstaller.Updates = $InstallCollection

    Write-LogMessage "Starting installation of $($InstallCollection.Count) update(s)..." -Level INFO
    Write-LogMessage "This may take several minutes..." -Level INFO

    $installResult = $null
    $usedAsync = $false

    try {
        $job = $UpdateInstaller.BeginInstall($null, $null)
        $usedAsync = $true

        $nextLogPercent = 0
        $isInteractive = Test-InteractiveSession
        while (-not $job.IsCompleted) {
            $progress = $null
            try { $progress = $job.GetProgress() } catch { $progress = $null }

            $pct = $null
            $statusText = $null
            if ($null -ne $progress) {
                try { $pct = [int]$progress.PercentComplete } catch { $pct = $null }
                try {
                    $idx = [int]$progress.CurrentUpdateIndex
                    $current = $InstallCollection.Item($idx)
                    if ($null -ne $current) {
                        $displayIdx = $idx + 1
                        $statusText = "[$displayIdx/$($InstallCollection.Count)] $($current.Title)"
                    }
                }
                catch {
                    $statusText = $null
                }
            }

            if ($isInteractive -and $null -ne $pct) {
                $st = if ($statusText) { "$pct% - $statusText" } else { "$pct%" }
                Write-Progress -Activity 'Installing Windows Updates' -Status $st -PercentComplete $pct
            }

            if ($null -ne $pct -and $pct -ge $nextLogPercent) {
                $logSt = if ($statusText) { "$pct% - $statusText" } else { "$pct%" }
                Write-LogMessage "Install progress: $logSt" -Level INFO
                $nextLogPercent = [Math]::Min(100, ($nextLogPercent + 10))
            }

            Start-Sleep -Seconds $PollSeconds
        }

        if ($isInteractive) {
            Write-Progress -Activity 'Installing Windows Updates' -Completed
        }

        $installResult = $UpdateInstaller.EndInstall($job)
    }
    catch {
        Write-LogMessage "Async install not available/failed; falling back to synchronous Install(). Details: $_" -Level WARNING
        $installResult = $UpdateInstaller.Install()
    }

    return [pscustomobject]@{
        InstallCollection = $InstallCollection
        InstallResult     = $installResult
        UsedAsync         = $usedAsync
    }
}

#endregion

#region SECTION 1-3 - SCAN/DOWNLOAD/INSTALL (MULTI-PASS)
# ============================================================================
# Multi-pass flow to handle rare dependency chains without requiring a second
# scheduled run. Stops early when reboot is required.
# ============================================================================

Write-LogSection "WINDOWS UPDATE EXECUTION"

try {
    Write-LogMessage "Creating Windows Update COM object..." -Level INFO
    $UpdateSession = New-Object -ComObject 'Microsoft.Update.Session'
}
catch {
    Write-LogMessage "ERROR: Failed to create Windows Update COM object" -Level ERROR
    Write-LogMessage "Error details: $_" -Level ERROR
    $UpdateFlowFailed = $true
    $ScriptExitCode = 1
}

$TotalUpdatesAttempted = 0
$TotalUpdatesFound = 0
$TotalPassesExecuted = 0

if (-not $UpdateFlowFailed) {
    for ($pass = 1; $pass -le $MaxPasses; $pass++) {
        $TotalPassesExecuted = $pass

        Write-LogSection ("PASS {0} of {1} - SCAN" -f $pass, $MaxPasses)
        try {
            Write-LogMessage "Searching for uninstalled updates (this may take a moment)..." -Level INFO
            Write-LogMessage "Applying filter: UpdateMode=$UpdateMode" -Level INFO
            if ($UpdateMode -eq 'Delayed') {
                Write-LogMessage "Delayed mode: DeferDays=$DeferDays, StrictDefer=$StrictDefer" -Level INFO
            }

            $scan = Invoke-WuaScan -UpdateSession $UpdateSession -UpdateMode $UpdateMode -DeferDays $DeferDays -StrictDefer:$StrictDefer
            $SearchResult = $scan.SearchResult
            $SelectedUpdates = $scan.SelectedUpdates
            $SkippedCount = $scan.SkippedCount

            $UpdateCount = $SelectedUpdates.Count

            if ($UpdateCount -gt 0) {
                $TotalUpdatesFound += $UpdateCount
            }

            if ($UpdateCount -eq 0) {
                if ($pass -eq 1) {
                    Write-LogMessage "No updates matched the selected filter (skipped: $SkippedCount)." -Level SUCCESS
                    $UpdateFlowSkipped = $true
                }
                else {
                    Write-LogMessage "No additional updates found after pass $($pass - 1)." -Level SUCCESS
                }
                break
            }

            Write-LogMessage "Found $UpdateCount update(s) after filtering (skipped: $SkippedCount):" -Level INFO
            foreach ($Update in $SelectedUpdates) {
                if ($UpdateMode -eq 'Latest') {
                    Write-LogMessage "  [!] $($Update.Title)" -Level INFO
                }
                else {
                    $cats = (Get-UpdateCategoryNames -Update $Update) -join ', '
                    $catsText = if ([string]::IsNullOrWhiteSpace($cats)) { 'Unknown' } else { $cats }
                    Write-LogMessage "  [!] $($Update.Title) (Categories: $catsText)" -Level INFO
                }
            }
        }
        catch {
            Write-LogMessage "ERROR: Failed to scan for updates" -Level ERROR
            Write-LogMessage "Error details: $_" -Level ERROR
            $UpdateFlowFailed = $true
            $ScriptExitCode = 1
            break
        }

        if ($ScanOnly) {
            Write-LogMessage "ScanOnly requested: skipping download/install phases." -Level INFO
            $UpdateFlowSkipped = $true
            break
        }

        Write-LogSection ("PASS {0} of {1} - DOWNLOAD" -f $pass, $MaxPasses)
        try {
            $download = Invoke-WuaDownload -UpdateSession $UpdateSession -Updates $SelectedUpdates -PollSeconds $ProgressPollSeconds
            $DownloadCollection = $download.DownloadCollection
            $DownloadResult = $download.DownloadResult

            # ResultCode: 0=NotStarted, 1=InProgress, 2=Succeeded, 3=SucceededWithErrors, 4=Failed, 5=Aborted
            if ($DownloadResult.ResultCode -eq 2 -or $DownloadResult.ResultCode -eq 3) {
                Write-LogMessage "Download successful (result code: $($DownloadResult.ResultCode))" -Level SUCCESS
            }
            else {
                Write-LogMessage "Download failed with result code: $($DownloadResult.ResultCode)" -Level ERROR
                $UpdateFlowFailed = $true
                $ScriptExitCode = 1
                break
            }
        }
        catch {
            Write-LogMessage "FATAL ERROR: Failed to download updates" -Level ERROR
            Write-LogMessage "Error details: $_" -Level ERROR
            $UpdateFlowFailed = $true
            $ScriptExitCode = 1
            break
        }

        Write-LogSection ("PASS {0} of {1} - INSTALL" -f $pass, $MaxPasses)
        try {
            $install = Invoke-WuaInstall -UpdateSession $UpdateSession -DownloadCollection $DownloadCollection -PollSeconds $ProgressPollSeconds
            $InstallCollection = $install.InstallCollection
            $InstallResult = $install.InstallResult
            $TotalUpdatesAttempted += $InstallCollection.Count

            # Log per-update results for troubleshooting
            for ($i = 0; $i -lt $InstallCollection.Count; $i++) {
                $u = $InstallCollection.Item($i)
                $r = $InstallResult.GetUpdateResult($i)
                Write-LogMessage "Update result: [$($r.ResultCode)] $($u.Title) (HResult: 0x$('{0:X8}' -f ($r.HResult -band 0xFFFFFFFF)))" -Level INFO
            }

            # ResultCode: 0=NotStarted, 1=InProgress, 2=Succeeded, 3=SucceededWithErrors, 4=Failed, 5=Aborted
            if ($InstallResult.ResultCode -eq 2 -or $InstallResult.ResultCode -eq 3) {
                Write-LogMessage "Installation successful (result code: $($InstallResult.ResultCode))" -Level SUCCESS
                if ($InstallResult.RebootRequired) {
                    Write-LogMessage "Installer indicates a reboot is required." -Level WARNING
                    break
                }
            }
            else {
                Write-LogMessage "Installation failed with result code: $($InstallResult.ResultCode)" -Level ERROR
                $UpdateFlowFailed = $true
                if ($ScriptExitCode -eq 0) { $ScriptExitCode = 1 }
                break
            }
        }
        catch {
            Write-LogMessage "FATAL ERROR: Failed to install updates" -Level ERROR
            Write-LogMessage "Error details: $_" -Level ERROR
            $UpdateFlowFailed = $true
            if ($ScriptExitCode -eq 0) { $ScriptExitCode = 1 }
            break
        }

        if ($pass -lt $MaxPasses) {
            Write-LogMessage "Pass $pass completed without reboot requirement. Will re-scan once to catch dependency-chained updates." -Level INFO
        }
    }
}

#endregion

#region SECTION 4 - DETECT IF RESTART NEEDED
# ============================================================================
# IMPORTANT DESIGN NOTE
# Follow Windows Update's own reboot requirement signals.
# Do NOT force/schedule reboots based on other "pending" indicators such as:
# - PendingFileRenameOperations
# - CBS RebootPending
# - UpdateExeVolatile
# Those may indicate a reboot is recommended or may block some future installs,
# but the decision to reboot belongs to the operator.
# ============================================================================

Write-LogSection "CHECKING IF RESTART IS REQUIRED"

Write-LogMessage "Checking Windows Update reboot requirement signals..." -Level INFO

# Initialize restart detection
$RestartNeeded = $false
$RestartReasons = @()

# CHECK #1: Installer result (most reliable when available)
if ($null -ne $InstallResult -and $InstallResult.RebootRequired) {
    Write-LogMessage "Checking installer reboot requirement..." -Level INFO
    Write-LogMessage " [!] Installer RebootRequired is TRUE" -Level WARNING
    $RestartNeeded = $true
    $RestartReasons += "Installer reported reboot required"
}

# CHECK #2: Windows Update automatic restart flag (this is a subkey, not a property)
Write-LogMessage "Checking Windows Update RebootRequired key..." -Level INFO
try {
    $RegistryPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
    if (Test-Path -Path $RegistryPath) {
        Write-LogMessage " [!] RebootRequired flag found" -Level WARNING
        $RestartNeeded = $true
        $RestartReasons += "Windows Update marked restart as required"
    }
    else {
        Write-LogMessage " [o] No RebootRequired flag" -Level INFO
    }
}
catch {
    Write-LogMessage " [!] Could not check (not critical): $_" -Level WARNING
}

# Additional indicators (INFORMATIONAL ONLY)
# These are logged for operator awareness but do not trigger reboot scheduling.
$InformationalRebootIndicators = @()

# INFO: Pending file rename operations (files to replace at boot)
Write-LogVerboseMessage "Checking PendingFileRenameOperations (informational only)..." -Level INFO
try {
    $RegistryPath = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager'
    $RegistryKey = Get-ItemProperty -Path $RegistryPath -Name 'PendingFileRenameOperations' -ErrorAction SilentlyContinue
    
    if ($RegistryKey -and $RegistryKey.PendingFileRenameOperations) {
        Write-LogVerboseMessage " [!] Pending file operations found (operator decision to reboot)" -Level WARNING
        $InformationalRebootIndicators += "PendingFileRenameOperations present"
    }
    else {
        Write-LogVerboseMessage " [o] No pending file operations" -Level INFO
    }
}
catch {
    Write-LogVerboseMessage " [!] Could not check (not critical): $_" -Level WARNING
}

# INFO: Component-Based Servicing pending reboot
Write-LogVerboseMessage "Checking CBS RebootPending (informational only)..." -Level INFO
try {
    $RegistryPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending'
    $RegistryKey = Get-Item -Path $RegistryPath -ErrorAction SilentlyContinue
    
    if ($RegistryKey) {
        Write-LogVerboseMessage " [!] CBS reboot pending flag found (operator decision to reboot)" -Level WARNING
        $InformationalRebootIndicators += "CBS RebootPending present"
    }
    else {
        Write-LogVerboseMessage " [o] No CBS reboot pending" -Level INFO
    }
}
catch {
    Write-LogVerboseMessage " [!] Could not check (not critical): $_" -Level WARNING
}

# INFO: UpdateExeVolatile (can block some future installs)
Write-LogVerboseMessage "Checking UpdateExeVolatile (informational only)..." -Level INFO
try {
    $RegistryPath = 'HKLM:\SOFTWARE\Microsoft\Updates'
    $RegistryKey = Get-ItemProperty -Path $RegistryPath -Name 'UpdateExeVolatile' -ErrorAction SilentlyContinue
    if ($RegistryKey -and $RegistryKey.UpdateExeVolatile -ne 0) {
        Write-LogVerboseMessage " [!] UpdateExeVolatile set (value: $($RegistryKey.UpdateExeVolatile)) (operator decision to reboot)" -Level WARNING
        $InformationalRebootIndicators += "UpdateExeVolatile non-zero"
    }
    else {
        Write-LogVerboseMessage " [o] UpdateExeVolatile not set" -Level INFO
    }
}
catch {
    Write-LogVerboseMessage " [!] Could not check (not critical): $_" -Level WARNING
}

# Summary of restart detection
Write-LogMessage "" -Level INFO

if ($RestartNeeded) {
    Write-LogMessage "RESULT: Restart IS REQUIRED" -Level WARNING
    Write-LogMessage "Reason(s):" -Level WARNING
    foreach ($Reason in $RestartReasons) {
        Write-LogMessage " [!] $Reason" -Level WARNING
    }
}
else {
    Write-LogMessage "RESULT: Restart NOT REQUIRED (per Windows Update)" -Level SUCCESS
    Write-LogMessage "Updates installed successfully without Windows Update requiring restart" -Level SUCCESS
    if ($InformationalRebootIndicators.Count -gt 0) {
        Write-LogVerboseMessage "Informational: system indicates pending operations (operator may choose to reboot):" -Level INFO
        foreach ($i in $InformationalRebootIndicators) {
            Write-LogVerboseMessage " [i] $i" -Level INFO
        }
    }
}

#endregion

#region SECTION 5 - SCHEDULE RESTART IF NEEDED
# ============================================================================
# If restart is needed, schedule it with a delay to allow graceful shutdown
# If not needed, we're done!
# ============================================================================

Write-LogSection "RESTART ACTION"

# Detect if THIS script already scheduled a restart previously (re-run safety)
$markerState = Get-RestartMarkerState -Path $RestartMarkerPath
if ($markerState.Exists) {
    $RestartAlreadyScheduled = $true
    $RestartScheduledFor = $markerState.ScheduledFor
    $whenText = if ($RestartScheduledFor) { $RestartScheduledFor.ToString('yyyy-MM-dd HH:mm:ss') } else { 'Unknown' }
    Write-LogMessage "Existing restart marker detected (scheduled for: $whenText)." -Level WARNING
    if ($markerState.Reason) {
        Write-LogMessage "Marker reason: $($markerState.Reason)" -Level INFO
    }
}

if ($RestartNeeded) {
    if (-not $AllowRestartScheduling) {
        Write-LogMessage "Restart is required, but restart scheduling is disabled (ScanOnly/NoRestart)." -Level WARNING
        Write-LogMessage "No restart will be scheduled by this run." -Level WARNING
    }
    else {
    Write-LogMessage "Scheduling restart..." -Level WARNING
    try {
        if ($RestartAlreadyScheduled) {
            Write-LogMessage "A restart is already scheduled (marker present). Not scheduling again." -Level WARNING
            $RestartScheduled = $true
        }
        else {
        # Calculate when restart will happen
        $RestartTime = (Get-Date).AddMinutes($RestartDelayMinutes)
        
        # Schedule restart using shutdown.exe
        # Use the call operator (&) so PowerShell handles quoting for args with spaces (e.g. /c comment)
        $ShutdownTimeoutSeconds = [Math]::Max(0, ($RestartDelayMinutes * 60))
        $ShutdownComment = 'Windows scheduled maintenance restart. Updates have been installed.'

        $shutdownArgs = @('/r', '/t', $ShutdownTimeoutSeconds, '/c', $ShutdownComment)
        if ($ForceRestart) {
            $shutdownArgs += '/f'
        }

        Write-LogMessage ("Running: shutdown.exe {0}" -f ($shutdownArgs -join ' ')) -Level INFO
        & shutdown.exe @shutdownArgs

        # 0 = scheduled OK
        # 1190 = a shutdown has already been scheduled (treat as already scheduled, do not fail)
        if ($LASTEXITCODE -eq 0) {
            $RestartScheduled = $true
            Set-RestartMarker -Path $RestartMarkerPath -ScheduledFor $RestartTime -Reason 'Scheduled by script'
        }
        elseif ($LASTEXITCODE -eq 1190) {
            Write-LogMessage "shutdown.exe reports a shutdown is already scheduled (exit code 1190)." -Level WARNING
            $RestartScheduled = $true
            # Record what we observed so re-runs don't keep trying
            Set-RestartMarker -Path $RestartMarkerPath -ScheduledFor $RestartTime -Reason 'shutdown.exe indicated already scheduled (1190)'
        }
        else {
            throw "shutdown.exe failed with exit code: $LASTEXITCODE"
        }
        
        # Log the restart scheduling
        Write-LogMessage "Restart scheduled successfully" -Level WARNING
        Write-LogMessage "Restart will occur at: $($RestartTime.ToString('yyyy-MM-dd HH:mm:ss'))" -Level WARNING
        Write-LogMessage "Delay: $RestartDelayMinutes minute(s)" -Level WARNING

        # Best-effort notification to interactive users.
        # Note: shutdown UI/toast often does NOT appear when running in Session 0 (scheduled task/WinRM).
        try {
            $notifyText = "Windows will restart at $($RestartTime.ToString('yyyy-MM-dd HH:mm:ss')) (in $RestartDelayMinutes minute(s)) for update maintenance."
            & msg.exe * /time:60 $notifyText 2>$null
            if (Test-InteractiveSession) {
                Write-LogMessage "User notification attempted (msg.exe)." -Level INFO
            }
        }
        catch {
            Write-LogMessage "Could not send user notification (not critical): $_" -Level WARNING
        }
        }
    }
    catch {
        Write-LogMessage "FATAL ERROR: Failed to schedule restart" -Level ERROR
        Write-LogMessage "Error details: $_" -Level ERROR
        if ($ScriptExitCode -eq 0) { $ScriptExitCode = 1 }
    }
    }
}
else {
    Write-LogMessage "No restart needed. Updates installed cleanly." -Level SUCCESS
    Write-LogMessage "Server will continue running without interruption." -Level SUCCESS
    if ($RestartAlreadyScheduled) {
        Write-LogMessage "Note: A restart appears to already be scheduled (marker present)." -Level WARNING
    }
}

#endregion

#region SECTION 6 - FINAL SUMMARY
# ============================================================================
# Log final status and execution time
# ============================================================================

Write-LogSection "EXECUTION SUMMARY"

# Calculate how long the script ran
$ExecutionDuration = (Get-Date) - $ScriptStartTime
$DurationFormatted = $ExecutionDuration.ToString('hh\:mm\:ss')

# Write final summary to log
Write-LogMessage "Execution completed: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Level SUCCESS
Write-LogMessage "Total execution time: $DurationFormatted" -Level INFO
if ($ScanOnly) {
    Write-LogMessage "Updates found (ScanOnly): $TotalUpdatesFound" -Level INFO
}
else {
    Write-LogMessage "Updates attempted (install): $TotalUpdatesAttempted" -Level INFO
    Write-LogMessage "Updates found (all passes): $TotalUpdatesFound" -Level INFO
}
Write-LogMessage "Passes executed: $TotalPassesExecuted" -Level INFO
Write-LogMessage "Restart scheduled: $RestartScheduled" -Level INFO
Write-LogMessage "Restart already scheduled (marker): $RestartAlreadyScheduled" -Level INFO

# Write to console as well
Write-Host ""
Write-Host "Script Summary:" -ForegroundColor White -BackgroundColor DarkCyan
Write-Host "  Duration: $DurationFormatted" -ForegroundColor Cyan
if ($ScanOnly) {
    Write-Host "  Updates Found:     $TotalUpdatesFound" -ForegroundColor Cyan
}
else {
    Write-Host "  Updates Attempted: $TotalUpdatesAttempted" -ForegroundColor Cyan
    Write-Host "  Updates Found:     $TotalUpdatesFound" -ForegroundColor Cyan
}
Write-Host "  Passes:   $TotalPassesExecuted" -ForegroundColor Cyan
Write-Host "  Restart:  $RestartScheduled" -ForegroundColor Cyan
Write-Host ""
Write-Host "Log file: $LogFile" -ForegroundColor Green
Write-Host ""

# Create/overwrite a shortcut in the script directory for quick log access.
# Shortcut filename is LastExecute.log.lnk (typically displayed as LastExecute.log when extensions are hidden).
try {
    $scriptPath = if ($PSCommandPath) { $PSCommandPath } else { $MyInvocation.MyCommand.Path }
    if (-not [string]::IsNullOrWhiteSpace($scriptPath)) {
        $scriptDir = Split-Path -Path $scriptPath -Parent
        if (Test-Path -Path $scriptDir) {
            $shortcutPath = Join-Path -Path $scriptDir -ChildPath 'LastExecute.log.lnk'
            $wsh = New-Object -ComObject WScript.Shell
            $sc = $wsh.CreateShortcut($shortcutPath)
            $sc.TargetPath = $LogFile
            $sc.WorkingDirectory = $LogPath
            $sc.Description = 'Shortcut to the most recent Windows Update install log.'
            $sc.Save()
            Write-LogMessage "Created/updated shortcut: $shortcutPath" -Level INFO
        }
    }
}
catch {
    Write-LogMessage "Could not create log shortcut in script directory (not critical): $_" -Level WARNING
}

# Decide final exit code (for automation):
# 0    = success, no reboot required
# 3010 = success, reboot required (common in deployment tooling)
$FinalExitCode = 0
if ($ScriptExitCode -ne 0) {
    $FinalExitCode = $ScriptExitCode
}
elseif ($RestartNeeded -or $RestartScheduled -or $RestartAlreadyScheduled) {
    $FinalExitCode = 3010
}

# Interactive behavior:
# Keep the PowerShell window open until the machine reboots/shuts down.
# Any key will close the window immediately. If a reboot was scheduled by this run,
# allow the user to press A to abort the reboot (shutdown.exe /a).
if (Test-InteractiveSession) {
    Write-Host "" 
    if ($RestartScheduled -or $RestartAlreadyScheduled) {
        Write-Host "A reboot/shutdown may be scheduled." -ForegroundColor Yellow
        Write-Host "Press 'A' to abort the scheduled reboot, or press ANY other key to close this window." -ForegroundColor Yellow
        Write-Host "If you do nothing, this window will stay open until shutdown/restart." -ForegroundColor Cyan
    }
    else {
        Write-Host "Press ANY key to close this window." -ForegroundColor Yellow
        Write-Host "If you do nothing, this window will stay open until shutdown/restart." -ForegroundColor Cyan
    }

    while (-not [Console]::KeyAvailable) {
        Start-Sleep -Seconds 1
    }

    $key = [Console]::ReadKey($true)
    if (($RestartScheduled -or $RestartAlreadyScheduled) -and ($key.Key -eq [ConsoleKey]::A)) {
        try {
            Write-LogMessage "User requested abort of scheduled reboot (shutdown.exe /a)." -Level WARNING
            & shutdown.exe /a | Out-Null
            Write-LogMessage "Abort request sent." -Level INFO
            # Keep window open after abort; any subsequent key closes.
            Write-Host "Reboot abort requested. Press ANY key to close this window." -ForegroundColor Yellow
            while (-not [Console]::KeyAvailable) {
                Start-Sleep -Seconds 1
            }
            [void][Console]::ReadKey($true)
        }
        catch {
            Write-LogMessage "Failed to abort scheduled reboot: $_" -Level WARNING
        }
    }
}

exit $FinalExitCode

#endregion

