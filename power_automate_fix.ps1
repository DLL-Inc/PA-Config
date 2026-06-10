# ============================================================
#  Fix-PowerAutomate-Cleanup.ps1
#  - Disables Power Automate's built-in run file cleanup
#  - Restarts UIFlowService to apply the fix
#  - Installs a Windows Task Scheduler job that deletes run
#    files older than X days every night at midnight
#
#  Requirements : Run as Administrator
# ============================================================

#Requires -RunAsAdministrator

function Write-Step { param($msg) Write-Host "`n== $msg" -ForegroundColor Cyan }
function Write-OK   { param($msg) Write-Host "   [OK] $msg" -ForegroundColor Green }
function Write-Warn { param($msg) Write-Host "   [!!] $msg" -ForegroundColor Yellow }
function Write-Fail { param($msg) Write-Host "   [XX] $msg" -ForegroundColor Red }


# ============================================================
# CONFIGURATION - Edit this section to your needs
# ============================================================

# How many days of run files to keep before deleting
# e.g. 7 = keep last 7 days, delete anything older
$keepDays = 1

# The folder where Power Automate stores run files
$runsFolder = "$env:ProgramData\Microsoft\Power Automate Desktop\Runs"

# Name for the scheduled task
$taskName = "PowerAutomate-RunFiles-Cleanup"


# ============================================================
# PART 1 - APPLY THE REGISTRY FIX
# ============================================================

Write-Step "Applying DisableRunFilesCleanup registry fix"

$regPath = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Power Automate Desktop\Global"

if (!(Test-Path $regPath)) {
    New-Item -Path $regPath -Force | Out-Null
    Write-Warn "Registry key did not exist - created it"
}

Set-ItemProperty -Path $regPath -Name "DisableRunFilesCleanup" -Value 1 -Type DWord
Write-OK "DisableRunFilesCleanup set to 1"


# ============================================================
# PART 2 - RESTART UIFlowService
# ============================================================

Write-Step "Restarting UIFlowService"

$svc = Get-Service -Name "UIFlowService" -ErrorAction SilentlyContinue

if ($null -eq $svc) {
    Write-Warn "UIFlowService not found - Power Automate may not be installed yet"
    Write-Warn "The registry fix is saved and will apply once Power Automate is installed"
} else {
    try {
        Stop-Service -Name "UIFlowService" -Force -ErrorAction Stop
        Write-OK "UIFlowService stopped"
        Start-Sleep -Seconds 2
        Start-Service -Name "UIFlowService" -ErrorAction Stop
        Write-OK "UIFlowService started"
    } catch {
        Write-Fail "Could not restart UIFlowService: $_"
        Write-Warn "Try restarting it manually: net stop UIFlowService && net start UIFlowService"
    }
}


# ============================================================
# PART 3 - CREATE THE NIGHTLY CLEANUP SCHEDULED TASK
# ============================================================
# The task runs as SYSTEM at midnight every day.
# It deletes any run file folders older than $keepDays days.
# A log is written to C:\ODT\Logs\PAD-Cleanup.log each run.
# ============================================================

Write-Step "Creating nightly cleanup scheduled task"

# Make sure log folder exists
New-Item -ItemType Directory -Force -Path "C:\ODT\Logs" | Out-Null

# The cleanup script that the task will execute
# Written as a single-line -Command string to avoid needing a separate .ps1 file
$cleanupScript = @"
`$runsFolder = '$runsFolder'
`$keepDays   = $keepDays
`$logFile    = 'C:\ODT\Logs\PAD-Cleanup.log'
`$cutoff     = (Get-Date).AddDays(-`$keepDays)
`$timestamp  = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

Add-Content `$logFile "`$timestamp - Starting cleanup (keeping last $keepDays days)"

if (!(Test-Path `$runsFolder)) {
    Add-Content `$logFile "`$timestamp - Runs folder not found, nothing to clean"
    exit 0
}

`$folders = Get-ChildItem -Path `$runsFolder -Directory -ErrorAction SilentlyContinue |
    Where-Object { `$_.LastWriteTime -lt `$cutoff }

if (`$folders.Count -eq 0) {
    Add-Content `$logFile "`$timestamp - No folders older than $keepDays days found"
} else {
    foreach (`$folder in `$folders) {
        try {
            Remove-Item -Path `$folder.FullName -Recurse -Force -ErrorAction Stop
            Add-Content `$logFile "`$timestamp - Deleted: `$(`$folder.FullName)"
        } catch {
            Add-Content `$logFile "`$timestamp - FAILED to delete `$(`$folder.FullName): `$_"
        }
    }
    Add-Content `$logFile "`$timestamp - Cleanup complete. Removed `$(`$folders.Count) folder(s)"
}
"@

# Encode the script so it survives being passed to -Command
$encodedScript = [Convert]::ToBase64String(
    [Text.Encoding]::Unicode.GetBytes($cleanupScript)
)

# Build the scheduled task
$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-NonInteractive -NoProfile -EncodedCommand $encodedScript"

$trigger = New-ScheduledTaskTrigger -Daily -At "00:00"

$principal = New-ScheduledTaskPrincipal `
    -UserId "SYSTEM" `
    -LogonType ServiceAccount `
    -RunLevel Highest

$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 30) `
    -RestartCount 2 `
    -RestartInterval (New-TimeSpan -Minutes 5) `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew

# Remove existing task with same name if present
if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    Write-Warn "Existing task removed - will be recreated fresh"
}

Register-ScheduledTask `
    -TaskName  $taskName `
    -Action    $action `
    -Trigger   $trigger `
    -Principal $principal `
    -Settings  $settings `
    -Description "Deletes Power Automate Desktop run files older than $keepDays days. Installed by Fix-PowerAutomate-Cleanup.ps1" `
    | Out-Null

Write-OK "Scheduled task '$taskName' created"
Write-OK "Runs every night at midnight as SYSTEM"
Write-OK "Deletes run files older than $keepDays days"
Write-OK "Cleanup log: C:\ODT\Logs\PAD-Cleanup.log"


# ============================================================
# DONE
# ============================================================
Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  All done! Summary:"                        -ForegroundColor Cyan
Write-Host "  - DisableRunFilesCleanup = 1 (reg fix)"   -ForegroundColor Cyan
Write-Host "  - UIFlowService restarted"                 -ForegroundColor Cyan
Write-Host "  - Nightly cleanup task installed"          -ForegroundColor Cyan
Write-Host "    Keeps last $keepDays days of run files"              -ForegroundColor Cyan
Write-Host "    Runs folder: $runsFolder" -ForegroundColor Cyan
Write-Host "    Log file: C:\ODT\Logs\PAD-Cleanup.log"  -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan