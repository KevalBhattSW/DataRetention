<#
.SYNOPSIS
    Launches the Azure DevOps agent (run.cmd) as a detached process, independent
    of the PowerShell session that starts it, and logs the process tree so you
    can verify the agent is NOT a child of your interactive console.

.DESCRIPTION
    Root cause traced: Agent.Listener.exe was a direct child of the interactive
    PowerShell process. Closing that console, hitting Ctrl+C, an RDP session
    disconnect affecting the window, or the shell crashing would silently kill
    the agent with no crash log and no OS-level event - matching the
    "agent not contactable" ADO failures.

    This script uses Start-Process to break that parent/child relationship.
    The agent still runs in the current interactive desktop session (required
    for the legacy Office COM automation), it just won't die if THIS console
    does.

.NOTES
    Run this from an interactive PowerShell session (RDP/console), same as
    before - just use this script instead of calling run.cmd directly.
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$AgentPath = "C:\path\to\agent",   # <-- UPDATE to your actual agent install path

    [Parameter(Mandatory = $false)]
    [string]$LogDir = "C:\logs"
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path $LogDir)) {
    New-Item -Path $LogDir -ItemType Directory -Force | Out-Null
}

$launchLog = Join-Path $LogDir "agent-launch-history.csv"
$runCmdPath = Join-Path $AgentPath "run.cmd"

if (-not (Test-Path $runCmdPath)) {
    Write-Error "run.cmd not found at $runCmdPath - update -AgentPath and try again."
    exit 1
}

# Warn if an agent process is already running - avoid accidentally launching a duplicate
$existing = Get-Process -Name "Agent.Listener" -ErrorAction SilentlyContinue
if ($existing) {
    Write-Warning "Agent.Listener.exe is already running (PID $($existing.Id)). Launching another instance may cause conflicts."
    $confirm = Read-Host "Continue anyway? (y/N)"
    if ($confirm -ne "y") {
        Write-Host "Aborted."
        exit 0
    }
}

Write-Host "Launching agent detached from this console..." -ForegroundColor Cyan

# Launch run.cmd as its own independent process tree (NOT a child of this shell)
$proc = Start-Process -FilePath $runCmdPath -WorkingDirectory $AgentPath -PassThru -WindowStyle Normal

Start-Sleep -Seconds 3

# Give run.cmd a moment to spawn the actual Agent.Listener.exe, then find it
$agentProc = Get-Process -Name "Agent.Listener" -ErrorAction SilentlyContinue |
    Sort-Object StartTime -Descending | Select-Object -First 1

if ($agentProc) {
    $agentInfo = Get-CimInstance Win32_Process -Filter "ProcessId = $($agentProc.Id)"
    $parentInfo = Get-CimInstance Win32_Process -Filter "ProcessId = $($agentInfo.ParentProcessId)" -ErrorAction SilentlyContinue

    $record = [PSCustomObject]@{
        Time                = Get-Date
        RunCmdPID           = $proc.Id
        AgentListenerPID    = $agentProc.Id
        AgentParentPID      = $agentInfo.ParentProcessId
        AgentParentName     = if ($parentInfo) { $parentInfo.Name } else { "N/A (parent already exited - good sign, fully detached)" }
        LaunchingShellPID   = $PID
    }

    $record | Export-Csv -Path $launchLog -Append -NoTypeInformation

    Write-Host ""
    Write-Host "Agent launched. Verification:" -ForegroundColor Green
    Write-Host "  Agent.Listener.exe PID : $($agentProc.Id)"
    Write-Host "  Agent's parent PID     : $($agentInfo.ParentProcessId)  ($($record.AgentParentName))"
    Write-Host "  This shell's PID       : $PID"

    if ($agentInfo.ParentProcessId -eq $PID) {
        Write-Warning "Agent is STILL a child of this shell (PID $PID). Detachment did not work as expected - investigate before closing this window."
    }
    else {
        Write-Host "  Confirmed: agent is NOT a child of this console. Safe to close this window / disconnect." -ForegroundColor Green
    }

    Write-Host ""
    Write-Host "Launch record appended to $launchLog"
}
else {
    Write-Warning "Could not find Agent.Listener.exe after launch - it may still be starting, or run.cmd failed. Check manually with: Get-Process Agent.Listener"
}
