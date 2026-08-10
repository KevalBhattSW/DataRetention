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
    [string]$AgentPath = "c:\users\UDRTagging\Agent",   # <-- UPDATE to your actual agent install path

    [Parameter(Mandatory = $false)]
    [string]$LogDir = "C:\Temp\logs"
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
# SIG # Begin signature block
# MIIFjgYJKoZIhvcNAQcCoIIFfzCCBXsCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUWNjU9m1vBiiKAWZ25LN3pY1q
# QySgggMeMIIDGjCCAgKgAwIBAgIQE7EWVmkfAJtDYwn54Vq9kjANBgkqhkiG9w0B
# AQsFADAlMSMwIQYDVQQDDBpVRFIgVGFnZ2luZyBTY3JpcHQgU2lnbmluZzAeFw0y
# NjA3MDMxMDAyNTVaFw0zMTA3MDMxMDEyNTFaMCUxIzAhBgNVBAMMGlVEUiBUYWdn
# aW5nIFNjcmlwdCBTaWduaW5nMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKC
# AQEA0ahFLuBWT8eQYEXp4UNGKqZ+aBrj+Wjn95UOc5E8cm78DK2RGMO88iQHRYsv
# lc8UwNOZA3VL9CZOdUilRTZfxlHtwpPQP9jdho+ceTOdSRQpznTTe2MKZT1WzQ4R
# v7dhYKDUAbb7jiSteYtb+L2tR0UPDuFHNekRpG35eq2J+/x96h2ZA+DhM54Mz83g
# Ie59Cs2T95sYsA+eSbcfRPJdfz2KPmz+DrQHYsgkTveRnOvr36b6vuyiVyBDXQzp
# CxnDGxTIxdLkJtIyefP0fyCgZgZdKH97GVmdfYjfoYxJ5sN/Hk4flUEjMuMlfTmM
# mJ40TmwHtOKtkwfeKXsE9SyOsQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYD
# VR0lBAwwCgYIKwYBBQUHAwMwHQYDVR0OBBYEFJ6YYuR91Gg43DzDNPmgIHbPYkoH
# MA0GCSqGSIb3DQEBCwUAA4IBAQA91UnsH6jMWAdf6URgNeuioWjnW1VcVnf8Rwta
# BbH6SOi2Ep/ILWGHJr/Y6vTgX5kNasmKlbdF4d9uCKdQMn7VVmIaJyHQXaH4Hxhn
# tds2kuJ9Jjmc27lx4jCVshlACn53hOhJNJLED0X+kxgedzY6kkS8bZXkMonfDesG
# CYYnMtVYuf1PinYa3zeUxuZBt6HhD5ny9KDv4R96KrPRzfAkDhHv/o6X0/pCQlF9
# ms5deEvBGRa0Lx1EkSzP+CyHzC8Ovi/LdjvP6chjA4eYr3DGRt5Nd/pLwShO5dJ/
# qqW+96CM0MKNB8+7wtVMpMfqDQz1GjzppehQz5qObQOfqjwFMYIB2jCCAdYCAQEw
# OTAlMSMwIQYDVQQDDBpVRFIgVGFnZ2luZyBTY3JpcHQgU2lnbmluZwIQE7EWVmkf
# AJtDYwn54Vq9kjAJBgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIBDDEKMAigAoAAoQKA
# ADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4wDAYK
# KwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUf0Zfj1j5dxtFtiGSRFP4nTQI6QUw
# DQYJKoZIhvcNAQEBBQAEggEArq+bjb2HwaTD7xMlV4bIKZJdgxfgutnTqLDKlcW7
# LObkV0EITv0yYHlWRiyZPNVGYzbwjOTUUGGmHKO9EON2McRS7NLEYdABMBMPGxdk
# 5pWtnfLubKmGth2txbqzntgjX4n/ZAzsS/1z0RPas7iklji6363Jls1V9ym7eKgz
# plfzgVKyP7fIpXqXnwtA/P4pVrA+ka/JXOYIIQ3VCkQbLYQImLa1+TqaHuUh+EtC
# uDfWaelQkkD+RQQla/AI/hC5/Grk6+uqMd5QTX5W6xYKiqePj0MUqSdAwjecKIS1
# Vap07p9VhTdFZVtneLIbceMb97JbCPead5Hp9emSZF6DMQ==
# SIG # End signature block
