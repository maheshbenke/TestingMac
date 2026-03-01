<#
.SYNOPSIS
    Registers (or removes) a Windows Scheduled Task that runs Scan-Outlook.ps1
    every hour.

.PARAMETER Action
    Install   – creates the scheduled task (default)
    Uninstall – removes the scheduled task

.PARAMETER ScriptPath
    Full path to Scan-Outlook.ps1. Defaults to the copy next to this script.

.PARAMETER HoursBack
    Passed through to Scan-Outlook.ps1 (-HoursBack). Default: 2
    Using 2 gives overlap so no emails are missed between hourly runs.

.PARAMETER IncludeBody
    Pass -IncludeBody to include longer body text in the export.

.NOTES
    Must be run elevated (Run as Administrator) to create scheduled tasks.
#>

[CmdletBinding()]
param(
    [ValidateSet('Install', 'Uninstall')]
    [string]$Action = 'Install',

    [string]$ScriptPath = (Join-Path $PSScriptRoot 'Scan-Outlook.ps1'),

    [int]$HoursBack = 2,

    [switch]$IncludeBody
)

$TaskName = 'OutlookEmailScanner'
$TaskPath = '\CopilotTools\'

if ($Action -eq 'Uninstall') {
    Write-Host "Removing scheduled task '$TaskName' ..."
    Unregister-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host 'Done.'
    return
}

# ── Install ──────────────────────────────────────────────────────────────────

if (-not (Test-Path $ScriptPath)) {
    throw "Scanner script not found at: $ScriptPath"
}

$bodyFlag = if ($IncludeBody) { ' -IncludeBody' } else { '' }
$arguments = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$ScriptPath`" -HoursBack $HoursBack$bodyFlag"

$action_obj = New-ScheduledTaskAction `
    -Execute 'powershell.exe' `
    -Argument $arguments `
    -WorkingDirectory (Split-Path $ScriptPath)

# Trigger: every 1 hour, starting now
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Hours 1)

# Settings: run whether logged in or not (for current user), allow on battery
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 10)

$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited

# Register
$existingTask = Get-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -ErrorAction SilentlyContinue
if ($existingTask) {
    Write-Host "Task already exists – updating..."
    Set-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath `
        -Action $action_obj -Trigger $trigger -Settings $settings -Principal $principal | Out-Null
}
else {
    Register-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath `
        -Action $action_obj -Trigger $trigger -Settings $settings -Principal $principal | Out-Null
}

Write-Host @"

Scheduled task '$TaskName' registered successfully.

  Schedule : Every 1 hour
  Script   : $ScriptPath
  HoursBack: $HoursBack
  Output   : $(Split-Path $ScriptPath)\outlook_emails.json

To run immediately:
  Start-ScheduledTask -TaskName '$TaskName' -TaskPath '$TaskPath'

To uninstall:
  .\Install-ScheduledTask.ps1 -Action Uninstall
"@
  
