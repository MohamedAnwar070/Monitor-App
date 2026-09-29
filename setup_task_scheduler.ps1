# ============================================================
#  setup_task_scheduler.ps1
#  Run this ONCE as Administrator to register the Task Scheduler jobs.
#  It creates TWO daily triggers: 10:00 AM and 4:30 PM IST.
#
#  HOW TO RUN:
#    Right-click PowerShell ? "Run as Administrator"
#    Then run: .\setup_task_scheduler.ps1
# ============================================================

$TaskName   = "GitHub_Website_Monitor_Trigger"
$ScriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$ScriptFile = Join-Path $ScriptDir "trigger_workflow.ps1"

# Remove old task if it exists
if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Host "??  Removed old task: $TaskName"
}

# Define Action: run the PowerShell script
$Action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptFile`""

# Define TWO daily triggers
# 10:00 AM IST
$Trigger1 = New-ScheduledTaskTrigger -Daily -At "10:00AM"
# 4:30 PM IST
$Trigger2 = New-ScheduledTaskTrigger -Daily -At "04:30PM"

# Settings
$Settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 10) `
    -StartWhenAvailable `
    -RunOnlyIfNetworkAvailable `
    -MultipleInstances IgnoreNew

# Principal: run as current user (INTERACTIVE)
$Principal = New-ScheduledTaskPrincipal `
    -UserId ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) `
    -LogonType Interactive `
    -RunLevel Highest

# Register the task
Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $Action `
    -Trigger $Trigger1, $Trigger2 `
    -Settings $Settings `
    -Principal $Principal `
    -Description "Triggers the GitHub Actions Website Monitor workflow at 10:00 AM and 4:30 PM daily (IST)" `
    -Force

Write-Host ""
Write-Host "============================================================"
Write-Host "  ? Task '$TaskName' registered successfully!"
Write-Host "  ? Runs at: 10:00 AM  and  4:30 PM  every day (IST)"
Write-Host "  ?? Script: $ScriptFile"
Write-Host "============================================================"
Write-Host ""
Write-Host "??  To verify: Open Task Scheduler ? Task Scheduler Library ? $TaskName"
Write-Host "??  To test now: Run the task manually from Task Scheduler or run:"
Write-Host "    Start-ScheduledTask -TaskName '$TaskName'"
