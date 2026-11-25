# Install the Cash Flow Carnival update on NobleOfficeMgr.
# Run this in Windows PowerShell as Kelly. It does not ask for the passcode
# and it does not print the passcode. Invoice checks can use this same
# Task Scheduler folder and the same Windows user later.
#
# Setup, once:
# 1. Clone the repo, if it is not already there:
#      git clone https://github.com/kellynoble26/noble-rep-payments.git C:\Users\kelly\NobleOfficeMgr\noble-rep-payments
# 2. Set the passcode for this Windows user only, then sign out and back in
#    so the scheduled task can see it. Do not put it in a file or in git:
#      [Environment]::SetEnvironmentVariable("FIELD_LEDGER_PASSCODE", "<the passcode>", "User")
# 3. Gmail, for mgr@noblewandd.com:
#      - In Google Cloud, create an OAuth desktop client for that account.
#      - Save the downloaded JSON as
#        C:\Users\kelly\NobleOfficeMgr\cash-flow-carnival\gmail-client.json
#      - From the repo folder, run:
#        py -3 tools\cash_flow_update.py --gmail-consent
#      - Sign in as mgr@noblewandd.com and allow read-only Gmail.
# 4. GitHub push:
#      - Install Git Credential Manager if it is not already installed.
#      - In the repo folder, run: git push origin main
#      - Sign in with an account that can push main. The credential stays
#        in Windows Credential Manager.
#      - If GitHub rejects the push because main requires a pull request,
#        a bypass for this PC's account is required before the live page
#        can update by itself.
# 5. Outlook must already be signed in as mgr@noblegcllc.com. The task runs
#    only while Kelly is logged on, so Outlook can be read. The screen can
#    be locked. If the PC is asleep, the missed run starts after it wakes
#    and looks back over the gap.
# 6. Install the task:
#      powershell -ExecutionPolicy Bypass -File tools\install_cash_flow_task.ps1
# 7. Check without publishing:
#      py -3 tools\cash_flow_update.py --check
#
# Failures are written to:
#   C:\Users\kelly\NobleOfficeMgr\logs\cash-flow-carnival.log
# and the task shows as failed in Task Scheduler under \NobleOfficeMgr.

[CmdletBinding()]
param(
    [string]$RepoPath = "C:\Users\kelly\NobleOfficeMgr\noble-rep-payments",
    [string]$TaskFolder = "\NobleOfficeMgr\",
    [string]$TaskName = "Cash Flow Carnival"
)

$ErrorActionPreference = "Stop"
New-Item -ItemType Directory -Force -Path "C:\Users\kelly\NobleOfficeMgr\logs" | Out-Null
New-Item -ItemType Directory -Force -Path "C:\Users\kelly\NobleOfficeMgr\cash-flow-carnival" | Out-Null

$scriptPath = Join-Path $RepoPath "tools\cash_flow_update.py"
if (-not (Test-Path $scriptPath)) {
    Write-Error "Clone the repo to $RepoPath first. The setup steps are at the top of this script."
}

$python = Get-Command py -ErrorAction SilentlyContinue
if (-not $python) {
    $python = Get-Command python -ErrorAction SilentlyContinue
}
if (-not $python) {
    Write-Error "Install Python 3 on this PC, then run this installer again."
}
$pythonExe = $python.Source
$pipArgs = @("-m", "pip", "install", "--user", "openpyxl", "cryptography", "pywin32", "google-auth", "google-auth-oauthlib", "google-api-python-client")
$runArgs = "`"$scriptPath`""
if ($python.Name -like "py*") {
    $pipArgs = @("-3") + $pipArgs
    $runArgs = "-3 $runArgs"
}
& $pythonExe @pipArgs
if ($LASTEXITCODE -ne 0) {
    Write-Error "Could not install the Python packages this task needs."
}

$passcode = [Environment]::GetEnvironmentVariable("FIELD_LEDGER_PASSCODE", "User")
if ([string]::IsNullOrEmpty($passcode)) {
    Write-Warning "FIELD_LEDGER_PASSCODE is not set for this Windows user. The task will fail until it is set and you sign in again."
} else {
    Write-Host "FIELD_LEDGER_PASSCODE is set. The value was not printed."
}
Remove-Variable passcode

$user = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$principal = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive -RunLevel Limited
$action = New-ScheduledTaskAction -Execute $pythonExe -Argument $runArgs -WorkingDirectory $RepoPath
$start = Get-Date "06:00"
$trigger = New-ScheduledTaskTrigger -Daily -At $start
$repeat = New-ScheduledTaskTrigger -Once -At $start -RepetitionInterval (New-TimeSpan -Hours 2) -RepetitionDuration (New-TimeSpan -Days 1)
$trigger.Repetition = $repeat.Repetition
$settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 30)

Register-ScheduledTask `
    -TaskName $TaskName `
    -TaskPath $TaskFolder `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -Principal $principal `
    -Force | Out-Null

Write-Host "Installed $TaskFolder$TaskName for $user."
Write-Host "It runs every 2 hours, and a missed run starts after the PC wakes."
Write-Host "A run with no new payments does not commit or push."
Write-Host "Log: C:\Users\kelly\NobleOfficeMgr\logs\cash-flow-carnival.log"
