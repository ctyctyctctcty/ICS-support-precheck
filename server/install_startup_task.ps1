$ErrorActionPreference = 'Stop'

$TaskName = 'ICS Support Precheck Web'
$RunScript = Join-Path $PSScriptRoot 'run_web_service.ps1'
$AppDir = Resolve-Path (Join-Path $PSScriptRoot '..')
$TaskRun = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$RunScript`""

$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw 'Please run this script as Administrator.'
}

Write-Host "Registering startup task: $TaskName"
Write-Host "App directory: $AppDir"

schtasks /Create /TN $TaskName /SC ONSTART /DELAY 0001:00 /RL HIGHEST /RU SYSTEM /TR $TaskRun /F | Out-Host
schtasks /Run /TN $TaskName | Out-Host

Write-Host ''
Write-Host 'Done.'
Write-Host 'The service will keep running after logout and will start automatically after reboot.'
Write-Host 'Local URL:  http://127.0.0.1:8010'
Write-Host 'Shared URL: http://SERVER_IP:8010'
Write-Host "Logs:       $AppDir\logs"
Read-Host 'Press Enter to close'
