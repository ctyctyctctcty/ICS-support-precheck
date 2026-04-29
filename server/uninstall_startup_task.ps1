$ErrorActionPreference = 'Stop'

$TaskName = 'ICS Support Precheck Web'

$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw 'Please run this script as Administrator.'
}

Write-Host "Stopping startup task if it is running: $TaskName"
schtasks /End /TN $TaskName 2>$null | Out-Host

Write-Host "Deleting startup task: $TaskName"
schtasks /Delete /TN $TaskName /F | Out-Host

Write-Host ''
Write-Host 'Done.'
Read-Host 'Press Enter to close'
