$TaskName = 'ICS Support Precheck Web'
$AppDir = Resolve-Path (Join-Path $PSScriptRoot '..')

Write-Host "Task status: $TaskName"
schtasks /Query /TN $TaskName /FO LIST /V

Write-Host ''
Write-Host 'Health check:'
try {
    $response = Invoke-WebRequest -Uri 'http://127.0.0.1:8010/api/health' -UseBasicParsing -TimeoutSec 3
    Write-Host "OK: HTTP $($response.StatusCode)"
}
catch {
    Write-Host "NG: $($_.Exception.Message)"
}

Write-Host ''
Write-Host "Logs: $AppDir\logs"
