$ErrorActionPreference = 'Stop'

$AppDir = Resolve-Path (Join-Path $PSScriptRoot '..')
$LogDir = Join-Path $AppDir 'logs'
$ServiceLog = Join-Path $LogDir 'web-service.log'
$RunLog = Join-Path $LogDir 'uvicorn.log'

New-Item -ItemType Directory -Force -Path $LogDir | Out-Null
Set-Location $AppDir

function Write-ServiceLog {
    param([string]$Message)
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Add-Content -Path $ServiceLog -Value "[$timestamp] $Message" -Encoding UTF8
}

function Get-PythonCommand {
    if ($env:ICS_PRECHECK_PYTHON) {
        return $env:ICS_PRECHECK_PYTHON
    }

    $py = Get-Command py -ErrorAction SilentlyContinue
    if ($py) {
        return $py.Source
    }

    $python = Get-Command python -ErrorAction SilentlyContinue
    if ($python) {
        return $python.Source
    }

    throw 'Python was not found. Install Python for all users or set ICS_PRECHECK_PYTHON.'
}

$pythonCommand = Get-PythonCommand
Write-ServiceLog "Service runner started. AppDir=$AppDir Python=$pythonCommand"

while ($true) {
    try {
        Write-ServiceLog 'Starting web service.'
        & $pythonCommand start_web.py *> $RunLog
        $exitCode = $LASTEXITCODE
        Write-ServiceLog "Web service stopped. ExitCode=$exitCode"
    }
    catch {
        Write-ServiceLog "Web service failed to start or crashed. $($_.Exception.Message)"
    }

    Start-Sleep -Seconds 10
}
