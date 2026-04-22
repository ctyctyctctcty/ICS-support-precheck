[CmdletBinding()]
param(
    [string]$EnvPath = (Join-Path $PSScriptRoot ".env")
)

$ErrorActionPreference = "Stop"

function Read-DotEnv {
    param([string]$Path)

    $values = @{}
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Env file not found: $Path. Copy .env.example to .env first."
    }

    foreach ($rawLine in Get-Content -LiteralPath $Path -Encoding UTF8) {
        $line = $rawLine.Trim()
        if (-not $line -or $line.StartsWith("#") -or -not $line.Contains("=")) {
            continue
        }
        $key, $value = $line.Split("=", 2)
        $values[$key.Trim()] = $value.Trim().Trim('"').Trim("'")
    }
    return $values
}

function Value-OrDefault {
    param(
        [hashtable]$Values,
        [string]$Key,
        [string]$Default
    )

    $value = [string]($Values[$Key])
    if ([string]::IsNullOrWhiteSpace($value)) {
        return $Default
    }
    return $value
}

$envValues = Read-DotEnv -Path $EnvPath
$taskName = Value-OrDefault -Values $envValues -Key "TASK_NAME" -Default "ICS Support Precheck DHCP Export"
$description = Value-OrDefault -Values $envValues -Key "TASK_DESCRIPTION" -Default "Export DHCP scope ranges for ICS support precheck."
$dayOfMonth = [int](Value-OrDefault -Values $envValues -Key "TASK_DAY_OF_MONTH" -Default "1")
$taskTime = Value-OrDefault -Values $envValues -Key "TASK_TIME" -Default "06:00"
$runAsUser = Value-OrDefault -Values $envValues -Key "TASK_RUN_AS_USER" -Default ""

if ($dayOfMonth -lt 1 -or $dayOfMonth -gt 31) {
    throw "TASK_DAY_OF_MONTH must be between 1 and 31."
}

$scriptPath = Join-Path $PSScriptRoot "export_dhcp_ranges.ps1"
if (-not (Test-Path -LiteralPath $scriptPath)) {
    throw "Export script not found: $scriptPath"
}

$taskCommand = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
$schtasksArgs = @(
    "/Create",
    "/TN", $taskName,
    "/TR", $taskCommand,
    "/SC", "MONTHLY",
    "/D", ([string]$dayOfMonth),
    "/ST", $taskTime,
    "/RL", "HIGHEST",
    "/F"
)

if ($runAsUser) {
    $credential = Get-Credential -UserName $runAsUser -Message "Password for the scheduled task account"
    $schtasksArgs += @("/RU", $credential.UserName, "/RP", $credential.GetNetworkCredential().Password)
}

& schtasks.exe @schtasksArgs | Out-Host
if ($LASTEXITCODE -ne 0) {
    throw "schtasks.exe failed with exit code $LASTEXITCODE"
}

Write-Host "Registered scheduled task: $taskName"
Write-Host "Description: $description"
Write-Host "Schedule: monthly on day $dayOfMonth at $taskTime"
Write-Host "Script: $scriptPath"
