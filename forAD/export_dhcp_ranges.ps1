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

function Get-RequiredValue {
    param(
        [hashtable]$Values,
        [string]$Key
    )

    $value = [string]($Values[$Key])
    if ([string]::IsNullOrWhiteSpace($value)) {
        throw "$Key is not configured in .env."
    }
    return $value
}

$envValues = Read-DotEnv -Path $EnvPath
$destination = Get-RequiredValue -Values $envValues -Key "EXPORT_DESTINATION"
$format = [string]$envValues["EXPORT_FORMAT"]
if ([string]::IsNullOrWhiteSpace($format)) {
    $format = "both"
}
$format = $format.Trim().ToLowerInvariant()

$prefix = [string]$envValues["EXPORT_PREFIX"]
if ([string]::IsNullOrWhiteSpace($prefix)) {
    $prefix = "dhcp_scopes"
}
$prefix = $prefix.Trim()

if ($format -notin @("csv", "json", "both")) {
    throw "EXPORT_FORMAT must be csv, json, or both."
}

if (-not (Test-Path -LiteralPath $destination)) {
    New-Item -ItemType Directory -Path $destination -Force | Out-Null
}

Import-Module DhcpServer -ErrorAction Stop

$exportedAt = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$serverName = $env:COMPUTERNAME

$scopes = Get-DhcpServerv4Scope | Sort-Object ScopeId | ForEach-Object {
    [pscustomobject]@{
        ScopeId = $_.ScopeId.ToString()
        Name = $_.Name
        StartRange = $_.StartRange.ToString()
        EndRange = $_.EndRange.ToString()
        SubnetMask = $_.SubnetMask.ToString()
        State = $_.State.ToString()
        LeaseDuration = $_.LeaseDuration.ToString()
        ExportedAt = $exportedAt
        ServerName = $serverName
    }
}

if (-not $scopes -or $scopes.Count -eq 0) {
    throw "No DHCP IPv4 scopes were found on this server."
}

$written = @()
if ($format -in @("csv", "both")) {
    $csvPath = Join-Path $destination "$prefix`_$timestamp.csv"
    $scopes | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
    $written += $csvPath
}

if ($format -in @("json", "both")) {
    $jsonPath = Join-Path $destination "$prefix`_$timestamp.json"
    [pscustomobject]@{
        exported_at = $exportedAt
        server_name = $serverName
        scopes = $scopes
    } | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $jsonPath -Encoding UTF8
    $written += $jsonPath
}

foreach ($path in $written) {
    Write-Host "Wrote $path"
}
