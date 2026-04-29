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

function Write-StatusFile {
    param(
        [string]$Destination,
        [hashtable]$Status
    )

    $statusPath = Join-Path $Destination "dhcp_export_status.json"
    $Status | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $statusPath -Encoding UTF8
}

function Assert-ScopeRows {
    param([object[]]$Rows)

    if (-not $Rows -or $Rows.Count -le 0) {
        throw "No DHCP IPv4 scopes were found on this server."
    }

    foreach ($row in $Rows) {
        foreach ($field in @("ScopeId", "Name", "StartRange", "EndRange")) {
            $value = [string]$row.$field
            if ([string]::IsNullOrWhiteSpace($value)) {
                throw "DHCP export validation failed: missing field $field."
            }
        }

        $start = [System.Net.IPAddress]::Parse([string]$row.StartRange).GetAddressBytes()
        $end = [System.Net.IPAddress]::Parse([string]$row.EndRange).GetAddressBytes()
        $startNum = [BitConverter]::ToUInt32(($start[3], $start[2], $start[1], $start[0]), 0)
        $endNum = [BitConverter]::ToUInt32(($end[3], $end[2], $end[1], $end[0]), 0)
        if ($startNum -gt $endNum) {
            throw "DHCP export validation failed: StartRange is greater than EndRange for scope $($row.ScopeId)."
        }
    }
}

function Assert-RangeRows {
    param(
        [object[]]$Rows,
        [string]$Label
    )

    foreach ($row in $Rows) {
        foreach ($field in @("ScopeId", "StartRange", "EndRange")) {
            $value = [string]$row.$field
            if ([string]::IsNullOrWhiteSpace($value)) {
                throw "$Label validation failed: missing field $field."
            }
        }

        $start = [System.Net.IPAddress]::Parse([string]$row.StartRange).GetAddressBytes()
        $end = [System.Net.IPAddress]::Parse([string]$row.EndRange).GetAddressBytes()
        $startNum = [BitConverter]::ToUInt32(($start[3], $start[2], $start[1], $start[0]), 0)
        $endNum = [BitConverter]::ToUInt32(($end[3], $end[2], $end[1], $end[0]), 0)
        if ($startNum -gt $endNum) {
            throw "$Label validation failed: StartRange is greater than EndRange for scope $($row.ScopeId)."
        }
    }
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

$exportedAt = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$serverName = $env:COMPUTERNAME
$written = @()

try {
    Import-Module DhcpServer -ErrorAction Stop

    $scopes = @(Get-DhcpServerv4Scope | Sort-Object ScopeId | ForEach-Object {
        [pscustomobject]@{
            RangeType = "scope"
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
    })

    Assert-ScopeRows -Rows $scopes

    $exclusions = @()
    foreach ($scope in $scopes) {
        $scopeId = $scope.ScopeId
        $scopeExclusions = @(Get-DhcpServerv4ExclusionRange -ScopeId $scopeId -ErrorAction SilentlyContinue)
        foreach ($excluded in $scopeExclusions) {
            $exclusions += [pscustomobject]@{
                RangeType = "exclusion"
                ScopeId = $scopeId
                Name = $scope.Name
                StartRange = $excluded.StartRange.ToString()
                EndRange = $excluded.EndRange.ToString()
                SubnetMask = $scope.SubnetMask
                State = ""
                LeaseDuration = ""
                ExportedAt = $exportedAt
                ServerName = $serverName
            }
        }
    }
    Assert-RangeRows -Rows $exclusions -Label "DHCP exclusion export"

    $allRows = @($scopes) + @($exclusions)

    if ($format -in @("csv", "both")) {
        $csvPath = Join-Path $destination "$prefix`_$timestamp.csv"
        $allRows | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
        $written += $csvPath
    }

    if ($format -in @("json", "both")) {
        $jsonPath = Join-Path $destination "$prefix`_$timestamp.json"
        [pscustomobject]@{
            exported_at = $exportedAt
            server_name = $serverName
            scopes = $scopes
            exclusions = $exclusions
        } | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $jsonPath -Encoding UTF8
        $written += $jsonPath
    }

    Write-StatusFile -Destination $destination -Status @{
        success = $true
        exported_at = $exportedAt
        server_name = $serverName
        scope_count = $scopes.Count
        exclusion_count = $exclusions.Count
        data_files = @($written | ForEach-Object { [System.IO.Path]::GetFileName($_) })
    }

    foreach ($path in $written) {
        Write-Host "Wrote $path"
    }
}
catch {
    Write-StatusFile -Destination $destination -Status @{
        success = $false
        exported_at = $exportedAt
        server_name = $serverName
        scope_count = 0
        exclusion_count = 0
        data_files = @()
        error = $_.Exception.Message
    }
    throw
}
