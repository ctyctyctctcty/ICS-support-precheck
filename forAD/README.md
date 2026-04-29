# DHCP Export for AD/DHCP Server

These scripts run on the DHCP server side. They do not use PowerShell remoting.

The goal is to export DHCP scope/range information once a month into a folder that the support precheck tool can read through `DHCP_REFERENCE_PATH`.

These scripts now include export-side validation:

- fail if no IPv4 scopes are found
- fail if required fields are missing
- fail if `StartRange` is greater than `EndRange`
- export DHCP exclusion ranges as `RangeType=exclusion`
- always write `dhcp_export_status.json` so the support tool can verify export freshness and success

## Files

- `.env.example`: sample settings. Copy it to `.env` and edit locally.
- `export_dhcp_ranges.ps1`: exports DHCP scopes to CSV/JSON.
- `register_monthly_task.ps1`: registers a monthly Windows Task Scheduler job.

## Setup

1. Copy `.env.example` to `.env`.
2. Set `EXPORT_DESTINATION` to the shared folder used by the support tool.
3. Run PowerShell as Administrator on the DHCP server.
4. Register the monthly task:

```powershell
cd C:\path\to\ICS-support-precheck\forAD
.\register_monthly_task.ps1
```

5. Run once manually to confirm output:

```powershell
.\export_dhcp_ranges.ps1
```

Expected output:

- one or more `dhcp_scopes_yyyyMMdd_HHmmss.csv` / `.json` files
- `dhcp_export_status.json`

The generated CSV includes columns such as:

```csv
RangeType,ScopeId,Name,StartRange,EndRange,SubnetMask,State,LeaseDuration,ExportedAt,ServerName
```

The generated status file includes fields such as:

```json
{
  "success": true,
  "exported_at": "2026-04-27 06:00:00",
  "server_name": "DHCP01",
  "scope_count": 12,
  "exclusion_count": 3,
  "data_files": ["dhcp_scopes_20260427_060000.csv"]
}
```

On the support tool side, point `config\.env` to this output folder:

```dotenv
DHCP_REFERENCE_PATH=\\fileserver\share\dhcp_exports
```

## Permission Notes

The scheduled task account needs:

- Permission to read DHCP scopes on the DHCP server.
- Write permission to `EXPORT_DESTINATION`.

No API keys or applicant data are required on the DHCP server.
