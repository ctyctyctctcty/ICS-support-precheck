# DHCP Export for AD/DHCP Server

These scripts run on the DHCP server side. They do not use PowerShell remoting.

The goal is to export DHCP scope/range information once a month into a folder that the support precheck tool can read through `DHCP_REFERENCE_PATH`.

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

The generated CSV includes columns such as:

```csv
ScopeId,Name,StartRange,EndRange,SubnetMask,State,LeaseDuration,ExportedAt,ServerName
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
