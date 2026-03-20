# Export 365 Azure Intune Settings

This repository contains a PowerShell utility for exporting configuration snapshots from a Microsoft 365 tenant across Entra ID, Exchange Online, SharePoint Online, Teams, and Intune-related workloads.

## Status

This is now a cleaner public utility built around section-based exports and safer defaults. It still relies on a few workload-specific Microsoft modules, but it no longer installs and connects to everything unconditionally.

## What It Does

- Connects only to the workloads required by the sections you select
- Exports both JSON and CSV where practical
- Uses Microsoft Graph for Entra and Intune-related exports
- Keeps Exchange, Teams, and SharePoint exports separate and optional

## Requirements

- PowerShell 7 recommended
- Permission to install required PowerShell modules for the current user
- Administrator or reader access appropriate for the Microsoft 365 workloads you want to export
- `SharePointAdminUrl` only if you want SharePoint export sections
- Internet access to bootstrap `PowerShellAdminHelpers` from GitHub the first time you run the script

## Usage

The script will automatically install `PowerShellAdminHelpers` from `TychoLoke/powershell-admin-helpers` if the shared module is not already installed.

```powershell
.\Export-M365TenantSettings.ps1 `
  -OutputDirectory "C:\Temp\M365-Exports" `
  -Sections Organization,EntraDirectorySettings,IntuneDeviceCompliance,Teams `
  -ExchangeUserPrincipalName "admin@contoso.com" `
  -SharePointAdminUrl "https://contoso-admin.sharepoint.com"
```

For compatibility, `powershell.ps1` remains as a wrapper that forwards to `Export-M365TenantSettings.ps1`.

## Available Sections

- `Organization`
- `EntraDirectorySettings`
- `ExchangeOrganization`
- `SharePointTenantSites`
- `Teams`
- `IntuneDeviceCompliance`
- `IntuneDeviceConfiguration`
- `IntuneMobileApps`
- `IntuneAppProtection`
- `IntuneAppConfiguration`

## Notes

- Not every export surface returns flat data, so JSON is always produced and CSV is generated when the structure is compatible.
- The script continues on section failures and logs warnings for failed sections instead of aborting the whole run.
- Review the generated output before using it as a compliance, migration, or backup source.

## License

This project is licensed under the MIT License.
