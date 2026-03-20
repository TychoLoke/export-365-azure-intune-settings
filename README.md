# Export 365 Azure Intune Settings

This repository contains a PowerShell export script for collecting configuration snapshots from a Microsoft 365 tenant across Azure AD, Exchange Online, SharePoint Online, Teams, and Intune-related workloads.

## Status

This is a legacy utility that has been cleaned up for public use, but it still depends on a mix of older Microsoft administration modules.
Treat it as a best-effort export tool and validate the output in a non-production tenant before relying on it operationally.

## What It Does

- Installs and imports the modules required by the selected export sections
- Connects to multiple Microsoft 365 services
- Runs a set of export commands for each workload
- Writes section-based CSV output to a chosen directory

## Requirements

- PowerShell 7 or Windows PowerShell 5.1
- Permission to install required PowerShell modules for the current user
- Administrator or reader access appropriate for the Microsoft 365 workloads you want to export

## Usage

```powershell
.\powershell.ps1 `
  -OutputDirectory "C:\Temp\M365-Exports" `
  -ExchangeUserPrincipalName "admin@contoso.com" `
  -SharePointAdminUrl "https://contoso-admin.sharepoint.com"
```

The script will prompt for interactive sign-in where required.

## Notes

- Some of the modules used here are legacy modules kept for compatibility with the original script.
- Not every command is available in every tenant or module combination. The script continues on section failures and logs warnings for failed sections.
- Review the generated CSV files before using them as a compliance or migration source.

## License

This project is licensed under the MIT License.
