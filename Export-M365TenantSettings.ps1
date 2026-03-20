[CmdletBinding()]
param(
    [string]$OutputDirectory = "C:\Temp\M365-Exports",
    [string[]]$Sections = @(
        "Organization",
        "EntraDirectorySettings",
        "ExchangeOrganization",
        "SharePointTenantSites",
        "Teams",
        "IntuneDeviceCompliance",
        "IntuneDeviceConfiguration",
        "IntuneMobileApps",
        "IntuneAppProtection",
        "IntuneAppConfiguration"
    ),
    [string]$ExchangeUserPrincipalName,
    [string]$SharePointAdminUrl
)

$ErrorActionPreference = "Stop"

function Initialize-PowerShellAdminHelpers {
    $moduleName = "PowerShellAdminHelpers"

    if (-not (Get-Module -ListAvailable -Name $moduleName)) {
        $installerPath = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath "Install-PowerShellAdminHelpers.ps1"
        Invoke-WebRequest -Uri "https://raw.githubusercontent.com/TychoLoke/powershell-admin-helpers/main/Install-PowerShellAdminHelpers.ps1" -OutFile $installerPath
        & $installerPath
    }

    Import-Module -Name $moduleName -Force -ErrorAction Stop
}

function Get-GraphCollection {
    param([Parameter(Mandatory = $true)][string]$Uri)

    $items = @()
    $nextUri = $Uri

    while ($nextUri) {
        $response = Invoke-MgGraphRequest -Method GET -Uri $nextUri
        if ($response.value) {
            $items += $response.value
        } else {
            $items += $response
            break
        }
        $nextUri = $response.'@odata.nextLink'
    }

    return $items
}

function Connect-GraphIfNeeded {
    param([string[]]$Scopes)

    if (-not $Scopes) {
        return
    }

    Connect-GraphWithScopes -Scopes $Scopes
}

function Connect-ExchangeIfNeeded {
    param([string]$UserPrincipalName)

    if (-not (Get-Command Connect-ExchangeOnline -ErrorAction SilentlyContinue)) {
        Ensure-Module -ModuleName "ExchangeOnlineManagement"
    }

    if (-not (Get-ConnectionInformation -ErrorAction SilentlyContinue)) {
        if ($UserPrincipalName) {
            Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName
        } else {
            Connect-ExchangeOnline
        }
    }
}

function Connect-SharePointIfNeeded {
    param([string]$AdminUrl)

    if (-not $AdminUrl) {
        throw "SharePointAdminUrl is required for SharePoint export sections."
    }

    Ensure-Module -ModuleName "PnP.PowerShell"

    try {
        Get-PnPConnection | Out-Null
    } catch {
        Connect-PnPOnline -Url $AdminUrl -Interactive
    }
}

function Connect-TeamsIfNeeded {
    Ensure-Module -ModuleName "MicrosoftTeams"

    if (-not (Get-CsTenant -ErrorAction SilentlyContinue)) {
        Connect-MicrosoftTeams
    }
}

Initialize-PowerShellAdminHelpers
Ensure-OutputDirectory -Path $OutputDirectory

$graphSectionScopes = @{
    Organization = @("Organization.Read.All")
    EntraDirectorySettings = @("Directory.Read.All")
    IntuneDeviceCompliance = @("DeviceManagementConfiguration.Read.All")
    IntuneDeviceConfiguration = @("DeviceManagementConfiguration.Read.All")
    IntuneMobileApps = @("DeviceManagementApps.Read.All")
    IntuneAppProtection = @("DeviceManagementApps.Read.All")
    IntuneAppConfiguration = @("DeviceManagementApps.Read.All")
}

$sectionHandlers = @{
    Organization = {
        Get-GraphCollection -Uri "https://graph.microsoft.com/v1.0/organization"
    }
    EntraDirectorySettings = {
        Get-GraphCollection -Uri "https://graph.microsoft.com/beta/settings"
    }
    ExchangeOrganization = {
        Connect-ExchangeIfNeeded -UserPrincipalName $ExchangeUserPrincipalName
        @(Get-OrganizationConfig)
    }
    SharePointTenantSites = {
        Connect-SharePointIfNeeded -AdminUrl $SharePointAdminUrl
        Get-PnPTenantSite
    }
    Teams = {
        Connect-TeamsIfNeeded
        Get-Team
    }
    IntuneDeviceCompliance = {
        Get-GraphCollection -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies"
    }
    IntuneDeviceConfiguration = {
        Get-GraphCollection -Uri "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations"
    }
    IntuneMobileApps = {
        Get-GraphCollection -Uri "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps"
    }
    IntuneAppProtection = {
        Get-GraphCollection -Uri "https://graph.microsoft.com/beta/deviceAppManagement/managedAppPolicies"
    }
    IntuneAppConfiguration = {
        Get-GraphCollection -Uri "https://graph.microsoft.com/beta/deviceAppManagement/mobileAppConfigurations"
    }
}

$requiredGraphScopes = @(
    foreach ($section in $Sections) {
        if ($graphSectionScopes.ContainsKey($section)) {
            $graphSectionScopes[$section]
        }
    }
) | Select-Object -Unique

Connect-GraphIfNeeded -Scopes $requiredGraphScopes

foreach ($section in $Sections) {
    if (-not $sectionHandlers.ContainsKey($section)) {
        Write-Warning "Skipping unknown section: $section"
        continue
    }

    try {
        Write-Host "Collecting $section..."
        $data = & $sectionHandlers[$section]
        Export-ObjectBundle -OutputDirectory $OutputDirectory -SectionName $section -Data $data
        Write-Host "Exported $section to $OutputDirectory"
    } catch {
        Write-Warning "Failed to export $section: $($_.Exception.Message)"
    }
}

if (Get-MgContext) {
    Disconnect-MgGraph | Out-Null
}
