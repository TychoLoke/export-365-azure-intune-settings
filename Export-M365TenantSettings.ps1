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

function Ensure-Module {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ModuleName
    )

    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Install-Module -Name $ModuleName -Scope CurrentUser -Force
    }

    Import-Module -Name $ModuleName -ErrorAction Stop
}

function Ensure-OutputDirectory {
    param([string]$Path)

    if (-not (Test-Path -Path $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Export-SectionData {
    param(
        [Parameter(Mandatory = $true)]
        [string]$OutputDirectory,
        [Parameter(Mandatory = $true)]
        [string]$SectionName,
        [Parameter(Mandatory = $true)]
        $Data
    )

    Ensure-OutputDirectory -Path $OutputDirectory

    $jsonPath = Join-Path -Path $OutputDirectory -ChildPath "$SectionName.json"
    $Data | ConvertTo-Json -Depth 10 | Set-Content -Path $jsonPath -Encoding UTF8
    Write-Host "Exported $SectionName JSON to $jsonPath"

    if ($Data -is [System.Collections.IEnumerable] -and -not ($Data -is [string])) {
        try {
            $csvPath = Join-Path -Path $OutputDirectory -ChildPath "$SectionName.csv"
            $Data | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
            Write-Host "Exported $SectionName CSV to $csvPath"
        } catch {
            Write-Warning "Skipped CSV export for $SectionName because the data structure is not flat enough."
        }
    }
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

    if (-not (Get-MgContext)) {
        Ensure-Module -ModuleName "Microsoft.Graph"
        Connect-MgGraph -Scopes $Scopes -NoWelcome
    }
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
        Export-SectionData -OutputDirectory $OutputDirectory -SectionName $section -Data $data
    } catch {
        Write-Warning "Failed to export $section: $($_.Exception.Message)"
    }
}

if (Get-MgContext) {
    Disconnect-MgGraph | Out-Null
}
