# Revised InstallAndImport-Module function
function InstallAndImport-Module {
    param (
        [Parameter(Mandatory=$true)] [string] $Name,
        [switch] $AllowClobber,
        [switch] $VerboseMode
    )
    if (-not (Get-InstalledModule -Name $Name -ErrorAction SilentlyContinue)) {
        Install-Module -Name $Name -AllowClobber:$AllowClobber -Verbose:$VerboseMode -Scope CurrentUser
    }
    Import-Module -Name $Name -Verbose:$VerboseMode
}

$requiredModules = @(
    "AzureAD",
    "MSOnline",
    "ExchangeOnlineManagement",
    "SharePointPnPPowerShellOnline",
    "Microsoft.Graph.Intune",
    "MicrosoftTeams",
    "Microsoft.Graph"
)

# Install and import required modules
foreach ($module in $requiredModules) {
    Write-Host "Installing and importing module: $module"
    InstallAndImport-Module -Name $module -AllowClobber -VerboseMode
}

# Check if the required cmdlets are available
$requiredCmdlets = @(
    "Get-AzureADMSDirectorySetting",
    "Get-MsolCompanyInformation",
    "Get-OrganizationConfig",
    "Get-PnPWeb",
    "Get-Team"
)

foreach ($cmdlet in $requiredCmdlets) {
    if (-not (Get-Command $cmdlet -ErrorAction SilentlyContinue)) {
        Write-Error "Cmdlet $cmdlet not found. Please check the module installation and imports."
        return
    }
}

$sections = @(
    @{
        Name = "AzureADSettings";
        ScriptBlock = { Get-AzureADMSDirectorySetting }
    },
    @{
        Name = "MSOLSettings";
        ScriptBlock = { Get-MsolCompanyInformation }
    },
    @{
        Name = "ExchangeSettings";
        ScriptBlock = { Get-OrganizationConfig }
    },
    @{
        Name = "SharePointSettings";
        ScriptBlock = { Get-PnPWeb }
    },
    @{
        Name = "TeamsSettings";
        ScriptBlock = { Get-Team }
    },
    @{
        Name = "DeviceComplianceSettings";
        ScriptBlock = { Get-IntuneDeviceCompliancePolicy }
    },
    @{
        Name = "DeviceConfigurationSettings";
        ScriptBlock = { Get-IntuneDeviceConfigurationPolicy }
    },
    @{
        Name = "MobileAppSettings";
        ScriptBlock = { Get-IntuneMobileApp }
    },
    @{
        Name = "AppProtectionSettings";
        ScriptBlock = { Get-IntuneAppProtectionPolicy }
    },
    @{
        Name = "AppConfigurationSettings";
        ScriptBlock = { Get-MgIntuneAppConfigurationPolicy }
    }
)

# Connect to your Microsoft 365/Azure environment
Connect-AzureAD
Connect-MsolService
Connect-ExchangeOnline -UserPrincipalName admin@bradonbureauvandetoekomst.nl
Connect-PnPOnline -Url https://bureauvandetoekomst.sharepoint.com -Credentials (Get-Credential)
Connect-MSGraph
Connect-MgGraph
Connect-MicrosoftTeams

foreach ($section in $sections) {
    try {
        $settingsObject = Invoke-Command -ScriptBlock $section.ScriptBlock
        Export-SettingsToCSV -CsvFilePath $csvFilePath -SettingsObject $settingsObject -SectionName $section.Name
    } catch {
        Write-Warning "Failed to retrieve $($section.Name) due to an error: $($_.Exception.Message)"
    }
}
``
