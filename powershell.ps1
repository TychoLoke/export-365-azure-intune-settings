# Revised InstallAndImport-Module function test
function InstallAndImport-Module {
    param (
        [Parameter(Mandatory=$true)] [string] $Name,
        [switch] $AllowClobber
       
    )
    if (-not (Get-InstalledModule -Name $Name -ErrorAction SilentlyContinue)) {
        Install-Module -Name $Name -AllowClobber:$AllowClobber -Scope CurrentUser
    }
    Import-Module -Name $Name
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
    InstallAndImport-Module -Name $module -AllowClobber 
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
        ScriptBlock = { Get-AzureADDirectorySetting }
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
Connect-ExchangeOnline -UserPrincipalName username@contoso.nl
Connect-PnPOnline -Url https://contoso.sharepoint.com -Credentials (Get-Credential)
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

