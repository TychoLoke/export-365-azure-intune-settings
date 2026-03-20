param(
    [string]$OutputDirectory = "C:\Temp\M365-Exports",
    [string]$ExchangeUserPrincipalName,
    [string]$SharePointAdminUrl
)

function Install-AndImportModule {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [switch]$AllowClobber
    )

    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Install-Module -Name $Name -AllowClobber:$AllowClobber -Scope CurrentUser -Force
    }

    Import-Module -Name $Name -ErrorAction Stop
}

function Export-SettingsToCsv {
    param(
        [Parameter(Mandatory = $true)]
        [string]$OutputDirectory,

        [Parameter(Mandatory = $true)]
        [string]$SectionName,

        [Parameter(Mandatory = $true)]
        $SettingsObject
    )

    if (-not (Test-Path -Path $OutputDirectory)) {
        New-Item -Path $OutputDirectory -ItemType Directory -Force | Out-Null
    }

    $csvFilePath = Join-Path -Path $OutputDirectory -ChildPath "$SectionName.csv"
    $SettingsObject | Export-Csv -Path $csvFilePath -NoTypeInformation -Encoding UTF8
    Write-Host "Exported $SectionName to $csvFilePath"
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

foreach ($module in $requiredModules) {
    Write-Host "Installing and importing module: $module"
    Install-AndImportModule -Name $module -AllowClobber
}

$sections = @(
    @{
        Name = "AzureADSettings"
        ScriptBlock = { Get-AzureADDirectorySetting }
    },
    @{
        Name = "MSOLSettings"
        ScriptBlock = { Get-MsolCompanyInformation }
    },
    @{
        Name = "ExchangeSettings"
        ScriptBlock = { Get-OrganizationConfig }
    },
    @{
        Name = "SharePointSettings"
        ScriptBlock = { Get-PnPWeb }
    },
    @{
        Name = "TeamsSettings"
        ScriptBlock = { Get-Team }
    },
    @{
        Name = "DeviceComplianceSettings"
        ScriptBlock = { Get-IntuneDeviceCompliancePolicy }
    },
    @{
        Name = "DeviceConfigurationSettings"
        ScriptBlock = { Get-IntuneDeviceConfigurationPolicy }
    },
    @{
        Name = "MobileAppSettings"
        ScriptBlock = { Get-IntuneMobileApp }
    },
    @{
        Name = "AppProtectionSettings"
        ScriptBlock = { Get-IntuneAppProtectionPolicy }
    },
    @{
        Name = "AppConfigurationSettings"
        ScriptBlock = { Get-MgIntuneAppConfigurationPolicy }
    }
)

Connect-AzureAD
Connect-MsolService

if ($ExchangeUserPrincipalName) {
    Connect-ExchangeOnline -UserPrincipalName $ExchangeUserPrincipalName
} else {
    Connect-ExchangeOnline
}

if ($SharePointAdminUrl) {
    Connect-PnPOnline -Url $SharePointAdminUrl -Credentials (Get-Credential)
} else {
    Write-Warning "Skipping SharePoint export because -SharePointAdminUrl was not provided."
}

Connect-MSGraph
Connect-MgGraph
Connect-MicrosoftTeams

foreach ($section in $sections) {
    try {
        if ($section.Name -eq "SharePointSettings" -and -not $SharePointAdminUrl) {
            continue
        }

        $settingsObject = & $section.ScriptBlock
        Export-SettingsToCsv -OutputDirectory $OutputDirectory -SectionName $section.Name -SettingsObject $settingsObject
    } catch {
        Write-Warning "Failed to retrieve $($section.Name): $($_.Exception.Message)"
    }
}
