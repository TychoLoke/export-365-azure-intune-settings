param(
    [string]$OutputDirectory = "C:\Temp\M365-Exports",
    [string[]]$Sections,
    [string]$ExchangeUserPrincipalName,
    [string]$SharePointAdminUrl
)

$scriptPath = Join-Path -Path $PSScriptRoot -ChildPath "Export-M365TenantSettings.ps1"
if (-not (Test-Path -Path $scriptPath)) {
    throw "Export-M365TenantSettings.ps1 was not found next to this wrapper script."
}

$invokeParams = @{
    OutputDirectory = $OutputDirectory
}

if ($Sections) {
    $invokeParams.Sections = $Sections
}

if ($ExchangeUserPrincipalName) {
    $invokeParams.ExchangeUserPrincipalName = $ExchangeUserPrincipalName
}

if ($SharePointAdminUrl) {
    $invokeParams.SharePointAdminUrl = $SharePointAdminUrl
}

& $scriptPath @invokeParams
