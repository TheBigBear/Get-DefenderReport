<#
.SYNOPSIS
    Creates a DPAPI-protected XML file for storing email settings.

.DESCRIPTION
    This script prompts for email settings and saves them in an encrypted XML file using DPAPI.
    The encrypted file can be used by other scripts to securely retrieve email settings.

.EXAMPLE
    .\Create-EmailSettings.ps1 -To "me@company.com" -From "DefenderReport@company.com" -SmtpServer "mail-server.company.com" -Port 587 -UserName "DefenderReport@company.com" -Password "SuperSecretPasswordNumber12!"

.NOTES
    Author: Your Name
    Version: 1.0
    Date: 2023-10-10
#>

#region Parameters
param (
    [Parameter(Mandatory = $true)]
    [string]$To,

    [Parameter(Mandatory = $true)]
    [string]$From,

    [Parameter(Mandatory = $true)]
    [string]$SmtpServer,

    [Parameter(Mandatory = $true)]
    [int]$Port,

    [Parameter(Mandatory = $true)]
    [string]$UserName,

    [Parameter(Mandatory = $true)]
    [string]$Password
)
#endregion

#region Configuration
$outputDir = "C:\Downloads"
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}

$emailSettingsPath = Join-Path -Path $outputDir -ChildPath "DefenderReportEmailSettings.xml"
#endregion

#region Create Email Settings Object
$emailSettings = @{
    To         = $To
    From       = $From
    SmtpServer = $SmtpServer
    Port       = $Port
    UserName   = $UserName
    Password   = ConvertTo-SecureString -String $Password -AsPlainText -Force
}
#endregion

#region Save Email Settings to Encrypted XML
try {
    $emailSettings | Export-Clixml -Path $emailSettingsPath -Force
    Write-Host "Email settings saved to $emailSettingsPath"
} catch {
    Write-Error "Failed to save email settings: $_"
    exit 1
}
#endregion