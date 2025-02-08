<#
.SYNOPSIS
    Creates an HTML report of the local server's Microsoft Defender status.

.DESCRIPTION
    This script checks the local server's Microsoft Defender status and generates an HTML report.
    The report includes information such as whether Defender is enabled, the status of the Defender service,
    the age of virus definitions, the last full scan date, and any detected threats.

.EXAMPLE
    .\Get-DefenderReport.ps1

.NOTES
    Author: Jason Dillman (adapted for local use)
    Version: 1.0 (Localized)
    Date: 2023-10-10
#>

# Set the output directory
$outputDir = "C:\Downloads"
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}

# Import the Report-Functions module
$modulePath = Join-Path -Path $outputDir -ChildPath "Report-Functions.psm1"
Import-Module -Name $modulePath -Force

# Get the FQDN of the server
$serverFQDN = [System.Net.Dns]::GetHostByName($env:COMPUTERNAME).HostName

# Generate a timestamp for the report filename
$timestamp = Get-Date -Format "yyyy-MM-dd-HH-mm-ss"

# HTML report file path
$htmlReportPath = Join-Path -Path $outputDir -ChildPath "$serverFQDN`_$timestamp.html"

# Get Microsoft Defender status
try {
    $defenderStatus = Get-MpComputerStatus -ErrorAction Stop
    $defenderThreats = Get-MpThreat -ErrorAction SilentlyContinue
} catch {
    Write-Error "Failed to retrieve Microsoft Defender status. Ensure Microsoft Defender is installed and running."
    exit
}

# Prepare Defender data for the report
$defenderData = @(
    [PSCustomObject]@{
        'Computer Name'       = $env:COMPUTERNAME
        'Defender Enabled'    = if ($defenderStatus.AntivirusEnabled) { "Enabled" } else { "Disabled" }
        'Real-Time Protection'= if ($defenderStatus.RealTimeProtectionEnabled) { "Enabled" } else { "Disabled" }
        'Definition Age'      = "$($defenderStatus.AntivirusSignatureAge) days"
        'Last Full Scan'      = if ($defenderStatus.FullScanEndTime) { $defenderStatus.FullScanEndTime.ToString() } else { "Never" }
        'Threats Found'       = if ($defenderThreats) { $defenderThreats.Count } else { "None" }
        'Color Coding'        = @(
            @{ Key = 'Defender Enabled'; Value = if (-not $defenderStatus.AntivirusEnabled) { 'ff0000' } else { '00ff00' } },
            @{ Key = 'Real-Time Protection'; Value = if (-not $defenderStatus.RealTimeProtectionEnabled) { 'ff0000' } else { '00ff00' } },
            @{ Key = 'Definition Age'; Value = if ($defenderStatus.AntivirusSignatureAge -gt 5) { 'ff7d00' } else { '00ff00' } },
            @{ Key = 'Last Full Scan'; Value = if ($defenderStatus.FullScanEndTime -lt (Get-Date).AddDays(-14)) { 'ff7d00' } else { '00ff00' } },
            @{ Key = 'Threats Found'; Value = if ($defenderThreats) { 'ff0000' } else { '00ff00' } }
        )
        'Primary Column Name' = 'Computer Name'
        'Sort'                = 0
    }
)

# Generate the HTML report
$htmlReport = $defenderData | New-HTMLReport

# Save the HTML report to file
$htmlReport | Out-File -FilePath $htmlReportPath -Encoding UTF8

Write-Host "Microsoft Defender status report saved to: $htmlReportPath"