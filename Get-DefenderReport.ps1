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

# HTML report file path
$htmlReportPath = Join-Path -Path $outputDir -ChildPath "DefenderStatusReport.html"

# Function to create the HTML report
function New-DefenderHTMLReport {
    param (
        [Parameter(Mandatory = $true)]
        [array]$DefenderData
    )

    # HTML header and style
    $htmlHeader = @"
<!DOCTYPE html>
<html>
<head>
    <title>Microsoft Defender Status Report</title>
    <style>
        body { font-family: Arial, sans-serif; }
        table { border-collapse: collapse; width: 100%; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
        tr:nth-child(even) { background-color: #f9f9f9; }
        tr:hover { background-color: #f1f1f1; }
        .red { background-color: #ffcccc; }
        .orange { background-color: #ffd699; }
    </style>
</head>
<body>
    <h1>Microsoft Defender Status Report</h1>
    <p>Generated on: $(Get-Date)</p>
    <table>
        <tr>
            <th>Property</th>
            <th>Value</th>
        </tr>
"@

    # HTML rows for Defender data
    $htmlRows = ""
    foreach ($item in $DefenderData) {
        $rowClass = ""
        if ($item.Value -like "*Disabled*" -or $item.Value -like "*Threat Found*") {
            $rowClass = "red"
        } elseif ($item.Value -like "*Outdated*") {
            $rowClass = "orange"
        }
        $htmlRows += @"
        <tr class="$rowClass">
            <td>$($item.Key)</td>
            <td>$($item.Value)</td>
        </tr>
"@
    }

    # HTML footer
    $htmlFooter = @"
    </table>
</body>
</html>
"@

    # Combine HTML parts
    $htmlReport = $htmlHeader + $htmlRows + $htmlFooter
    return $htmlReport
}

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
    @{ Key = "Defender Enabled"; Value = if ($defenderStatus.AntivirusEnabled) { "Enabled" } else { "Disabled" } },
    @{ Key = "Real-Time Protection"; Value = if ($defenderStatus.RealTimeProtectionEnabled) { "Enabled" } else { "Disabled" } },
    @{ Key = "Antivirus Definitions Age"; Value = "$($defenderStatus.AntivirusSignatureAge) days" },
    @{ Key = "Last Full Scan"; Value = if ($defenderStatus.FullScanEndTime) { $defenderStatus.FullScanEndTime.ToString() } else { "Never" } },
    @{ Key = "Threats Found"; Value = if ($defenderThreats) { $defenderThreats.Count } else { "None" } }
)

# Generate the HTML report
$htmlReport = New-DefenderHTMLReport -DefenderData $defenderData

# Save the HTML report to file
$htmlReport | Out-File -FilePath $htmlReportPath -Encoding UTF8

Write-Host "Microsoft Defender status report saved to: $htmlReportPath"