<#
.SYNOPSIS
    Creates an HTML report of Microsoft Defender status for one or more servers.

.DESCRIPTION
    This script checks Microsoft Defender status for servers listed in a CSV file and generates individual HTML reports for each server.
    It also creates an overview report summarizing the status of all servers. Email reports can be sent if configured.

.EXAMPLE
    .\Get-DefenderReport.ps1 -CsvPath "C:\Downloads\defender-servers.csv" -Parallel 10

.NOTES
    Author: Adapted for PowerShell Core 7.x
    Version: 2.5
    Date: 2023-10-10
#>

#region Parameters
param (
    [string]$CsvPath = "C:\Downloads\defender-servers.csv", # Path to CSV file with server names
    [int]$Parallel = 5, # Number of servers to process in parallel
    [switch]$DebugMode, # Enable debug output
    [switch]$SendEmail # Send email report
)
#endregion

#region Check PowerShell Version
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Error "This script requires PowerShell Core 7.x or higher. Exiting."
    exit 1
}
#endregion

#region Configuration
$outputDir = "C:\Downloads"
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}

# Import the Report-Functions module
$modulePath = Join-Path -Path $outputDir -ChildPath "Report-Functions.psm1"
Import-Module -Name $modulePath -Force

# Load email settings from DPAPI-protected XML
$emailSettingsPath = Join-Path -Path $outputDir -ChildPath "DefenderReportEmailSettings.xml"
if (Test-Path $emailSettingsPath) {
    try {
        $emailSettings = Import-Clixml -Path $emailSettingsPath
    } catch {
        Write-Error "Failed to load email settings. Ensure the file exists and is properly encrypted. Exiting."
        exit 1
    }
} else {
    Write-Error "Email settings file not found. Exiting."
    exit 1
}
#endregion

#region Functions
function New-DefenderHTMLReport {
    param (
        [Parameter(Mandatory = $true)]
        [array]$DefenderData
    )

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
            <th>Server Name</th>
            <th>Defender Enabled</th>
            <th>Real-Time Protection</th>
            <th>Definition Age</th>
            <th>Last Full Scan</th>
            <th>Threats Found</th>
        </tr>
"@

    $htmlRows = ""
    foreach ($item in $DefenderData) {
        $rowClass = if ($item.ThreatsFound -gt 0) { "red" } elseif ($item.DefinitionAge -gt 5) { "orange" }
        $htmlRows += @"
        <tr class="$rowClass">
            <td>$($item.ServerName)</td>
            <td>$($item.DefenderEnabled)</td>
            <td>$($item.RealTimeProtection)</td>
            <td>$($item.DefinitionAge)</td>
            <td>$($item.LastFullScan)</td>
            <td>$($item.ThreatsFound)</td>
        </tr>
"@
    }

    $htmlFooter = @"
    </table>
</body>
</html>
"@

    return $htmlHeader + $htmlRows + $htmlFooter
}

function Send-EmailReport {
    param (
        [string]$Body,
        [string]$Subject
    )

    $emailParams = @{
        To         = $emailSettings.To
        From       = $emailSettings.From
        Subject    = $Subject
        Body       = $Body
        SmtpServer = $emailSettings.SmtpServer
        Port       = $emailSettings.Port
        Credential = New-Object System.Management.Automation.PSCredential -ArgumentList $emailSettings.UserName, ($emailSettings.Password | ConvertTo-SecureString)
        UseSsl     = $true
        BodyAsHtml = $true
    }

    try {
        Send-MailMessage @emailParams -ErrorAction Stop
        Write-Host "Email report sent successfully."
    } catch {
        Write-Error "Failed to send email report: $($_.Exception.Message)"
    }
}
#endregion

#region Main Script
# Read server names from CSV
try {
    $servers = Import-Csv -Path $CsvPath -Header "ServerName" | ForEach-Object { $_.ServerName }
    if (-not $servers) {
        Write-Error "No servers found in CSV file. Exiting."
        exit 1
    }
} catch {
    Write-Error "Failed to read CSV file: $($_.Exception.Message)"
    exit 1
}

# Process servers in parallel
$defenderResults = $servers | ForEach-Object -Parallel {
    # Define the Test-Online function inside the parallel block
    function Test-Online {
        param (
            [string]$ComputerName
        )

        try {
            $pingResult = Test-Connection -ComputerName $ComputerName -Count 2 -ErrorAction Stop
            return $pingResult.Status -eq "Success"
        } catch {
            Write-Warning "Failed to ping $ComputerName: $($_.Exception.Message)"
            return $false
        }
    }

    # Define the Get-DefenderStatus function inside the parallel block
    function Get-DefenderStatus {
        param (
            [string]$ServerName
        )

        try {
            $defenderStatus = Invoke-Command -ComputerName $ServerName -ScriptBlock {
                Get-MpComputerStatus -ErrorAction Stop
            } -ErrorAction Stop

            $defenderThreats = Invoke-Command -ComputerName $ServerName -ScriptBlock {
                Get-MpThreat -ErrorAction SilentlyContinue
            } -ErrorAction SilentlyContinue

            [PSCustomObject]@{
                ServerName           = $ServerName
                DefenderEnabled      = if ($defenderStatus.AntivirusEnabled) { "Enabled" } else { "Disabled" }
                RealTimeProtection   = if ($defenderStatus.RealTimeProtectionEnabled) { "Enabled" } else { "Disabled" }
                DefinitionAge        = "$($defenderStatus.AntivirusSignatureAge) days"
                LastFullScan         = if ($defenderStatus.FullScanEndTime) { $defenderStatus.FullScanEndTime.ToString() } else { "Never" }
                ThreatsFound         = if ($defenderThreats) { $defenderThreats.Count } else { "None" }
            }
        } catch {
            Write-Warning "Failed to retrieve Defender status for $ServerName: $($_.Exception.Message)"
            return $null
        }
    }

    $server = $_
    if ($using:DebugMode) {
        Write-Host "Processing server: $server"
    }

    if (Test-Online -ComputerName $server) {
        Get-DefenderStatus -ServerName $server
    } else {
        Write-Warning "Server $server is offline or unreachable."
        $null
    }
} -ThrottleLimit $Parallel

# Generate individual reports
foreach ($result in $defenderResults) {
    if ($result) {
        $timestamp = Get-Date -Format "yyyy-MM-dd-HH-mm-ss"
        $htmlReportPath = Join-Path -Path $outputDir -ChildPath "$($result.ServerName)_$timestamp.html"
        $htmlReport = New-DefenderHTMLReport -DefenderData @($result)
        $htmlReport | Out-File -FilePath $htmlReportPath -Encoding UTF8
        Write-Host "Report saved for $($result.ServerName) at $htmlReportPath"
    }
}

# Generate overview report
$overviewReportPath = Join-Path -Path $outputDir -ChildPath "Overview-Defender-Report_$(Get-Date -Format 'yyyy-MM-dd-HH-mm-ss').html"
$overviewReport = New-DefenderHTMLReport -DefenderData $defenderResults
$overviewReport | Out-File -FilePath $overviewReportPath -Encoding UTF8
Write-Host "Overview report saved at $overviewReportPath"

# Send email report if enabled
if ($SendEmail) {
    Send-EmailReport -Body $overviewReport -Subject "Microsoft Defender Overview Report"
}
#endregion