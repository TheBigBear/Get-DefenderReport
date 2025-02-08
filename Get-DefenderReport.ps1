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
            [string]${ComputerName}
        )

        try {
            $pingResult = Test-Connection -ComputerName ${ComputerName} -Count 2 -ErrorAction Stop
            return $pingResult.Status -eq "Success"
        } catch {
            Write-Warning "Failed to ping ${ComputerName}: $($_.Exception.Message)"
            return $false
        }
    }

    # Define the Get-DefenderStatus function inside the parallel block
    function Get-DefenderStatus {
        param (
            [string]${ServerName}
        )

        try {
            $defenderStatus = Invoke-Command -ComputerName ${ServerName} -ScriptBlock {
                Get-MpComputerStatus -ErrorAction Stop
            } -ErrorAction Stop

            $defenderThreats = Invoke-Command -ComputerName ${ServerName} -ScriptBlock {
                Get-MpThreat -ErrorAction SilentlyContinue
            } -ErrorAction SilentlyContinue

            [PSCustomObject]@{
                ServerName           = ${ServerName}
                DefenderEnabled      = if ($defenderStatus.AntivirusEnabled) { "Enabled" } else { "Disabled" }
                RealTimeProtection   = if ($defenderStatus.RealTimeProtectionEnabled) { "Enabled" } else { "Disabled" }
                DefinitionAge        = "$($defenderStatus.AntivirusSignatureAge) days"
                LastFullScan         = if ($defenderStatus.FullScanEndTime) { $defenderStatus.FullScanEndTime.ToString() } else { "Never" }
                ThreatsFound         = if ($defenderThreats) { $defenderThreats.Count } else { "None" }
            }
        } catch {
            Write-Warning "Failed to retrieve Defender status for ${ServerName}: $($_.Exception.Message)"
            return $null
        }
    }

    $server = $_
    if ($using:DebugMode) {
        Write-Host "Processing server: $server"
    }

    if (Test-Online -ComputerName ${server}) {
        Get-DefenderStatus -ServerName ${server}
    } else {
        Write-Warning "Server ${server} is offline or unreachable."
        $null
    }
} -ThrottleLimit $Parallel

# Check if $defenderResults is null or empty
if (-not $defenderResults) {
    Write-Warning "No Defender status data was retrieved. Check the CSV file and ensure the servers are online and accessible."
    exit 1
}

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