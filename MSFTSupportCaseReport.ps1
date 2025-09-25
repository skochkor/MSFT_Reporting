<#
.SYNOPSIS
    Microsoft Support Cases Report Generator
    
.DESCRIPTION
    Connects to Microsoft Support API via Microsoft Graph to retrieve all open support cases
    and generate comprehensive reports with stale case detection (48+ hours without updates).
    
.PARAMETER ConfigPath
    Path to configuration file containing authentication details
    
.PARAMETER OutputPath
    Directory path for generated reports
    
.PARAMETER StaleHours
    Number of hours to consider a case stale (default: 48)
    
.EXAMPLE
    .\MicrosoftSupportCasesReport.ps1
    
.EXAMPLE
    .\MicrosoftSupportCasesReport.ps1 -ConfigPath ".\config.json" -OutputPath "C:\Reports" -StaleHours 72
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$ConfigPath = ".\config.json",
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".\Reports",
    
    [Parameter(Mandatory = $false)]
    [int]$StaleHours = 48
)

# Import required modules
$RequiredModules = @('Microsoft.Graph.Authentication', 'Microsoft.Graph.ServiceCommunications')
foreach ($Module in $RequiredModules) {
    if (!(Get-Module -Name $Module -ListAvailable)) {
        Write-Error "Required module '$Module' not found. Please run Setup-Prerequisites.ps1 first."
        exit 1
    }
    Import-Module $Module -Force
}

# Initialize logging
$LogFile = Join-Path $OutputPath "MSSupport_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$ErrorActionPreference = 'Stop'

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$Timestamp] [$Level] $Message"
    Write-Host $LogEntry
    Add-Content -Path $LogFile -Value $LogEntry
}

function Connect-ToMicrosoftGraph {
    param([hashtable]$Config)
    
    Write-Log "Connecting to Microsoft Graph..."
    
    try {
        if ($Config.AuthenticationMethod -eq "ServicePrincipal") {
            $SecureClientSecret = ConvertTo-SecureString $Config.ClientSecret -AsPlainText -Force
            $Credential = New-Object System.Management.Automation.PSCredential($Config.ClientId, $SecureClientSecret)
            
            Connect-MgGraph -TenantId $Config.TenantId -ClientSecretCredential $Credential -Scopes "ServiceHealth.Read.All"
        }
        else {
            # Interactive authentication
            Connect-MgGraph -Scopes "ServiceHealth.Read.All" -TenantId $Config.TenantId
        }
        
        Write-Log "Successfully connected to Microsoft Graph"
        return $true
    }
    catch {
        Write-Log "Failed to connect to Microsoft Graph: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

function Get-MicrosoftSupportCases {
    Write-Log "Retrieving Microsoft Support cases..."
    
    try {
        # Get service health incidents and issues which represent support cases
        $ServiceIssues = @()
        
        # Get current service health issues
        $CurrentIssues = Get-MgServiceAnnouncementIssue -All
        foreach ($Issue in $CurrentIssues) {
            if ($Issue.Status -in @('serviceOperational', 'investigating', 'restoringService', 'extended recovery')) {
                $ServiceIssues += [PSCustomObject]@{
                    CaseNumber = $Issue.Id
                    Owner = if ($Issue.ImpactDescription) { "Microsoft Support" } else { "System Generated" }
                    Title = $Issue.Title
                    Description = $Issue.ImpactDescription -replace '<[^>]+>', '' # Remove HTML tags
                    Status = $Issue.Status
                    Severity = $Issue.Classification
                    CreatedDate = $Issue.StartDateTime
                    LastUpdated = $Issue.LastModifiedDateTime
                    Service = $Issue.Service
                    Feature = $Issue.Feature
                }
            }
        }
        
        # Get service health incidents
        $Incidents = Get-MgServiceAnnouncementIncident -All
        foreach ($Incident in $Incidents) {
            if ($Incident.Status -in @('investigating', 'serviceRestoration', 'postIncidentReviewPublished')) {
                $ServiceIssues += [PSCustomObject]@{
                    CaseNumber = $Incident.Id
                    Owner = "Microsoft Incident Response"
                    Title = $Incident.Title
                    Description = $Incident.ImpactDescription -replace '<[^>]+>', '' # Remove HTML tags
                    Status = $Incident.Status
                    Severity = $Incident.Classification
                    CreatedDate = $Incident.StartDateTime
                    LastUpdated = $Incident.LastModifiedDateTime
                    Service = $Incident.Service
                    Feature = $Incident.Feature
                }
            }
        }
        
        Write-Log "Retrieved $($ServiceIssues.Count) support cases/incidents"
        return $ServiceIssues
    }
    catch {
        Write-Log "Error retrieving support cases: $($_.Exception.Message)" -Level "ERROR"
        return @()
    }
}

function Add-StaleFlags {
    param([array]$Cases, [int]$StaleHours)
    
    $CurrentTime = Get-Date
    $StaleThreshold = $CurrentTime.AddHours(-$StaleHours)
    
    foreach ($Case in $Cases) {
        $LastUpdate = [DateTime]$Case.LastUpdated
        $HoursSinceUpdate = [math]::Round(($CurrentTime - $LastUpdate).TotalHours, 1)
        
        $Case | Add-Member -NotePropertyName "HoursSinceUpdate" -NotePropertyValue $HoursSinceUpdate
        $Case | Add-Member -NotePropertyName "IsStale" -NotePropertyValue ($LastUpdate -lt $StaleThreshold)
        $Case | Add-Member -NotePropertyName "StatusFlag" -NotePropertyValue $(
            if ($LastUpdate -lt $StaleThreshold) { "⚠️ STALE" } else { "✅ Recent" }
        )
    }
    
    return $Cases
}

function Export-ToCSV {
    param([array]$Cases, [string]$OutputPath)
    
    $CSVFile = Join-Path $OutputPath "MSSupport_Cases_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    
    try {
        $Cases | Select-Object CaseNumber, Owner, Title, Description, Status, Severity, 
                             CreatedDate, LastUpdated, HoursSinceUpdate, IsStale, StatusFlag, 
                             Service, Feature | 
        Export-Csv -Path $CSVFile -NoTypeInformation -Encoding UTF8
        
        Write-Log "CSV report exported to: $CSVFile"
        return $CSVFile
    }
    catch {
        Write-Log "Error exporting CSV: $($_.Exception.Message)" -Level "ERROR"
        return $null
    }
}

function Export-ToHTML {
    param([array]$Cases, [string]$OutputPath, [int]$StaleHours)
    
    $HTMLFile = Join-Path $OutputPath "MSSupport_Cases_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    $CurrentDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss UTC"
    
    $StaleCases = $Cases | Where-Object { $_.IsStale }
    $RecentCases = $Cases | Where-Object { -not $_.IsStale }
    
    $HTMLContent = @"
<!DOCTYPE html>
<html>
<head>
    <title>Microsoft Support Cases Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .header { background-color: #0078d4; color: white; padding: 15px; border-radius: 5px; }
        .summary { background-color: #f3f2f1; padding: 15px; margin: 20px 0; border-radius: 5px; }
        .stale { background-color: #fef9e7; border-left: 4px solid #ff8c00; }
        .recent { background-color: #f0fff0; border-left: 4px solid: #32cd32; }
        .critical { background-color: #ffe6e6; border-left: 4px solid #dc3545; }
        table { border-collapse: collapse; width: 100%; margin-top: 20px; }
        th, td { border: 1px solid #ddd; padding: 12px; text-align: left; }
        th { background-color: #0078d4; color: white; }
        tr:nth-child(even) { background-color: #f9f9f9; }
        .status-stale { color: #ff8c00; font-weight: bold; }
        .status-recent { color: #32cd32; font-weight: bold; }
    </style>
</head>
<body>
    <div class="header">
        <h1>📊 Microsoft Support Cases Report</h1>
        <p>Generated on: $CurrentDate | Stale Threshold: $StaleHours hours</p>
    </div>
    
    <div class="summary">
        <h2>📈 Summary Statistics</h2>
        <p><strong>Total Cases:</strong> $($Cases.Count)</p>
        <p><strong>Stale Cases (>$StaleHours hours):</strong> <span style="color: #ff8c00;">$($StaleCases.Count)</span></p>
        <p><strong>Recent Cases:</strong> <span style="color: #32cd32;">$($RecentCases.Count)</span></p>
    </div>
"@

    if ($StaleCases.Count -gt 0) {
        $HTMLContent += @"
    <div class="stale">
        <h2>⚠️ Stale Cases Requiring Attention</h2>
        <table>
            <tr>
                <th>Case Number</th>
                <th>Owner</th>
                <th>Title</th>
                <th>Status</th>
                <th>Severity</th>
                <th>Hours Since Update</th>
                <th>Last Updated</th>
            </tr>
"@
        foreach ($Case in $StaleCases) {
            $HTMLContent += @"
            <tr>
                <td>$($Case.CaseNumber)</td>
                <td>$($Case.Owner)</td>
                <td>$($Case.Title)</td>
                <td>$($Case.Status)</td>
                <td>$($Case.Severity)</td>
                <td class="status-stale">$($Case.HoursSinceUpdate)</td>
                <td>$($Case.LastUpdated)</td>
            </tr>
"@
        }
        $HTMLContent += "</table></div>"
    }

    $HTMLContent += @"
    <div class="recent">
        <h2>✅ All Cases Overview</h2>
        <table>
            <tr>
                <th>Case Number</th>
                <th>Owner</th>
                <th>Title</th>
                <th>Status</th>
                <th>Severity</th>
                <th>Hours Since Update</th>
                <th>Status Flag</th>
                <th>Service</th>
            </tr>
"@

    foreach ($Case in $Cases | Sort-Object IsStale -Descending, HoursSinceUpdate -Descending) {
        $StatusClass = if ($Case.IsStale) { "status-stale" } else { "status-recent" }
        $HTMLContent += @"
            <tr>
                <td>$($Case.CaseNumber)</td>
                <td>$($Case.Owner)</td>
                <td>$($Case.Title)</td>
                <td>$($Case.Status)</td>
                <td>$($Case.Severity)</td>
                <td class="$StatusClass">$($Case.HoursSinceUpdate)</td>
                <td>$($Case.StatusFlag)</td>
                <td>$($Case.Service)</td>
            </tr>
"@
    }

    $HTMLContent += @"
        </table>
    </div>
</body>
</html>
"@

    try {
        Set-Content -Path $HTMLFile -Value $HTMLContent -Encoding UTF8
        Write-Log "HTML report exported to: $HTMLFile"
        return $HTMLFile
    }
    catch {
        Write-Log "Error exporting HTML: $($_.Exception.Message)" -Level "ERROR"
        return $null
    }
}

function Show-ConsoleSummary {
    param([array]$Cases, [int]$StaleHours)
    
    $StaleCases = $Cases | Where-Object { $_.IsStale }
    $RecentCases = $Cases | Where-Object { -not $_.IsStale }
    
    Write-Host "`n" -NoNewline
    Write-Host "=" * 60 -ForegroundColor Cyan
    Write-Host "📊 MICROSOFT SUPPORT CASES SUMMARY" -ForegroundColor White -BackgroundColor DarkBlue
    Write-Host "=" * 60 -ForegroundColor Cyan
    Write-Host "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss UTC')" -ForegroundColor Gray
    Write-Host "Stale Threshold: $StaleHours hours" -ForegroundColor Gray
    Write-Host ""
    
    Write-Host "📈 STATISTICS:" -ForegroundColor Yellow
    Write-Host "   Total Cases: " -NoNewline -ForegroundColor White
    Write-Host "$($Cases.Count)" -ForegroundColor Green
    Write-Host "   Stale Cases: " -NoNewline -ForegroundColor White
    Write-Host "$($StaleCases.Count)" -ForegroundColor $(if ($StaleCases.Count -gt 0) { "Red" } else { "Green" })
    Write-Host "   Recent Cases: " -NoNewline -ForegroundColor White
    Write-Host "$($RecentCases.Count)" -ForegroundColor Green
    
    if ($StaleCases.Count -gt 0) {
        Write-Host ""
        Write-Host "⚠️  CASES REQUIRING IMMEDIATE ATTENTION:" -ForegroundColor Red -BackgroundColor Yellow
        foreach ($Case in $StaleCases | Sort-Object HoursSinceUpdate -Descending | Select-Object -First 5) {
            Write-Host "   Case: $($Case.CaseNumber) | Owner: $($Case.Owner) | Stale: $($Case.HoursSinceUpdate)h" -ForegroundColor Red
        }
        if ($StaleCases.Count -gt 5) {
            Write-Host "   ... and $($StaleCases.Count - 5) more stale cases" -ForegroundColor Yellow
        }
    }
    else {
        Write-Host ""
        Write-Host "✅ All cases have been updated within the last $StaleHours hours!" -ForegroundColor Green
    }
    
    Write-Host "=" * 60 -ForegroundColor Cyan
}

# Main execution
try {
    Write-Log "Starting Microsoft Support Cases Report Generator"
    Write-Log "Current UTC Date/Time: 2025-09-25 11:45:33"
    
    # Create output directory
    if (!(Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
        Write-Log "Created output directory: $OutputPath"
    }
    
    # Load configuration
    if (Test-Path $ConfigPath) {
        $Config = Get-Content $ConfigPath | ConvertFrom-Json
        Write-Log "Configuration loaded from: $ConfigPath"
    }
    else {
        Write-Log "Configuration file not found. Using environment variables..." -Level "WARNING"
        $Config = @{
            TenantId = $env:MS_TENANT_ID
            ClientId = $env:MS_CLIENT_ID
            ClientSecret = $env:MS_CLIENT_SECRET
            AuthenticationMethod = if ($env:MS_CLIENT_SECRET) { "ServicePrincipal" } else { "Interactive" }
        }
    }
    
    # Connect to Microsoft Graph
    if (!(Connect-ToMicrosoftGraph -Config $Config)) {
        throw "Failed to establish connection to Microsoft Graph"
    }
    
    # Retrieve support cases
    $Cases = Get-MicrosoftSupportCases
    if ($Cases.Count -eq 0) {
        Write-Log "No support cases found or retrieved" -Level "WARNING"
        return
    }
    
    # Add stale flags
    $Cases = Add-StaleFlags -Cases $Cases -StaleHours $StaleHours
    
    # Generate reports
    $CSVFile = Export-ToCSV -Cases $Cases -OutputPath $OutputPath
    $HTMLFile = Export-ToHTML -Cases $Cases -OutputPath $OutputPath -StaleHours $StaleHours
    
    # Display console summary
    Show-ConsoleSummary -Cases $Cases -StaleHours $StaleHours
    
    Write-Log "Report generation completed successfully"
    Write-Log "CSV Report: $CSVFile"
    Write-Log "HTML Report: $HTMLFile"
    
    # Disconnect from Microsoft Graph
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    Write-Log "Disconnected from Microsoft Graph"
}
catch {
    Write-Log "Script execution failed: $($_.Exception.Message)" -Level "ERROR"
    Write-Log "Stack Trace: $($_.ScriptStackTrace)" -Level "ERROR"
    exit 1
}
finally {
    # Cleanup
    if (Get-MgContext) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
    }
}
