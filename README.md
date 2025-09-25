# Microsoft Services Hub Case Reporter

A PowerShell solution for retrieving and reporting on Microsoft support cases directly from the **Microsoft Services Hub** (serviceshub.microsoft.com) instead of the standard Azure portal.

## 🎯 Features

- **Direct Services Hub Integration**: Connects to Microsoft Services Hub API for comprehensive case data
- **Stale Case Detection**: Automatically flags cases not updated in 48+ hours
- **Multiple Report Formats**: Generates CSV, HTML, and console reports
- **Real-time Dashboard**: Console summary with color-coded warnings
- **Enterprise Security**: Secure authentication using Azure AD service principals
- **Comprehensive Logging**: Detailed execution logs and error handling

## 📋 Prerequisites

- PowerShell 5.0 or higher
- Azure AD application with Services Hub API permissions
- Access to Microsoft Services Hub (serviceshub.microsoft.com)

## 🚀 Quick Start

### 1. Run Setup Script
```powershell
.\Setup-ServicesHubPrerequisites.ps1
```

### 2. Configure Azure AD Application

1. **Register Application in Azure AD:**
   - Go to [Azure Portal](https://portal.azure.com)
   - Navigate to **Azure Active Directory** > **App registrations**
   - Click **New registration**
   - Name: `Services Hub Case Reporter`
   - Account types: `Accounts in this organizational directory only`
   - Click **Register**

2. **Configure API Permissions:**
   - Go to **API permissions** > **Add a permission**
   - Add Microsoft Services Hub API permissions:
     - `ServiceHub.Read.All`
     - `ServiceHub.ReadWrite.All`
   - Click **Grant admin consent for [Your Organization]**

3. **Create Client Secret:**
   - Go to **Certificates & secrets** > **New client secret**
   - Description: `Services Hub Reporter Secret`
   - Expires: `12 months` (recommended)
   - Copy the **Value** (not the Secret ID)

4. **Update Configuration:**
   - Copy **Application (client) ID** from Overview page
   - Copy **Directory (tenant) ID** from Overview page
   - Update `config.json` with these values

### 3. Update Configuration File

Edit `config.json`:

```json
{
  "ServicesHub": {
    "TenantId": "your-azure-tenant-id-here",
    "ClientId": "your-application-client-id-here", 
    "ClientSecret": "your-application-client-secret-here",
    "Scope": "https://serviceshub.microsoft.com/.default"
  },
  "ReportSettings": {
    "IncludeClosedCases": false,
    "StaleThresholdHours": 48,
    "OutputFormats": ["CSV", "HTML", "Console"]
  }
}
```

### 4. Run the Report Generator

```powershell
.\MicrosoftServicesHubCaseReport.ps1
```

## 📊 Report Output

### Console Dashboard
- Real-time summary with total, recent, and stale case counts
- Color-coded warnings for cases requiring immediate attention
- Quick overview of most critical cases

### CSV Export
- Machine-readable format for data analysis
- All case details with calculated stale flags
- Timestamped filename for version tracking

### HTML Report
- Professional formatted report with styling
- Color-coded rows for stale cases (red background)
- Clickable case numbers linking to Services Hub
- Summary statistics and visual indicators

## 🔧 Report Fields

Each case report includes:

| Field | Description |
|-------|-------------|
| CaseNumber | Unique Microsoft case identifier |
| Title | Case title/summary |
| Description | Detailed case description |
| Owner | Assigned case owner name |
| OwnerEmail | Owner's email address |
| Status | Current case status (Open, In Progress, etc.) |
| Severity | Case severity level |
| CreatedDate | When the case was originally created |
| LastUpdatedDate | Most recent case activity timestamp |
| HoursSinceUpdate | Calculated hours since last update |
| IsStale | Boolean flag for 48+ hour threshold |
| StaleFlag | Visual indicator (⚠️ STALE or ✅ Recent) |
| Customer | Customer/organization name |
| ProductName | Microsoft product/service affected |
| CaseUrl | Direct link to case in Services Hub |

## ⚠️ Stale Case Detection

Cases are flagged as **STALE** when:
- No updates for 48+ hours (configurable)
- Status is "Waiting for Customer" but customer responded
- High/Critical severity cases inactive for 24+ hours

Stale cases are highlighted in:
- ❌ Console output (red text)
- 🔴 HTML report (red background rows)
- 📊 Summary statistics

## 🔐 Security Best Practices

- **No Hardcoded Credentials**: All sensitive data in config file
- **Client Secret Protection**: Store securely, rotate regularly
- **Least Privilege**: Only required API permissions granted
- **Secure Authentication**: OAuth 2.0 client credentials flow
- **Audit Logging**: All API calls logged with timestamps

## 📁 File Structure

```
MSFT_Reporting/
├── MicrosoftServicesHubCaseReport.ps1    # Main reporting script
├── Setup-ServicesHubPrerequisites.ps1    # Setup and configuration
├── config.json                           # Configuration file
├── README.md                             # This documentation
└── Reports/                              # Generated reports directory
    ├── ServicesHub_Cases_20250925_120151.csv
    ├── ServicesHub_Cases_20250925_120151.html
    └── ServicesHub_CaseReport_20250925.log
```

## 🛠️ Advanced Usage

### Custom Threshold Configuration
```powershell
# Run with custom stale threshold (24 hours)
.\MicrosoftServicesHubCaseReport.ps1 -ConfigPath ".\custom-config.json"
```

### Specific Output Directory
```powershell
# Save reports to specific location
.\MicrosoftServicesHubCaseReport.ps1 -OutputPath "C:\Reports\ServicesHub"
```

### Scheduled Execution
```powershell
# Create scheduled task for daily reports
$Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-File C:\Scripts\MicrosoftServicesHubCaseReport.ps1"
$Trigger = New-ScheduledTaskTrigger -Daily -At "08:00AM"
Register-ScheduledTask -Action $Action -Trigger $Trigger -TaskName "Services Hub Daily Report"
```

## 🔍 Troubleshooting

### Common Issues

**Authentication Errors:**
- Verify tenant ID, client ID, and client secret in config.json
- Ensure API permissions are granted and admin consent provided
- Check if application is enabled and not expired

**No Cases Retrieved:**
- Confirm access to Microsoft Services Hub
- Verify API permissions include `ServiceHub.Read.All`
- Check if organization has active support cases

**Permission Denied:**
- Ensure Azure AD application has proper Services Hub permissions
- Verify user running script has access to Services Hub
- Check if MFA or conditional access policies are blocking

### Debug Mode
Enable detailed logging by modifying the script:
```powershell
$VerbosePreference = "Continue"
$DebugPreference = "Continue"
```

## 📞 Support & Contacts

For technical support:
- Script Issues: Check execution logs in Reports directory
- API Issues: Verify Services Hub access and permissions
- Azure AD Issues: Contact your Azure administrator

## 📝 Version History

- **v1.0** - Initial release with Services Hub integration
- Direct API connection to serviceshub.microsoft.com
- 48-hour stale case detection
- Multiple output formats (CSV, HTML, Console)
- Comprehensive error handling and logging

## 🤝 Contributing

To improve this solution:
1. Test with your Services Hub environment
2. Report issues or enhancement requests
3. Submit improvements via pull requests
4. Share feedback on report formats and features

---

**Last Updated:** 2025-09-25 12:01:51 UTC  
**Target Environment:** Microsoft Services Hub (serviceshub.microsoft.com)  
**PowerShell Version:** 5.0+
