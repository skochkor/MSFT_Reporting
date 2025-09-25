<#
.SYNOPSIS
    Setup script for Microsoft Support Cases Report Generator
    
.DESCRIPTION
    Installs required PowerShell modules and sets up the environment
    for the Microsoft Support Cases reporting tool.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [switch]$Force
)

Write-Host "🔧 Microsoft Support Cases Report - Prerequisites Setup" -ForegroundColor Cyan
Write-Host "=" * 60

# Check PowerShell version
$PSVersion = $PSVersionTable.PSVersion
Write-Host "PowerShell Version: $PSVersion" -ForegroundColor Green

if ($PSVersion.Major -lt 5) {
    Write-Error "PowerShell 5.0 or higher is required. Current version: $PSVersion"
    exit 1
}

# Required modules
$RequiredModules = @(
    @{ Name = "Microsoft.Graph.Authentication"; MinVersion = "1.19.0" },
    @{ Name = "Microsoft.Graph.ServiceCommunications"; MinVersion = "1.19.0" }
)

Write-Host "`n📦 Installing Required Modules..."

foreach ($Module in $RequiredModules) {
    Write-Host "Checking module: $($Module.Name)" -ForegroundColor Yellow
    
    $InstalledModule = Get-Module -Name $Module.Name -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
    
    if ($InstalledModule -and $InstalledModule.Version -ge $Module.MinVersion -and -not $Force) {
        Write-Host "✅ $($Module.Name) v$($InstalledModule.Version) is already installed" -ForegroundColor Green
    }
    else {
        try {
            Write-Host "📥 Installing $($Module.Name)..." -ForegroundColor Yellow
            Install-Module -Name $Module.Name -MinimumVersion $Module.MinVersion -Force -AllowClobber -Scope CurrentUser
            Write-Host "✅ Successfully installed $($Module.Name)" -ForegroundColor Green
        }
        catch {
            Write-Error "Failed to install $($Module.Name): $($_.Exception.Message)"
            exit 1
        }
    }
}

# Create configuration template
Write-Host "`n📄 Creating Configuration Template..."

$ConfigTemplate = @{
    TenantId = "your-tenant-id-here"
    ClientId = "your-client-id-here" 
    ClientSecret = "your-client-secret-here"
    AuthenticationMethod = "ServicePrincipal"
    Notes = @{
        ServicePrincipal = "Use for unattended/scheduled execution"
        Interactive = "Use for manual execution with user login"
        RequiredPermissions = @(
            "ServiceHealth.Read.All",
            "ServiceMessage.Read.All"
        )
    }
}

$ConfigPath = ".\config-template.json"
$ConfigTemplate | ConvertTo-Json -Depth 3 | Set-Content -Path $ConfigPath -Encoding UTF8

Write-Host "✅ Configuration template created: $ConfigPath" -ForegroundColor Green

# Create Reports directory
$ReportsPath = ".\Reports"
if (!(Test-Path $ReportsPath)) {
    New-Item -ItemType Directory -Path $ReportsPath -Force | Out-Null
    Write-Host "✅ Created Reports directory: $ReportsPath" -ForegroundColor Green
}

# Display setup completion and next steps
Write-Host "`n" -NoNewline
Write-Host "🎉 SETUP COMPLETED SUCCESSFULLY!" -ForegroundColor White -BackgroundColor Green
Write-Host "`n📋 NEXT STEPS:"
Write-Host "1. Register an Azure AD Application:" -ForegroundColor Yellow
Write-Host "   - Go to https://portal.azure.com" -ForegroundColor Gray
Write-Host "   - Navigate to Azure Active Directory > App registrations" -ForegroundColor Gray
Write-Host "   - Click 'New registration'" -ForegroundColor Gray
Write-Host "   - Name: 'Microsoft Support Cases Reporter'" -ForegroundColor Gray
Write-Host "   - Account types: 'Accounts in this organizational directory only'" -ForegroundColor Gray

Write-Host "`n2. Configure API Permissions:" -ForegroundColor Yellow
Write-Host "   - Add Application permissions:" -ForegroundColor Gray
Write-Host "     • ServiceHealth.Read.All" -ForegroundColor Gray
Write-Host "     • ServiceMessage.Read.All" -ForegroundColor Gray
Write-Host "   - Grant admin consent" -ForegroundColor Gray

Write-Host "`n3. Create Client Secret:" -ForegroundColor Yellow
Write-Host "   - Go to 'Certificates & secrets'" -ForegroundColor Gray
Write-Host "   - Click 'New client secret'" -ForegroundColor Gray
Write-Host "   - Copy the secret value (you won't see it again!)" -ForegroundColor Gray

Write-Host "`n4. Update Configuration:" -ForegroundColor Yellow
Write-Host "   - Copy config-template.json to config.json" -ForegroundColor Gray
Write-Host "   - Update with your Tenant ID, Client ID, and Client Secret" -ForegroundColor Gray
Write-Host "   - OR set environment variables:" -ForegroundColor Gray
Write-Host "     • MS_TENANT_ID" -ForegroundColor Gray
Write-Host "     • MS_CLIENT_ID" -ForegroundColor Gray  
Write-Host "     • MS_CLIENT_SECRET" -ForegroundColor Gray

Write-Host "`n5. Test the Setup:" -ForegroundColor Yellow
Write-Host "   .\MicrosoftSupportCasesReport.ps1" -ForegroundColor Gray

Write-Host "`n🔐 SECURITY REMINDER:" -ForegroundColor Red
Write-Host "- Never commit config.json with secrets to version control" -ForegroundColor Yellow
Write-Host "- Use environment variables or Azure Key Vault in production" -ForegroundColor Yellow
Write-Host "- Regularly rotate client secrets" -ForegroundColor Yellow

Write-Host "`n" -NoNewline
Write-Host "=" * 60 -ForegroundColor Cyan
Write-Host "Setup completed! You're ready to generate Microsoft Support reports." -ForegroundColor Green
