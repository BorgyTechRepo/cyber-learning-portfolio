<#
#########################################################
.SYNOPSIS
    This script will pull for all or specific Entra ID users and validate their per user MFA status, Active Directory status, and License status & export the results to a CSV file. 
    The goal being automating this process and sending the report to the helpdesk or security team for review.

.DESCRIPTION
    Author      : Michael Beauregard
    Created     : April 15, 2025
    Version     : 1.0
    Notes       : Microsoft.Graph , Microsoft.Graph.Beta.Users , ActiveDirectory modules must be installed prior to running this script.
                  This version does not prompt for module installation.
                  WARNING: This script accesses sensitive Microsoft Entra / AD data. Only run in environments where you have proper authorization.
    GitHub      : https://github.com/BorgyTechRepo/cyber-learning-portfolio/blob/main/automation/scripts/powershell/Get-MFAReport.ps1
#########################################################
#>

# Import Essential Modules
Write-Host "Importing required modules..." -ForegroundColor Cyan
Import-Module Microsoft.Graph.Beta.Users
Import-Module ActiveDirectory

# Mapping License Type Variable
$skuMap = @{
    "ENTERPRISEPACK" = "Microsoft 365 E3"
    "E5" = "Microsoft 365 E5"
    "POWERAPPS_DEV" = "Power Apps Developer Plan"
    "FLOW_FREE" = "Power Automate Free"
    "Microsoft_365_Copilot" = "Microsoft 365 Copilot"
    "BUSINESS_PREMIUM" = "Microsoft 365 Business Premium"
    "EXCHANGESTANDARD" = "Exchange Online Plan 1"
    "EMS" = "Enterprise Mobility + Security E3"
    "SPE_E3" = "Microsoft 365 E3 (Enterprise Mobility + Security + Windows + Office)"
    "SPE_E5" = "Microsoft 365 E5 (Enterprise Mobility + Security + Windows + Office)"
    "PROJECTPROFESSIONAL" = "Project Plan 3"
    "VISIOONLINE_PLAN2" = "Visio Plan 2"
    "D365_ENTERPRISE_PLAN1" = "Dynamics 365 Enterprise Plan 1"
    "WINDOWS_STORE" = "Windows Store for Business"
    "AAD_PREMIUM" = "Azure Active Directory Premium P1"
    "AAD_PREMIUM_P2" = "Azure Active Directory Premium P2"
    "M365_F1" = "Microsoft 365 F1"
    "M365_F3" = "Microsoft 365 F3"
    "M365_BUSINESS_BASIC" = "Microsoft 365 Business Basic"
    "M365_BUSINESS_STANDARD" = "Microsoft 365 Business Standard"
    "M365_BUSINESS_PREMIUM" = "Microsoft 365 Business Premium"
}

# Connect to Microsoft Graph with required scopes
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "User.Read.All", "Policy.ReadWrite.AuthenticationMethod", "Directory.Read.All", "UserAuthenticationMethod.ReadWrite.All" -NoWelcome
Write-Host "Connected to Microsoft Graph." -ForegroundColor Green

# Ask user for input mode
$mode = Read-Host "Check MFA status for (A)ll users or (S)pecific users? [A/S]"

# Get users based on mode
if ($mode -eq "A") {
    Write-Host "Fetching all users..." -ForegroundColor Cyan
    $users = Get-MgUser -All
    Write-Host "Total users found: $($users.Count)" -ForegroundColor Green
} elseif ($mode -eq "S") {
    $upns = Read-Host "Enter comma-separated User Principal Names"
    $userList = $upns -split ',' | ForEach-Object { $_.Trim() }
    Write-Host "Fetching details for specific users: $($userList -join ', ')" -ForegroundColor Cyan
    $users = foreach ($upn in $userList) {
        try {
            $fetchedUser = Get-MgUser -UserId $upn
            Write-Host "Found user: $upn" -ForegroundColor Green
            $fetchedUser
        } catch {
            Write-Host "User not found: $upn" -ForegroundColor Red
        }
    }
} else {
    Write-Host "Invalid option. Please run the script again and choose A or S." -ForegroundColor Yellow
    return
}

# Prepare a list to store report data
$reportData = @()
$totalUsers = $users.Count
$currentCount = 0

foreach ($user in $users) {
    $currentCount++
    Write-Host "`nProcessing user [$currentCount/$totalUsers]: $($user.UserPrincipalName)" -ForegroundColor Magenta

    try {
        # Get MFA status
        Write-Host "Fetching MFA status..." -ForegroundColor Cyan
        $uri = "/beta/users/$($user.Id)/authentication/requirements"
        $result = Invoke-MgGraphRequest -Method GET -Uri $uri
        $mfaState = $result.perUserMfaState

        # Check AD status
        Write-Host "Checking Active Directory status..." -ForegroundColor Cyan
        try {
            $adUser = Get-ADUser -Filter {UserPrincipalName -eq $user.UserPrincipalName} -Properties Enabled
            if ($adUser) {
                $adStatus = if ($adUser.Enabled) { "Enabled" } else { "Disabled" }
                Write-Host "AD status: $adStatus" -ForegroundColor Green
            } else {
                $adStatus = "Not found in AD"
                Write-Host "AD status: Not found" -ForegroundColor Yellow
            }
        } catch {
            $adStatus = "Error: $_"
            Write-Host "Error checking AD status: $_" -ForegroundColor Red
        }

        # Check license status
        Write-Host "Checking license status..." -ForegroundColor Cyan
        try {
            $licenseDetails = Get-MgUserLicenseDetail -UserId $user.Id
            if ($licenseDetails) {
                $licenseStatus = "Licensed"
                $licenseTypes = ($licenseDetails | ForEach-Object {
                    $skuMap[$_.SkuPartNumber] ?? $_.SkuPartNumber
                }) -join ", "
                Write-Host "License status: Licensed ($licenseTypes)" -ForegroundColor Green
            } else {
                $licenseStatus = "Unlicensed"
                $licenseTypes = "None"
                Write-Host "License status: Unlicensed" -ForegroundColor Yellow
            }
        } catch {
            $licenseStatus = "Error"
            $licenseTypes = "Could not retrieve"
            Write-Host "Error checking license status: $_" -ForegroundColor Red
        }

        # Add to report data
        $reportData += [PSCustomObject]@{
            UserPrincipalName = $user.UserPrincipalName
            MFAState          = $mfaState
            ADStatus          = $adStatus
            LicenseStatus     = $licenseStatus
            LicenseTypes      = $licenseTypes
        }

        Write-Host "✅ Completed user [$currentCount/$totalUsers]" -ForegroundColor Green

    } catch {
        Write-Host "❌ Error processing user: $($user.UserPrincipalName): $_" -ForegroundColor Red
    }
}

# Export report to CSV
Write-Host "`nGenerating report of users' MFA, AD, and license status..." -ForegroundColor Cyan
$ReportPath = [Environment]::GetFolderPath("Desktop")
$ReportFileName = "MfaStatusReport_$((Get-Date -format yyyy-MM-dd_HH-mm-ss).ToString()).csv"
$ReportFullPath = Join-Path -Path $ReportPath -ChildPath $ReportFileName

$reportData | Export-Csv -Path $ReportFullPath -NoTypeInformation -Encoding UTF8
Write-Host "✅ Process completed. Report saved to: $ReportFullPath" -ForegroundColor Green
