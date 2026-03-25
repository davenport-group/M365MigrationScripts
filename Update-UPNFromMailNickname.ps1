<#
.SYNOPSIS
    Updates user UPNs in Entra ID using Mail Nickname attribute from Microsoft Graph.

.DESCRIPTION
    This script connects to a Microsoft 365 tenant using Microsoft Graph PowerShell SDK and:
    1. Retrieves all members (excluding guests and external users)
    2. For each user, gets the Mail Nickname attribute from Entra ID
    3. Updates the User Principal Name to use the Mail Nickname with the specified domain
    4. Provides WhatIf preview mode to review changes before applying

.PARAMETER TenantId
    The Azure AD Tenant ID. If not specified, uses interactive authentication.

.PARAMETER TargetDomain
    The domain suffix to append to the Mail Nickname (e.g., 'contoso.com', 'contoso.onmicrosoft.com').
    If not specified, will use the tenant's default domain.

.PARAMETER ExcludedUPNs
    An array of UPNs to exclude from the update (e.g., service accounts). These users will be skipped.

.NOTES
    Requires: Microsoft.Graph.Users, Microsoft.Graph.Identity.DirectoryManagement modules
    Author: UPN Migration Script
    Version: 1.1
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $false)]
    [string]$TenantId,
    
    [Parameter(Mandatory = $false)]
    [string]$TargetDomain,
    
    [Parameter(Mandatory = $false)]
    [string[]]$ExcludedUPNs
)

# Function to ensure required modules are installed
function Ensure-GraphModules {
    $requiredModules = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Users'
    )
    
    foreach ($module in $requiredModules) {
        if (-not (Get-Module -ListAvailable -Name $module)) {
            Write-Host "Installing module: $module" -ForegroundColor Cyan
            Install-Module -Name $module -Repository PSGallery -Force -AllowClobber
        }
    }
    
    Write-Host "Importing required modules..." -ForegroundColor Green
    Import-Module Microsoft.Graph.Authentication -Force
    Import-Module Microsoft.Graph.Users -Force
}

# Function to get tenant's default domain
function Get-TenantDefaultDomain {
    try {
        $organization = Get-MgOrganization -ErrorAction Stop
        $domains = $organization.VerifiedDomains | Where-Object { $_.IsDefault -eq $true }
        
        if ($domains) {
            return $domains[0].Name
        }
        else {
            # Fallback to first verified domain
            $allDomains = $organization.VerifiedDomains | Select-Object -First 1
            return $allDomains.Name
        }
    }
    catch {
        Write-Error "Failed to retrieve tenant default domain: $_"
        return $null
    }
}

# Function to get all members excluding guests and external users
function Get-TenantMembers {
    try {
        $members = @()
        $pageSize = 999
        
        Write-Host "Retrieving all members (excluding guests and external users)..." -ForegroundColor Green
        
        # Filter to exclude guests and external users
        $filter = "userType eq 'Member' and accountEnabled eq true"
        
        # Get users with pagination
        $users = Get-MgUser -Filter $filter -All -PageSize $pageSize -Property @(
            'id',
            'userPrincipalName',
            'mailNickname',
            'displayName',
            'mail',
            'userType',
            'accountEnabled'
        ) -ErrorAction Stop
        
        if ($users) {
            $members = @($users)
            Write-Host "Found $($members.Count) members" -ForegroundColor Green
        }
        else {
            Write-Host "No members found matching filter" -ForegroundColor Yellow
        }
        
        return $members
    }
    catch {
        Write-Error "Failed to retrieve members: $_"
        return @()
    }
}

# Function to update user UPN
function Update-UserUPN {
    param(
        [Parameter(Mandatory = $true)]
        $User,
        
        [Parameter(Mandatory = $true)]
        [string]$NewUPN,
        
        [bool]$WhatIfMode = $true
    )
    
    try {
        if ($NewUPN -eq $User.UserPrincipalName) {
            Write-Host "  ℹ️  SKIP: UPN already matches ($NewUPN)" -ForegroundColor Gray
            return $false
        }
        
        if ($WhatIfMode) {
            Write-Host "  📋 WHAT IF: Would update UPN" -ForegroundColor Cyan
            Write-Host "      From: $($User.UserPrincipalName)" -ForegroundColor Gray
            Write-Host "      To:   $NewUPN" -ForegroundColor Gray
            return $true
        }
        else {
            Write-Host "  🔄 Updating UPN..." -ForegroundColor Yellow
            Update-MgUser -UserId $User.Id -UserPrincipalName $NewUPN -ErrorAction Stop
            Write-Host "  ✅ UPN updated successfully" -ForegroundColor Green
            return $true
        }
    }
    catch {
        Write-Host "  ❌ ERROR: Failed to update UPN: $_" -ForegroundColor Red
        return $false
    }
}

# Main script logic
function Main {
    try {
        Write-Host "╔════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
        Write-Host "║  Microsoft Graph UPN Update from Mail Nickname Script      ║" -ForegroundColor Cyan
        Write-Host "╚════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
        Write-Host ""
        
        # Ensure modules are installed
        Ensure-GraphModules
        Write-Host ""
        
        # Connect to Microsoft Graph
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Green
        
        $connectParams = @{
            Scopes = @('User.ReadWrite.All', 'Directory.ReadWrite.All')
        }
        
        if ($TenantId) {
            $connectParams['TenantId'] = $TenantId
        }
        
        Connect-MgGraph @connectParams -NoWelcome -ErrorAction Stop
        Write-Host "✅ Connected to Microsoft Graph" -ForegroundColor Green
        Write-Host ""
        
        # Get or determine target domain
        if (-not $TargetDomain) {
            Write-Host "Determining target domain..." -ForegroundColor Green
            $TargetDomain = Get-TenantDefaultDomain
            
            if (-not $TargetDomain) {
                Write-Error "Could not determine target domain. Please specify -TargetDomain parameter."
                return
            }
        }
        
        Write-Host "Target domain: $TargetDomain" -ForegroundColor Cyan
        Write-Host ""
        
        # Get all members
        $members = Get-TenantMembers
        
        if ($members.Count -eq 0) {
            Write-Host "No members found to process." -ForegroundColor Yellow
            return
        }
        
        Write-Host ""
        Write-Host "Processing members..." -ForegroundColor Green
        Write-Host ""
        
        if ($WhatIfPreference) {
            Write-Host "⚠️  WHAT IF MODE ENABLED - No changes will be made" -ForegroundColor Yellow
            Write-Host ""
        }
        
        $successCount = 0
        $skipCount = 0
        $errorCount = 0
        
        foreach ($user in $members) {
            # Check if user is in excluded list
            if ($ExcludedUPNs -contains $user.UserPrincipalName) {
                Write-Host "⊘ EXCLUDED: User '$($user.DisplayName)' ($($user.UserPrincipalName)) is in the exclusion list" -ForegroundColor Magenta
                $skipCount++
                continue
            }
            
            if ([string]::IsNullOrWhiteSpace($user.MailNickname)) {
                Write-Host "❌ SKIP: User '$($user.DisplayName)' has no Mail Nickname" -ForegroundColor Yellow
                $skipCount++
                continue
            }
            
            $newUPN = "$($user.MailNickname)@$TargetDomain"
            
            Write-Host "User: $($user.DisplayName)" -ForegroundColor White
            
            if (Update-UserUPN -User $user -NewUPN $newUPN -WhatIfMode $WhatIfPreference) {
                $successCount++
            }
            else {
                $errorCount++
            }
            
            Write-Host ""
        }
        
        # Summary
        Write-Host "╔════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
        Write-Host "║  SUMMARY                                                   ║" -ForegroundColor Cyan
        Write-Host "╚════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
        Write-Host "Total members processed: $($members.Count)"
        Write-Host "✅ Updated/Processed: $successCount"
        Write-Host "⊘ Skipped: $skipCount"
        Write-Host "❌ Errors: $errorCount"
        Write-Host ""
        
        if ($WhatIfPreference) {
            Write-Host "To apply these changes, run the script again with -WhatIf -WhatIfPreference:`$false" -ForegroundColor Yellow
        }
        
        # Disconnect
        Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Green
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Write-Host "✅ Disconnected" -ForegroundColor Green
    }
    catch {
        Write-Error "Script error: $_"
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        exit 1
    }
}

# Run main script
Main
