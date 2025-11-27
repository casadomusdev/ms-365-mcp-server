<#
.SYNOPSIS
    Checks mailbox delegation permissions for a specific user in Exchange Online

.DESCRIPTION
    This script connects to Exchange Online using an access token and queries
    which shared mailboxes a specific user has Full Access or SendAs permissions for.
    
    The script outputs a JSON array of mailboxes with their permissions.

.PARAMETER UserEmail
    The email address (UPN) of the user to check permissions for

.PARAMETER AccessToken
    The access token for authenticating with Exchange Online

.PARAMETER Organization
    The tenant ID or domain name (e.g., "contoso.onmicrosoft.com")

.EXAMPLE
    .\check-mailbox-permissions.ps1 -UserEmail "user@contoso.com" -AccessToken "eyJ0..." -Organization "contoso.onmicrosoft.com"

.NOTES
    Requires Exchange Online PowerShell Module V3+
    Uses access token authentication (no certificate required)
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$UserEmail,
    
    [Parameter(Mandatory=$true)]
    [string]$CertificatePath,
    
    [Parameter(Mandatory=$true)]
    [string]$CertificatePassword,
    
    [Parameter(Mandatory=$true)]
    [string]$AppId,
    
    [Parameter(Mandatory=$true)]
    [string]$Organization
)

# Set strict mode and error action preference
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Function to write structured error output
function Write-ErrorOutput {
    param([string]$Message, [System.Management.Automation.ErrorRecord]$ErrorRecord)
    
    $errorObj = @{
        error = $true
        message = $Message
        details = if ($ErrorRecord) { $ErrorRecord.Exception.Message } else { $null }
    }
    
    Write-Error ($errorObj | ConvertTo-Json -Compress)
}

# Main script execution
try {
    Write-Verbose "Starting mailbox permission check for user: $UserEmail"
    Write-Verbose "Organization: $Organization"
    
    # Connect to Exchange Online using certificate
    Write-Verbose "Connecting to Exchange Online with certificate authentication..."
    try {
        # Convert certificate password to SecureString
        $secureCertPassword = ConvertTo-SecureString -String $CertificatePassword -AsPlainText -Force
        
        Write-Verbose "Certificate path: $CertificatePath"
        Write-Verbose "App ID: $AppId"
        Write-Verbose "Organization: $Organization"
        
        Connect-ExchangeOnline `
            -CertificateFilePath $CertificatePath `
            -CertificatePassword $secureCertPassword `
            -AppId $AppId `
            -Organization $Organization `
            -ShowBanner:$false `
            -ErrorAction Stop
            
        Write-Verbose "Successfully connected to Exchange Online"
    }
    catch {
        Write-ErrorOutput -Message "Failed to connect to Exchange Online" -ErrorRecord $_
        exit 1
    }
    
    # Initialize results array
    $mailboxResults = @()
    
    # Get all shared mailboxes in the organization
    Write-Verbose "Querying shared mailboxes..."
    try {
        $sharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited -ErrorAction Stop
        Write-Verbose "Found $($sharedMailboxes.Count) shared mailbox(es)"
    }
    catch {
        Write-ErrorOutput -Message "Failed to query shared mailboxes" -ErrorRecord $_
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        exit 1
    }
    
    # Check permissions for each shared mailbox
    foreach ($mailbox in $sharedMailboxes) {
        Write-Verbose "Checking permissions for mailbox: $($mailbox.PrimarySmtpAddress)"
        
        $permissions = @()
        $hasAccess = $false
        
        # Check Full Access permission
        try {
            $fullAccessPerms = Get-MailboxPermission -Identity $mailbox.PrimarySmtpAddress -ErrorAction Stop | 
                Where-Object { 
                    $_.User -like "*$UserEmail*" -and 
                    $_.AccessRights -contains 'FullAccess' -and
                    $_.IsInherited -eq $false
                }
            
            if ($fullAccessPerms) {
                $permissions += 'fullAccess'
                $hasAccess = $true
                Write-Verbose "  - User has Full Access"
            }
        }
        catch {
            Write-Verbose "  - Error checking Full Access: $($_.Exception.Message)"
        }
        
        # Check SendAs permission
        try {
            $sendAsPerms = Get-RecipientPermission -Identity $mailbox.PrimarySmtpAddress -ErrorAction Stop | 
                Where-Object { 
                    $_.Trustee -like "*$UserEmail*" -and 
                    $_.AccessRights -contains 'SendAs'
                }
            
            if ($sendAsPerms) {
                $permissions += 'sendAs'
                $hasAccess = $true
                Write-Verbose "  - User has SendAs"
            }
        }
        catch {
            Write-Verbose "  - Error checking SendAs: $($_.Exception.Message)"
        }
        
        # If user has any access to this mailbox, add it to results
        if ($hasAccess) {
            $mailboxResults += @{
                id = $mailbox.ExternalDirectoryObjectId
                email = $mailbox.PrimarySmtpAddress
                displayName = $mailbox.DisplayName
                permissions = $permissions
            }
            Write-Verbose "  - Added to results with permissions: $($permissions -join ', ')"
        }
    }
    
    # Disconnect from Exchange Online
    Write-Verbose "Disconnecting from Exchange Online..."
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
        Write-Verbose "Disconnected successfully"
    }
    catch {
        # Non-fatal - log but continue
        Write-Verbose "Warning: Failed to disconnect cleanly: $($_.Exception.Message)"
    }
    
    # Output results as JSON
    Write-Verbose "Returning $($mailboxResults.Count) mailbox(es) with permissions"
    $mailboxResults | ConvertTo-Json -Compress -Depth 10
    
    exit 0
}
catch {
    # Catch-all error handler
    Write-ErrorOutput -Message "Unexpected error during script execution" -ErrorRecord $_
    
    # Attempt to disconnect
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    }
    catch {
        # Ignore disconnect errors
    }
    
    exit 1
}
