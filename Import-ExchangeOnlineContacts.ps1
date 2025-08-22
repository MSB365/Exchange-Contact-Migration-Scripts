#region Description
<#     
.NOTES
==============================================================================
Created on:         2025/08/22
Created by:         Drago Petrovic
Organization:       MSB365.blog
Filename:           Import-ExchangeOnlineContacts.ps1
Current version:    V1.0     

Find us on:
* Website:         https://www.msb365.blog
* Technet:         https://social.technet.microsoft.com/Profile/MSB365
* LinkedIn:        https://www.linkedin.com/in/drago-petrovic/
* MVP Profile:     https://mvp.microsoft.com/de-de/PublicProfile/5003446
==============================================================================

.SYNOPSIS
    Import contacts from JSON file to Exchange Online
.DESCRIPTION
    This script imports contacts from a JSON file (created by Export-ExchangeContacts.ps1)
    to Exchange Online. Existing contacts are updated only if there are changes.
.PARAMETER JsonFilePath
    Path to the JSON file containing contact data
.PARAMETER TenantId
    Azure AD Tenant ID for Exchange Online
.PARAMETER Credential
    PSCredential object for Exchange Online authentication (optional - will prompt if not provided)
.PARAMETER WhatIf
    Show what would be done without making changes
.PARAMETER Force
    Skip confirmation prompts
.EXAMPLE
    .\Import-ExchangeOnlineContacts.ps1 -JsonFilePath "C:\Migration\contacts.json" -TenantId "your-tenant-id"

.COPYRIGHT
Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), 
to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, 
and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, 
WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
===========================================================================
.CHANGE LOG
V1.00, 2025/08/22 - DrPe - Initial version



--- keep it simple, but significant ---


--- by MSB365 Blog ---

#>

param(
    [Parameter(Mandatory=$true)]
    [string]$JsonFilePath,
    
    [Parameter(Mandatory=$true)]
    [string]$TenantId,
    
    [Parameter(Mandatory=$false)]
    [System.Management.Automation.PSCredential]$Credential,
    
    [Parameter(Mandatory=$false)]
    [switch]$WhatIf,
    
    [Parameter(Mandatory=$false)]
    [switch]$Force
)

# Function to write log messages
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $(
        switch($Level) {
            "ERROR" { "Red" }
            "WARNING" { "Yellow" }
            "SUCCESS" { "Green" }
            "SKIP" { "Cyan" }
            default { "White" }
        }
    )
}

# Function to compare contact properties
function Compare-ContactProperties {
    param($ExistingContact, $ImportContact)
    
    $differences = @()
    
    # Define properties to compare (excluding read-only and system properties)
    $PropertiesToCompare = @(
        'DisplayName', 'FirstName', 'LastName', 'Title', 'Department', 'Company', 'Office',
        'Phone', 'HomePhone', 'MobilePhone', 'Fax', 'StreetAddress', 'City', 'StateOrProvince',
        'PostalCode', 'CountryOrRegion', 'Notes', 'WebPage', 'AssistantName',
        'CustomAttribute1', 'CustomAttribute2', 'CustomAttribute3', 'CustomAttribute4', 'CustomAttribute5',
        'CustomAttribute6', 'CustomAttribute7', 'CustomAttribute8', 'CustomAttribute9', 'CustomAttribute10',
        'CustomAttribute11', 'CustomAttribute12', 'CustomAttribute13', 'CustomAttribute14', 'CustomAttribute15'
    )
    
    foreach ($Property in $PropertiesToCompare) {
        $ExistingValue = $ExistingContact.$Property
        $ImportValue = $ImportContact.$Property
        
        # Handle null/empty values
        if ([string]::IsNullOrWhiteSpace($ExistingValue)) { $ExistingValue = $null }
        if ([string]::IsNullOrWhiteSpace($ImportValue)) { $ImportValue = $null }
        
        if ($ExistingValue -ne $ImportValue) {
            $differences += [PSCustomObject]@{
                Property = $Property
                ExistingValue = $ExistingValue
                NewValue = $ImportValue
            }
        }
    }
    
    return $differences
}

try {
    Write-Log "Starting Exchange Online contact import process..."
    
    # Validate JSON file
    if (-not (Test-Path $JsonFilePath)) {
        throw "JSON file not found: $JsonFilePath"
    }
    
    # Load contact data from JSON
    Write-Log "Loading contact data from JSON file: $JsonFilePath"
    $ContactData = Get-Content -Path $JsonFilePath -Raw | ConvertFrom-Json
    Write-Log "Loaded $($ContactData.Count) contacts from JSON file"
    
    # Get credentials if not provided
    if (-not $Credential) {
        Write-Log "Please provide Exchange Online credentials:"
        $Credential = Get-Credential -Message "Enter Exchange Online credentials (user@domain.com)"
    }
    
    # Connect to Exchange Online
    Write-Log "Connecting to Exchange Online..."
    try {
        Import-Module ExchangeOnlineManagement -Force -ErrorAction Stop
        Connect-ExchangeOnline -Credential $Credential -ShowProgress $true -ErrorAction Stop
        Write-Log "Connected to Exchange Online successfully" -Level "SUCCESS"
    } catch {
        throw "Failed to connect to Exchange Online: $($_.Exception.Message)"
    }
    
    # Initialize counters
    $CreatedCount = 0
    $UpdatedCount = 0
    $SkippedCount = 0
    $ErrorCount = 0
    $ProcessedCount = 0
    
    # Process each contact
    foreach ($ImportContact in $ContactData) {
        try {
            $ProcessedCount++
            Write-Progress -Activity "Processing Contacts" -Status "Processing $($ImportContact.DisplayName)" -PercentComplete (($ProcessedCount / $ContactData.Count) * 100)
            
            # Check if contact already exists (by external email address)
            $ExistingContact = $null
            try {
                $ExistingContact = Get-MailContact -Identity $ImportContact.ExternalEmailAddress -ErrorAction SilentlyContinue
            } catch {
                # Contact doesn't exist, which is fine
            }
            
            if ($ExistingContact) {
                # Contact exists - check if update is needed
                Write-Log "Contact exists: $($ImportContact.DisplayName)"
                
                # Get detailed existing contact info
                $ExistingContactDetails = Get-Contact -Identity $ExistingContact.Identity
                
                # Compare properties
                $Differences = Compare-ContactProperties -ExistingContact $ExistingContactDetails -ImportContact $ImportContact
                
                if ($Differences.Count -gt 0) {
                    Write-Log "Found $($Differences.Count) differences for $($ImportContact.DisplayName)"
                    
                    if ($WhatIf) {
                        Write-Log "[WHATIF] Would update contact: $($ImportContact.DisplayName)" -Level "WARNING"
                        foreach ($diff in $Differences) {
                            Write-Log "[WHATIF] $($diff.Property): '$($diff.ExistingValue)' -> '$($diff.NewValue)'" -Level "WARNING"
                        }
                    } else {
                        # Update the contact
                        $UpdateParams = @{}
                        
                        # Build update parameters
                        foreach ($diff in $Differences) {
                            if (-not [string]::IsNullOrWhiteSpace($diff.NewValue)) {
                                $UpdateParams[$diff.Property] = $diff.NewValue
                            }
                        }
                        
                        if ($UpdateParams.Count -gt 0) {
                            Set-Contact -Identity $ExistingContact.Identity @UpdateParams -ErrorAction Stop
                            Write-Log "Updated contact: $($ImportContact.DisplayName)" -Level "SUCCESS"
                            $UpdatedCount++
                        }
                    }
                } else {
                    Write-Log "No changes needed for: $($ImportContact.DisplayName)" -Level "SKIP"
                    $SkippedCount++
                }
                
            } else {
                # Contact doesn't exist - create new one
                Write-Log "Creating new contact: $($ImportContact.DisplayName)"
                
                if ($WhatIf) {
                    Write-Log "[WHATIF] Would create contact: $($ImportContact.DisplayName)" -Level "WARNING"
                } else {
                    # Create new mail contact
                    $NewContactParams = @{
                        Name = $ImportContact.DisplayName
                        ExternalEmailAddress = $ImportContact.ExternalEmailAddress
                        DisplayName = $ImportContact.DisplayName
                    }
                    
                    # Add optional parameters if they exist
                    if ($ImportContact.FirstName) { $NewContactParams.FirstName = $ImportContact.FirstName }
                    if ($ImportContact.LastName) { $NewContactParams.LastName = $ImportContact.LastName }
                    if ($ImportContact.Alias) { $NewContactParams.Alias = $ImportContact.Alias }
                    
                    $NewContact = New-MailContact @NewContactParams -ErrorAction Stop
                    
                    # Update additional properties
                    $UpdateParams = @{}
                    
                    # Personal Information
                    if ($ImportContact.Title) { $UpdateParams.Title = $ImportContact.Title }
                    if ($ImportContact.Department) { $UpdateParams.Department = $ImportContact.Department }
                    if ($ImportContact.Company) { $UpdateParams.Company = $ImportContact.Company }
                    if ($ImportContact.Office) { $UpdateParams.Office = $ImportContact.Office }
                    
                    # Phone Numbers
                    if ($ImportContact.Phone) { $UpdateParams.Phone = $ImportContact.Phone }
                    if ($ImportContact.HomePhone) { $UpdateParams.HomePhone = $ImportContact.HomePhone }
                    if ($ImportContact.MobilePhone) { $UpdateParams.MobilePhone = $ImportContact.MobilePhone }
                    if ($ImportContact.Fax) { $UpdateParams.Fax = $ImportContact.Fax }
                    
                    # Address Information
                    if ($ImportContact.StreetAddress) { $UpdateParams.StreetAddress = $ImportContact.StreetAddress }
                    if ($ImportContact.City) { $UpdateParams.City = $ImportContact.City }
                    if ($ImportContact.StateOrProvince) { $UpdateParams.StateOrProvince = $ImportContact.StateOrProvince }
                    if ($ImportContact.PostalCode) { $UpdateParams.PostalCode = $ImportContact.PostalCode }
                    if ($ImportContact.CountryOrRegion) { $UpdateParams.CountryOrRegion = $ImportContact.CountryOrRegion }
                    
                    # Additional Information
                    if ($ImportContact.Notes) { $UpdateParams.Notes = $ImportContact.Notes }
                    if ($ImportContact.WebPage) { $UpdateParams.WebPage = $ImportContact.WebPage }
                    if ($ImportContact.AssistantName) { $UpdateParams.AssistantName = $ImportContact.AssistantName }
                    
                    # Custom Attributes
                    for ($i = 1; $i -le 15; $i++) {
                        $CustomAttr = "CustomAttribute$i"
                        if ($ImportContact.$CustomAttr) {
                            $UpdateParams.$CustomAttr = $ImportContact.$CustomAttr
                        }
                    }
                    
                    # Update contact with additional properties
                    if ($UpdateParams.Count -gt 0) {
                        Set-Contact -Identity $NewContact.Identity @UpdateParams -ErrorAction Stop
                    }
                    
                    Write-Log "Created contact: $($ImportContact.DisplayName)" -Level "SUCCESS"
                    $CreatedCount++
                }
            }
            
        } catch {
            Write-Log "Error processing contact $($ImportContact.DisplayName): $($_.Exception.Message)" -Level "ERROR"
            $ErrorCount++
        }
    }
    
    # Summary
    Write-Log "Import process completed!" -Level "SUCCESS"
    Write-Log "Summary:"
    Write-Log "  Total processed: $ProcessedCount"
    Write-Log "  Created: $CreatedCount"
    Write-Log "  Updated: $UpdatedCount"
    Write-Log "  Skipped (no changes): $SkippedCount"
    Write-Log "  Errors: $ErrorCount"
    
    if ($WhatIf) {
        Write-Log "This was a WhatIf run - no actual changes were made" -Level "WARNING"
    }
    
} catch {
    Write-Log "Critical error during import: $($_.Exception.Message)" -Level "ERROR"
    exit 1
} finally {
    # Disconnect from Exchange Online
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        Write-Log "Disconnected from Exchange Online"
    } catch {
        # Ignore disconnect errors
    }
    Write-Progress -Activity "Processing Contacts" -Completed
}
