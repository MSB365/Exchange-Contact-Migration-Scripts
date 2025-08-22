#region Description
<#     
.NOTES
==============================================================================
Created on:         2025/08/22
Created by:         Drago Petrovic
Organization:       MSB365.blog
Filename:           Export-ExchangeContacts.ps1
Current version:    V1.0     

Find us on:
* Website:         https://www.msb365.blog
* Technet:         https://social.technet.microsoft.com/Profile/MSB365
* LinkedIn:        https://www.linkedin.com/in/drago-petrovic/
* MVP Profile:     https://mvp.microsoft.com/de-de/PublicProfile/5003446
==============================================================================

.SYNOPSIS
    Export all Exchange contacts from on-premise Exchange Server to JSON file
.DESCRIPTION
    This script connects to an on-premise Exchange Server and exports all contacts
    with detailed information to a JSON file for migration purposes.
.PARAMETER ExchangeServer
    The FQDN or IP address of the Exchange Server
.PARAMETER OutputPath
    Path where the JSON file will be saved (default: .\ExchangeContacts.json)
.PARAMETER Credential
    PSCredential object for Exchange authentication (optional - will prompt if not provided)
.EXAMPLE
    .\Export-ExchangeContacts.ps1 -ExchangeServer "exchange.contoso.com" -OutputPath "C:\Migration\contacts.json"

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
    [string]$ExchangeServer,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = ".\ExchangeContacts.json",
    
    [Parameter(Mandatory=$false)]
    [System.Management.Automation.PSCredential]$Credential
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
            default { "White" }
        }
    )
}

try {
    Write-Log "Starting Exchange contact export process..."
    
    # Get credentials if not provided
    if (-not $Credential) {
        Write-Log "Please provide Exchange Server credentials:"
        $Credential = Get-Credential -Message "Enter Exchange Server credentials"
    }
    
    # Import Exchange Management Shell
    Write-Log "Loading Exchange Management Shell..."
    if (Get-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue) {
        Write-Log "Exchange Management Shell already loaded"
    } else {
        Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction Stop
        Write-Log "Exchange Management Shell loaded successfully"
    }
    
    # Connect to Exchange Server
    Write-Log "Connecting to Exchange Server: $ExchangeServer"
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$ExchangeServer/PowerShell/" -Authentication Kerberos -Credential $Credential -ErrorAction Stop
    Import-PSSession $Session -DisableNameChecking -AllowClobber -ErrorAction Stop
    Write-Log "Connected to Exchange Server successfully" -Level "SUCCESS"
    
    # Get all mail contacts
    Write-Log "Retrieving all mail contacts..."
    $Contacts = Get-MailContact -ResultSize Unlimited | Sort-Object DisplayName
    Write-Log "Found $($Contacts.Count) contacts to export"
    
    # Initialize array for contact data
    $ContactData = @()
    $ProcessedCount = 0
    
    foreach ($Contact in $Contacts) {
        try {
            $ProcessedCount++
            Write-Progress -Activity "Exporting Contacts" -Status "Processing $($Contact.DisplayName)" -PercentComplete (($ProcessedCount / $Contacts.Count) * 100)
            
            # Get detailed contact information
            $ContactDetails = Get-MailContact -Identity $Contact.Identity
            $ContactInfo = Get-Contact -Identity $Contact.Identity
            
            # Create contact object with all available properties
            $ContactObject = [PSCustomObject]@{
                # Basic Identity Information
                Identity = $Contact.Identity.ToString()
                DisplayName = $Contact.DisplayName
                Alias = $Contact.Alias
                Name = $Contact.Name
                
                # Email Information
                ExternalEmailAddress = $Contact.ExternalEmailAddress.ToString()
                PrimarySmtpAddress = $Contact.PrimarySmtpAddress.ToString()
                EmailAddresses = @($Contact.EmailAddresses | ForEach-Object { $_.ToString() })
                
                # Personal Information
                FirstName = $ContactInfo.FirstName
                LastName = $ContactInfo.LastName
                Initials = $ContactInfo.Initials
                Title = $ContactInfo.Title
                Department = $ContactInfo.Department
                Company = $ContactInfo.Company
                Office = $ContactInfo.Office
                
                # Phone Numbers
                Phone = $ContactInfo.Phone
                HomePhone = $ContactInfo.HomePhone
                MobilePhone = $ContactInfo.MobilePhone
                Fax = $ContactInfo.Fax
                Pager = $ContactInfo.Pager
                
                # Address Information
                StreetAddress = $ContactInfo.StreetAddress
                City = $ContactInfo.City
                StateOrProvince = $ContactInfo.StateOrProvince
                PostalCode = $ContactInfo.PostalCode
                CountryOrRegion = $ContactInfo.CountryOrRegion
                
                # Additional Information
                Notes = $ContactInfo.Notes
                WebPage = $ContactInfo.WebPage
                Manager = if ($ContactInfo.Manager) { $ContactInfo.Manager.ToString() } else { $null }
                AssistantName = $ContactInfo.AssistantName
                
                # Exchange Specific Properties
                OrganizationalUnit = $Contact.OrganizationalUnit.ToString()
                RecipientType = $Contact.RecipientType.ToString()
                RecipientTypeDetails = $Contact.RecipientTypeDetails.ToString()
                HiddenFromAddressListsEnabled = $Contact.HiddenFromAddressListsEnabled
                RequireSenderAuthenticationEnabled = $Contact.RequireSenderAuthenticationEnabled
                
                # Custom Attributes
                CustomAttribute1 = $Contact.CustomAttribute1
                CustomAttribute2 = $Contact.CustomAttribute2
                CustomAttribute3 = $Contact.CustomAttribute3
                CustomAttribute4 = $Contact.CustomAttribute4
                CustomAttribute5 = $Contact.CustomAttribute5
                CustomAttribute6 = $Contact.CustomAttribute6
                CustomAttribute7 = $Contact.CustomAttribute7
                CustomAttribute8 = $Contact.CustomAttribute8
                CustomAttribute9 = $Contact.CustomAttribute9
                CustomAttribute10 = $Contact.CustomAttribute10
                CustomAttribute11 = $Contact.CustomAttribute11
                CustomAttribute12 = $Contact.CustomAttribute12
                CustomAttribute13 = $Contact.CustomAttribute13
                CustomAttribute14 = $Contact.CustomAttribute14
                CustomAttribute15 = $Contact.CustomAttribute15
                
                # Extension Custom Attributes
                ExtensionCustomAttribute1 = $Contact.ExtensionCustomAttribute1
                ExtensionCustomAttribute2 = $Contact.ExtensionCustomAttribute2
                ExtensionCustomAttribute3 = $Contact.ExtensionCustomAttribute3
                ExtensionCustomAttribute4 = $Contact.ExtensionCustomAttribute4
                ExtensionCustomAttribute5 = $Contact.ExtensionCustomAttribute5
                
                # Timestamps
                WhenCreated = $Contact.WhenCreated
                WhenChanged = $Contact.WhenChanged
                
                # Export metadata
                ExportedOn = Get-Date
                ExportedBy = $env:USERNAME
                SourceServer = $ExchangeServer
            }
            
            $ContactData += $ContactObject
            Write-Log "Exported: $($Contact.DisplayName)" -Level "SUCCESS"
            
        } catch {
            Write-Log "Error processing contact $($Contact.DisplayName): $($_.Exception.Message)" -Level "ERROR"
        }
    }
    
    # Export to JSON
    Write-Log "Exporting $($ContactData.Count) contacts to JSON file: $OutputPath"
    $ContactData | ConvertTo-Json -Depth 10 | Out-File -FilePath $OutputPath -Encoding UTF8
    
    Write-Log "Export completed successfully!" -Level "SUCCESS"
    Write-Log "Total contacts exported: $($ContactData.Count)"
    Write-Log "Output file: $OutputPath"
    Write-Log "File size: $([math]::Round((Get-Item $OutputPath).Length / 1MB, 2)) MB"
    
} catch {
    Write-Log "Critical error during export: $($_.Exception.Message)" -Level "ERROR"
    exit 1
} finally {
    # Clean up session
    if ($Session) {
        Remove-PSSession $Session -ErrorAction SilentlyContinue
        Write-Log "Exchange session closed"
    }
    Write-Progress -Activity "Exporting Contacts" -Completed
}
