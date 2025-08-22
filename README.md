# Exchange Contact Migration Scripts

This repository contains PowerShell scripts to migrate contacts from on-premise Exchange Server to Exchange Online (Microsoft 365).

## Overview

The migration process consists of two main scripts:

1. **Export-ExchangeContacts.ps1** - Exports all contacts from on-premise Exchange to a JSON file
2. **Import-ExchangeOnlineContacts.ps1** - Imports contacts from JSON file to Exchange Online

## Prerequisites

### For On-Premise Export Script

- **Exchange Management Shell** installed on the machine running the script
- **PowerShell 5.1** or later
- **Administrative privileges** on the Exchange Server
- **Network connectivity** to the Exchange Server
- **Valid credentials** with Exchange Organization Management or Recipient Management permissions

### For Exchange Online Import Script

- **PowerShell 5.1** or later
- **ExchangeOnlineManagement module** installed
- **Valid Microsoft 365 credentials** with Exchange Administrator or Global Administrator permissions
- **Internet connectivity**

## Installation

### Install Exchange Online Management Module

```powershell
# Install the Exchange Online Management module
Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber

# Import the module
Import-Module ExchangeOnlineManagement
