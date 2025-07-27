# Micrsoft Graph Reporting

### Description: 
This script fetches usage reports from Microsoft Graph API for SharePoint, OneDrive, and Teams, and saves them to CSV files.
This was part of my internship and was my bachelor project.

### Dependencies:
- PowerShell version 7 or later
- Microsoft.Graph module (Install using Install-Module Microsoft.Graph if not already installed)
- An Azure AD application registered with appropriate permissions to access Microsoft Graph API.
  - App ID (Client ID)
  - Tenant ID
  - Certificate Thumbprint for app-only authentication

### Usage:
1. Modify the variables $tenantId, $clientId, and $thumbprnt with your Azure AD application details that can be found in config.txt.
2. Run the script in a PowerShell environment.
