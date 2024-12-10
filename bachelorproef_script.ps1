[CmdletBinding()]
param(
    [Parameter(
        Mandatory=$true,
        HelpMessage="This parameter decides if the GetReportData function should fetch Microsoft Graph reports."
    )]
    [bool]$GetReportData,

    [Parameter(
        Mandatory=$true,
        HelpMessage="This parameter determines if the UploadCSV function should upload the CSV files fetched in GetReportData to a blob storage."
    )]
    [bool]$UploadCSV,

    [Parameter(
        Mandatory=$true,
        HelpMessage="This parameter determines if the NewDB function should create a new database based on the fetched data."
    )]
    [bool]$NewDB,

    [Parameter(
        Mandatory=$true,
        HelpMessage="This parameter determines if the InsertDB function should insert data into the database tables."
    )]
    [bool]$InsertDB=$false,

    [Parameter(
        Mandatory=$true,
        HelpMessage="This parameter determines if the InsertDB function should insert data into the database tables."
    )]
    [string]$StorageAccountName, #bacherlorproefreporting

    [Parameter(
        Mandatory=$true,
        HelpMessage="This parameter determines if the InsertDB function should insert data into the database tables."
    )]
    [string]$ContainerName, #m365

    [Parameter(
        Mandatory=$true,
        HelpMessage="This parameter determines if the InsertDB function should insert data into the database tables."
    )]
    [string]$ResourceGroupName #M365
)


<#
Script Name: Microsoft Graph Reporting Script
Author: Bram Van Hout
Date: [17/4/2024]
Description: This script fetches usage reports from Microsoft Graph API for SharePoint, OneDrive, and Teams, and saves them to CSV files.

Dependencies:
- PowerShell version 7 or later
- Microsoft.Graph module (Install using Install-Module Microsoft.Graph if not already installed)
- An Azure AD application registered with appropriate permissions to access Microsoft Graph API.
  - App ID (Client ID)
  - Tenant ID
  - Certificate Thumbprint for app-only authentication

Usage:
1. Modify the variables $tenantId, $clientId, and $thumbprnt with your Azure AD application details that can be found in config.txt.
2. Run the script in a PowerShell environment.

Output:
- 77 csv files (As of 19/04/2024)

Note: Ensure that the Azure AD application has necessary permissions and access to fetch the reports(Read.Reports.All).
#>
#----------------------------------------------------------------------------------------------------------------------------------
#Importing necessary modules:
Import-Module Az.Accounts
Import-Module Az.Storage
Import-Module Microsoft.Graph.Authentication
Import-Module SqlServer


#----------------------------------------------------------------------------------------------------------------------------------
#Variable declaration:
# Ensures you do not inherit an AzContext in your runbook
#Disable-AzContextAutosave -Scope Process


# Ensures you do not inherit an AzContext in your runbook
Disable-AzContextAutosave -Scope Process

# Connect to Azure with user-assigned managed identity
$AzureContext = (Connect-AzAccount -Identity -AccountId "<account-id>").context

# set and store context
$AzureContext = Set-AzContext -SubscriptionName $AzureContext.Subscription -DefaultProfile $AzureContext


# Debug: Check current Azure context before attempting to connect
Write-Output "Current Azure context before connection:"
Get-AzContext

# Connect to Azure with user-assigned managed identity
#$AzureContext = Connect-AzAccount -Identity -AccountId "<account-id>"

# Debug: Check the context after connecting to ensure it's set
Write-Output "Azure context after connection:"
Get-AzContext

# Check if the subscription is available in the context
if ($AzureContext.Subscription -eq $null) {
    throw "No subscription found in the context. Please ensure the credentials provided are authorized to access an Azure subscription."
}

# Set and store context
#Set-AzContext -SubscriptionName $AzureContext.Subscription -DefaultProfile $AzureContext

# Debug: Verify the context after setting it
#Write-Output "Azure context after setting subscription:"
#Get-AzContext

# Define a known directory for temporary files
$tempDirectory = "C:\temp"

# Create the directory if it doesn't exist
if (-not (Test-Path -Path $TempDirectory)) {
    New-Item -ItemType Directory -Path $TempDirectory
}

# Get a reference to your storage account
$StorageAccount = Get-AzStorageAccount -Name $StorageAccountName -ResourceGroupName $ResourceGroupName

# Get storage account key
$StorageAccountKey = (Get-AzStorageAccountKey -ResourceGroupName $ResourceGroupName -Name $StorageAccountName).Value[0]

# Explicitly declare the type of storage context
[Microsoft.Azure.Commands.Common.Authentication.Abstractions.IStorageContext]$StorageContext = New-AzStorageContext -StorageAccountName $StorageAccountName -StorageAccountKey $StorageAccountKey

# Debug: Verify the storage context
Write-Output "Storage context created:"
$StorageContext

# Specify the folder within the container
$folderName = "Config"  # Replace "your-folder-name" with the desired folder name

# Download endpoints.txt
$endpointBlobName = "Config/endpoints.txt"
Write-Output "Downloading $endpointBlobName to $tempDirectory"
Get-AzStorageBlobContent -Blob $endpointBlobName -Container $ContainerName -Destination $tempDirectory -Context $StorageContext

# Download config.txt
$configBlobName = "Config/config.txt"
Write-Output "Downloading $configBlobName to $tempDirectory"
Get-AzStorageBlobContent -Blob $configBlobName -Container $ContainerName -Destination $tempDirectory -Context $StorageContext


# List files in the temporary directory to confirm download
Write-Output "Files in ${tempDirectory}:"
Get-ChildItem -Path $tempDirectory -Recurse

# Construct the full paths to the endpoint and config files
$endpointPath = Join-Path -Path $tempDirectory -ChildPath "Config/endpoints.txt"
$configPath = Join-Path -Path $tempDirectory -ChildPath "Config/config.txt"

# Check if the files exist before attempting to read them
if (-not (Test-Path -Path $endpointPath)) {
    throw "File not found: $endpointPath"
}

if (-not (Test-Path -Path $configPath)) {
    throw "File not found: $configPath"
}

# Read and convert the JSON content from the files
$arrays = Get-Content -Path $endpointPath | ConvertFrom-Json
$vars = Get-Content -Path $configPath | ConvertFrom-Json

Write-Host $arrays

# Assign the arrays/variables to variables
$endpoints = $arrays.endpoints
$files = $arrays.files
#$ContainerName = $vars.ContainerName
#$DestinationPath = $vars.DestinationPath
#$LocalFolderPath = $vars.LocalFolderPath
#$StorageAccountName = $vars.StorageAccountName
#$ResourceGroupName = $vars.ResourceGroupName
$serverName = $vars.serverName
$databaseName = $vars.databaseName
$username = $vars.username
$VaultName = $vars.VaultName

#Declare the date in the format suitable for SQL
$todaysDate = Get-Date -Format 'yyyy-MM-dd'

# This takes sensitive information stored in the AZ key vault and stores it in these variables
$password = Get-AzKeyVaultSecret -VaultName $VaultName -Name "Database-Password" -AsPlainText
$tenantId = Get-AzKeyVaultSecret -VaultName $VaultName -Name "tenantId" -AsPlainText
$clientId = Get-AzKeyVaultSecret -VaultName $VaultName -Name "clientId" -AsPlainText

# Retrieve the certificate from Azure Key Vault
$cert = Get-AzKeyVaultCertificate -VaultName $VaultName -Name "testtestpleasework"


# Convert the secret value to a byte array
#$certBytes = [System.Convert]::FromBase64String($certSecret.SecretValueText)

# Load the certificate with the private key using the constructor
#$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList @($certBytes, $null, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::Exportable)

# Check if the certificate has a private key
#if (-not $cert.HasPrivateKey) {
#    Write-Error "The certificate does not have a private key."
#    exit
#}

# Disconnect from Azure Account
Disconnect-AzAccount -Scope Process


#----------------------------------------------------------------------------------------------------------------------------------
# Below this script declares all the functions that will be used:

function GetReportData {
    param($endpoint, 
          $outFilePath, 
          $logFilePath, 
          [bool]$getReport = $false)

    $responseHeader = @()
    $retryCount = 5
    $retryDelay = 1 #1 second

    $failedEndpoints = @() # Initialize array to store failed endpoints
    
    for ($i = 1; $i -le $retryCount; $i++){
        try {
            Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0$endpoint" -Method Get -OutputFilePath $outFilePath -ResponseHeadersVariable responseHeader
            $getReport = $true
            return $responseHeader, $getReport
        } catch {
            $failedEndpoints += $endpoint # Add failed endpoint to the array
            Write-Host "Failed to retrieve data from endpoint: $endpoint (Attempt $i)"
            Write-Host $_.Exception.Message
        }
        Start-Sleep -Seconds $retryDelay
    }
    
    # Write failed endpoints to log file
    if ($failedEndpoints.Count -gt 0) {
        $failedEndpoints | Out-File -FilePath $logFilePath -Append
    }
    
    Write-Host "Failed to retrieve data"
    return $null
}


function Upload-CSV {
    param (
        $ContainerName,
        $DestinationPath,
        $StorageAccountName,
        $ResourceGroupName,
        [bool]$uploaded = $false,
	$StorageContext
    )

    # Define the folder name
    $FolderName = "CSVFiles"

    # Get storage account context
    #$StorageAccountKey = (Get-AzStorageAccountKey -ResourceGroupName $ResourceGroupName -Name $StorageAccountName).Value[0]
    #$StorageContext = New-AzStorageContext -StorageAccountName $StorageAccountName -StorageAccountKey $StorageAccountKey

    # Check if the folder exists in the Azure Storage container
    $BlobList = Get-AzStorageBlob -Container $ContainerName -Context $StorageContext -Prefix $FolderName

    if ($BlobList.Count -eq 0) {
        Write-Host "The folder '$FolderName' does not exist in the Azure Storage container '$ContainerName'. Creating it."
    }

    # Upload CSV files if the local folder contains them
    if (Test-Path $DestinationPath -PathType Container) {
        $CSVFiles = Get-ChildItem -Path $DestinationPath -Filter '*.csv'

        if ($CSVFiles.Count -gt 0) {
            # Upload CSV files
            $CSVFiles | ForEach-Object {
                $BlobName = "$FolderName/$($_.Name)"  # Include folder name in blob name
                Set-AzStorageBlobContent -Container $ContainerName -Blob $BlobName -File $_.FullName -Context $StorageContext
                Write-Host "Uploaded $BlobName to $ContainerName"
            }
            $uploaded = $true
            return $uploaded
        } else {
            Write-Host "No CSV files found in $DestinationPath"
        }
    } else {
        Write-Host "The specified folder $DestinationPath does not exist"
    }

    # If the function reaches this point, it means the upload was not successful
    return $false
}

function New-DB {
    param (
        $serverName,
        $databaseName,
        $username,
        $password,
        $DestinationPath,
        [bool]$createdDB = $false
    )

    if (Test-Path $DestinationPath -PathType Container) {
        $csvFiles = Get-ChildItem -Path $DestinationPath -Filter "*.csv"
        
        foreach ($file in $csvFiles) {
            try {
                # Import the CSV file
                $csvData = Import-Csv -Path $file.FullName

                if ($null -eq $csvData -or $csvData.Count -eq 0) {
                    Write-Host "CSV file $($file.FullName) is empty or could not be read."
                    continue  # Skip to the next CSV file
                }

                $tableName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)

                # Check if table exists
                $tableExists = Invoke-SqlCmd -ServerInstance $serverName -Database $databaseName -Username $username -Password $password -Query "IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = '$tableName') SELECT 1 ELSE SELECT 0"

                if ($tableExists -eq 0) {
                    # Construct table creation script based on headers
                    $tableCreationScript = "CREATE TABLE [$tableName] ("
                    foreach ($row in $csvData[0].PSObject.Properties) {
                        $columnName = $row.Name.Replace(" ", "")  # Remove spaces from column name
                        $columnType = "NVARCHAR(MAX)"  # You can customize this based on your CSV data
                        $tableCreationScript += "[$columnName] $columnType,"
                    }
                    # Add a column for the insert date
                    $tableCreationScript += "[InsertDate] Date,"

                    $tableCreationScript = $tableCreationScript.TrimEnd(',') + ");"

                    # Execute table creation script
                    Invoke-SqlCmd -ServerInstance $serverName -Database $databaseName -Username $username -Password $password -Query $tableCreationScript
                } else {
                    Write-Host "Table $tableName already exists."
                }
            } catch {
                Write-Host "Error creating table from $($file.FullName): $_"
            }
        }
    } else {
        Write-Host "The specified folder $DestinationPath does not exist"
    }

    # Set the return value based on whether tables were created
    $createdDB = $true
    return $createdDB
}


function Insert-Data {
    param (
        $serverName,
        $databaseName,
        $username,
        $password,
        $DestinationPath,
        $todaysDate 
    )

    if (Test-Path $DestinationPath -PathType Container) {
        $csvFiles = Get-ChildItem -Path $DestinationPath -Filter "*.csv"
        
        foreach ($file in $csvFiles) {
            $tableName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)

            # Import the CSV file
            $csvData = Import-Csv -Path $file.FullName

            if ($null -eq $csvData -or $csvData.Count -eq 0) {
                Write-Host "CSV file $($file.FullName) is empty or could not be read."
                continue  # Skip to the next CSV file
            }

            # Construct the INSERT INTO query
            $insertQuery = "INSERT INTO [$tableName] ("
            $insertQuery += ($csvData[0].PSObject.Properties | ForEach-Object { "[{0}]" -f $_.Name.Replace(" ", "") }) -join ','
            $insertQuery += ", [InsertDate]) VALUES "

            foreach ($row in $csvData) {
                $values = ($row.PSObject.Properties.Value | ForEach-Object { "'$_'" }) -join ','
                $values += ",'$todaysDate'"
                $insertQuery += "($values),"
            }

            $insertQuery = $insertQuery.TrimEnd(',')

            # Execute the INSERT INTO query
            Invoke-SqlCmd -ServerInstance $serverName -Database $databaseName -Username $username -Password $password -Query $insertQuery
        }
    } else {
        Write-Host "The specified folder $DestinationPath does not exist"
    }
}



#----------------------------------------------------------------------------------------------------------------------------------
#Connection to Microsoft Graph is made:

Connect-MgGraph -ClientID $clientId -TenantId $tenantId -CertificateThumbprint $cert.Thumbprint -Verbose

#Connect-MgGraph -ClientID $clientId -TenantId $tenantId <#-Certificate $cert#> -ClientSecretCredential $credential -Verbose

#----------------------------------------------------------------------------------------------------------------------------------
# Below are all the function calls:

if($GetReportData){
    if($endpoints.Count -eq $files.Count){
        for($i = 0; $i -lt $endpoints.Count;$i++){
            $outFilePath = $tempDirectory + "\" + $files[$i]
            $response, $getReport = GetReportData -endpoint $endpoints[$i] -outFilePath $outFilePath -logFilePath "$tempDirectory/logfilepath.txt"
            Start-Sleep -Seconds "5"
        }
    }
} else {
    $getReport = $true
}

if($UploadCSV){
    if($getReport){
        $uploaded = Upload-CSV -ContainerName $ContainerName -DestinationPath $tempDirectory -StorageAccountName $StorageAccountName -ResourceGroupName $ResourceGroupName -StorageContext $StorageContext
    }
} else {
    $uploaded = $true
}


if($NewDB){
    if($uploaded){
        $createdDB = New-DB -serverName $serverName -databaseName $databaseName -username $username -password $password -DestinationPath $tempDirectory
    }
} else {
    $createdDB = $true
}

if($InsertDB){
    if($createdDB){
        Insert-Data -serverName $serverName -databaseName $databaseName -username $username -password $password -DestinationPath $tempDirectory -todaysDate $todaysDate
    }
}

#----------------------------------------------------------------------------------------------------------------------------------
# This closes the connection to Microsoft Graph
Disconnect-Graph
