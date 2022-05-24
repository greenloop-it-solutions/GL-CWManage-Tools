<#
.SYNOPSIS
    Mass create firewall tickets in Manage
.DESCRIPTION
    The purpose of this script is to mass create tickets to address a potential vulnerability in firmware or current settings
    It ingests a CSV file containing a list of firewalls from IT Glue (and the MySonicWALL portal if needed) to create tickets in Manage

    This avoids creating tickets by hand if you need to assign technicians quickly

    Lastly, the option to include a CSV file exported from MySonicWALL is so that you can target a subset of firewalls by firmware version
    IT Glue does not provide this data so the script will cross reference both files to narrow the scope
    This means that you'll need to have filtered your list by firmware version first if you're targeting a specific number of devices
.NOTES
    Company: GreenLoop IT Solutions
    Version 1.0 - Initial release
#>
using namespace System.Runtime.InteropServices
Add-Type -AssemblyName System.Windows.Forms

# Install the 'ConnectWiseMangeAPI' module if missing
$moduleName = 'ConnectWiseManageAPI'
if (Get-Module -ListAvailable | Where-Object {$_.Name -eq $moduleName}) {
    Import-Module $moduleName
} else {
    Write-Host "'$($moduleName)' not found, installing..."
    Install-Module -Name $moduleName -Force -Scope CurrentUser
    Import-Module $moduleName
}

# Functions
function Get-CsvFile {
    # Create the Windows form object
    $fileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
        InitialDirectory = "$env:USERPROFILE\Downloads"
        Title            = 'Select file'
        Filter           = 'CSV (*.csv)|*.csv'
    }

    $null = $fileBrowser.ShowDialog()
    if ([string]::IsNullOrEmpty($fileBrowser.FileName)) {
        return
    } else {
        $csvFilePath = $fileBrowser.FileName
        $file = Import-Csv -Path $csvFilePath
    }
    return $file
}

function New-FirewallTicket {
    param (
        [Parameter(Mandatory)]
        [array]$Firewall
    )

    $companyNameString = $Firewall.organization -replace "&", "%26" -replace "'", "\'"
    $firewallDesc = "$($Firewall.id) - $($Firewall.name) - $($Firewall.serial_number)"
    $companyId = (Get-CWMCompany -condition "name like '$companyNameString'").id
    $descriptionNote = 'Configure remote access to new GL standards and update documentation.'

    $ticketBody = @{
        summary         = "firewall remote access config | S/N $($Firewall.serial_number)"
        recordType      = 'ServiceTicket'
        board           = @{
            id   = 46
            name = 'Implementation (MS)'
        }
        company         = @{
            id = $companyId
        }
        priority        = @{
            id   = 3
            name = 'Priority 4 - Unclassified'
        }
        serviceLocation = @{
            id   = 2
            name = 'Remote'
        }
        status          = @{
            id   = 1305
            name = 'Backend-NoSLA'
        }
        budgetHours     = 0.50
        severity        = 'Medium'
        impact          = 'Medium'
    }

    $ticketId = (New-CWMTicket @ticketBody -initialDescription $descriptionNote).id

    $internalNote = @"
    This is a low-priority ticket that any tech can take to update the remote management
    of this firewall to the standards documented here: https://greenloop.itglue.com/1286312/docs/6253769

    30 minutes
"@
    New-CWMTicketNote -ticketId $ticketId -text $internalNote -internalAnalysisFlag $true
    New-CWMTicketNote -ticketId $ticketId -text $firewallDesc -detailDescriptionFlag $true

    $task1 = 'Configure WAN management with restrictions'
    $task2 = 'Add TOTP for admin account'
    $task3 = 'Update documentation: admin password TOTP & URL should both be updated'

    New-CWMTicketTask -ticketId $ticketId -notes $task1 -priority 1
    New-CWMTicketTask -ticketId $ticketId -notes $task2 -priority 2
    New-CWMTicketTask -ticketId $ticketId -notes $task3 -priority 3
}

# Manage server info
$manageServerFqdn = Read-Host 'Manage Server FQDN'
$connectWiseCompany = Read-Host 'ConnectWise Company ID'
$publicAPIKey = Read-Host 'PublicAPIKey'
$privateAPIKey = Read-Host 'PrivateAPIKey' -AsSecureString

# Confirm the $clientID is a valid GUID format (36 characters including hyphens)
do {
    $clientID = Read-Host 'ConnectWise Manage Client ID (Go to https://developer.connectwise.com/ if you need to acquire one)'
    $validGuid = [guid]::TryParse($clientID, $([ref][guid]::Empty))
    if (-not $validGuid) {
        Write-Warning 'Not a valid GUID format. Try again.'
    }
} until ($validGuid)

$connection = @{
    Server     = $manageServerFqdn
    Company    = $connectWiseCompany
    PubKey     = $publicAPIKey
    PrivateKey = [Marshal]::PtrToStringAuto([Marshal]::SecureStringToBSTR($privateAPIKey))
    ClientID   = $clientID
}

Connect-CWM @connection

# Here we are looking for a single file OR two files if we're going to compare and perform logic on a subset of firewalls
# Default value is $false - meaning, we start with one file only
$firewallSubset = $false

# This NEEDS to be a list of ALL firewalls exported from IT Glue
Write-Host ''
Write-Host 'Provide a CSV file exported from IT Glue...'
$file1 = Get-CsvFile
if (-not $file1) {
    throw 'File path is empty. Exiting script.'
}

Write-Host ''
$additionalFile = Read-Host 'Provide a CSV file exported from the MySonicWALL portal? (y/n)'
if ($additionalFile -eq 'y') {
    # This NEEDS to be a list of firewalls exported from the MySonicWall portal (this should be a specific list of devices)
    # This is required to narrow the scope of select firmware versions - IT Glue does not provide these values
    Write-Host ''
    $file2 = Get-CsvFile
    if (-not $file2) {
        throw 'File path is empty. Exiting script.'
    }
    $firewallSubset = $true
}

$firewallsProcessed = 0
if ($firewallSubset) {
    # Process a specific list of firewalls
    foreach ($serialNo in $file2) {
        foreach ($firewall in $file1) {
            if ($firewall.serial_number -contains $serialNo.'Serial no.') {
                New-FirewallTicket -Firewall $firewall
                $firewallsProcessed++
            }
        }
    }
} else {
    # Process ALL firewalls in the list
    foreach ($firewall in $file1) {
        New-FirewallTicket -Firewall $firewalls
        $firewallsProcessed++
    }
}

Write-Host ''
Write-Host "Total firewalls processed: $firewallsProcessed"

# Remove sensitive variables
$varsToClear = @('privateAPIKey', 'clientID', 'connection')
Remove-Variable $varsToClear -ErrorAction SilentlyContinue
