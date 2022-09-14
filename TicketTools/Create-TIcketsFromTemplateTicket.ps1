<#
.SYNOPSIS
    Mass create service tickets in Manage, from a template ticket.
.DESCRIPTION
   Script takes a CSV file of organizations, and a template ticket. All tickets will be created under that associated company.
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

function New-TicketFromTemplate {
    param (
        [Parameter(Mandatory)]
        [array]$templateTicket,
        [array]$CSVline
    )

    $companyNameString = $CSVline.organization -replace "&", "%26" -replace "'", "\'"
    $companyId = (Get-CWMCompany -condition "(name like '$companyNameString*') and (status/name = 'Active') and (deletedFlag = false)").id
    if(!$companyId) {$companyId = 2 }
    $ticketSummary = $templateTicket.summary
    $ticketId = $templateTicket.id

    $ticketBody = @{
        summary         = "$ticketSummary | $companyNameString"[0..99] -Join ""
        recordType      = 'ServiceTicket'
        board           = @{
            id = $templateTicket.board.id
        }
        company         = @{
            id = $companyId
        }
        priority        = @{
            id = $templateTicket.priority.id
        }
        serviceLocation = @{
            id = $templateTicket.serviceLocation.id
        }
        status          = @{
            id = $templateTicket.status.id
        }
        budgetHours     = $templateTicket.budgetHours
        severity        = $templateTicket.severity
        impact          = $templateTicket.impact
    }

    $ticketNotes = Get-CWMTicketNote -ticketId $ticketId
    $ticketTasks = Get-CWMTicketTask -ticketId $ticketId

    $ticket = New-CWMTicket @ticketBody

    # Set task notes
    foreach ($task in $ticketTasks) {
        New-CWMTicketTask -ticketId $($ticket.id) -notes $($task.notes)
    }

    foreach ($note in $ticketnotes) {
        New-CWMTicketNote -parentId $($ticket.id) -text $($note.text) -detailDescriptionFlag $($note.detailDescriptionFlag) -internalAnalysisFlag $($note.internalAnalysisFlag) -resolutionFlag $($note.resolutionFlag) -dateCreated $($note.dateCreated) -createdBy $($note.createdBy)
    }
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

$sourceTicket = Read-Host "Provide the ticket ID of a ticket to serve as a template"

$templateTicket = Get-CWMTicket -ticketId $sourceTicket

# This NEEDS to be a list of ALL firewalls exported from IT Glue
Write-Host ''
Write-Host "Provide a CSV file, with company names in a column for 'organization'."
$CSVfile = Get-CsvFile
if (-not $CSVfile) {
    throw 'File path is empty. Exiting script.'
}

# Process a specific list of firewalls
foreach ($line in $CSVfile) {
    New-TicketFromTemplate -templateTicket $templateTicket -CSVline $line
}

# Remove sensitive variables
$varsToClear = @('privateAPIKey', 'clientID', 'connection')
Remove-Variable $varsToClear -ErrorAction SilentlyContinue
