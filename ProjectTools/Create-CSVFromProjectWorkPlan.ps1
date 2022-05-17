<#
.SYNOPSIS
    Creates a CSV file from a project work plan

.DESCRIPTION
    This script exports a CSV file from a project Work Plan with phases and tickets (including tasks & notes)

.NOTES
    Company: GreenLoop IT Solutions
    Version 1.0 - Initial release
#>
using namespace System.Runtime.InteropServices
using namespace System.Collections.Generic

# Install the 'ConnectWiseMangeAPI' module if missing
$moduleName = 'ConnectWiseManageAPI'
if (Get-Module -ListAvailable | Where-Object {$_.Name -eq $moduleName}) {
    Import-Module $moduleName
} else {
    Write-Host "'$($moduleName)' not found, installing..."
    Install-Module -Name $moduleName -Force -Scope CurrentUser
    Import-Module $moduleName
}

$manageServerFqdn = Read-Host "Manage Server FQDN"
$connectWiseCompany = Read-Host "ConnectWise Company ID"
$publicAPIKey = Read-Host "PublicAPIKey"
$privateAPIKey = Read-Host "PrivateAPIKey" -AsSecureString

# Confirm the clientID is a valid GUID format (36 characters including hyphens)
do {
    $clientID = Read-Host "ConnectWise Manage Client ID (Go to https://developer.connectwise.com/ if you need to acquire one)"
    $validGuid = [guid]::TryParse($clientID, $([ref][guid]::Empty))
    if (-not $validGuid) {
        Write-Warning "Not a valid GUID format. Try again."
    }
} until ($validGuid)

$connection = @{
    Server     = $manageServerFqdn
    Company    = $connectWiseCompany
    PubKey     = $publicAPIKey
    PrivateKey = [Marshal]::PtrToStringAuto([Marshal]::SecureStringToBSTR($privateAPIKey))
    ClientID   = $clientID
}

# Make the connection to Manage
Connect-CWM @connection

# Run at least once; we will prompt at the end to see whether the user wants to run again.
do {
    # Prompt for the project ID and confirm it exists
    do {
        $parentProjectID = Read-Host "Provide the Project ID of the Manage Project you want to generate a CSV export of"
        $projectFound = Get-CWMProject -id $parentProjectID -ErrorAction SilentlyContinue
        if (-not $projectFound) {
            Write-Warning "Project not found. Try again."
        } else {
            $conditionSearch = "project/id = $parentProjectID"
        }
    } until ($projectFound)

    # Project phases
    $projectPhases = Get-CWMProjectPhase -parentId $parentProjectID
    $projectPhases = foreach ($phase in $projectPhases) {
        [pscustomobject]@{
            WBS     = $phase.wbsCode
            Type    = "Phase"
            Summary = $phase.description
        }
    }

    # Project tickets
    $projectTickets = Get-CWMProjectTicket -condition $conditionSearch -all
    $projectTickets = foreach ($ticket in $projectTickets) {
        [pscustomobject]@{
            WBS           = $ticket.wbsCode
            Type          = "Ticket"
            Summary       = $ticket.summary
            'Time Budget' = $ticket.budgetHours
            TicketID      = $ticket.id
            TicketURL     = "https://$manageServerFqdn/v4_6_release/services/system_io/Service/fv_sr100_request.rails?service_recid=$($ticket.id)&companyName=$connectWiseCompany"
        }
    }

    # Ticket tasks
    $ticketTasks = foreach ($ticket in $projectTickets.TicketID) {
        Get-CWMTicketTask -parentId $ticket
    }
    $ticketTasks = foreach ($task in $ticketTasks) {
        [pscustomobject]@{
            Type     = "Task"
            Summary  = ($task.notes | Out-String).Trim()
            Priority = $task.priority
            TicketID = $task.ticketId
        }
    }

    # Ticket notes
    $ticketNotes = foreach ($ticket in $projectTickets.TicketID) {
        Get-CWMTicketNote -parentId $ticket
    }
    $ticketInternalNotes = $ticketNotes | Where-Object {$_.internalAnalysisFlag -eq $true}
    $ticketInternalNotes = foreach ($intNote in $ticketInternalNotes) {
        [pscustomobject]@{
            Type     = "Internal Note"
            Summary  = ($intNote.text | Out-String).Trim()
            TicketID = $intNote.ticketId
        }
    }

    $ticketDescriptionNotes = $ticketNotes | Where-Object {$_.detailDescriptionFlag -eq $true}
    $ticketDescriptionNotes = foreach ($descNote in $ticketDescriptionNotes) {
        [pscustomobject]@{
            Type     = "Description Note"
            Summary  = ($descNote.text | Out-String).Trim()
            TicketID = $descNote.ticketId
        }
    }

    # Build the Work Plan for export
    # Sorting by WBS code allows us to insert all tasks and notes sequentially based on item type
    [List[object]]$projectWorkPlan = @()
    $combinedObjects = ($projectPhases + $projectTickets) | Sort-Object {[regex]::Replace($_.WBS, '\d+', {$args[0].Value.PadLeft(2)})}

    foreach ($item in $combinedObjects) {
        if ($item.Type -eq 'Phase') {
            $projectWorkPlan.Add($item)
            continue
        }
        if ($item.Type -eq 'Ticket') {
            $projectWorkPlan += $item
            foreach ($task in $ticketTasks) {
                if ($task.TicketID -eq $item.TicketID) {
                    $projectWorkPlan.Add($task)
                }
            }
            foreach ($intNote in $ticketInternalNotes) {
                if ($intNote.TicketID -eq $item.TicketID) {
                    $projectWorkPlan.Add($intNote)
                }
            }
            foreach ($descNote in $ticketDescriptionNotes) {
                if ($descNote.TicketID -eq $item.TicketID) {
                    $projectWorkPlan.Add($descNote)
                }
            }
        }
    }

    # Create the Windows form object
    Add-Type -AssemblyName System.Windows.Forms
    $fileBrowser = New-Object System.Windows.Forms.SaveFileDialog -Property @{
        Title            = "Save file"
        FileName         = "Project Work Plan #$parentProjectID"
        InitialDirectory = "$env:USERPROFILE\Downloads"
        Filter           = 'CSV (*.csv)|*.csv'
    }

    # Prompt to save
    Write-Host ""
    Write-Host "Saving the CSV file...`n"
    $null = $fileBrowser.ShowDialog()
    if ([string]::IsNullOrEmpty($fileBrowser.FileName)) {
        throw "File not saved. Exiting script."
    } else {
        $fileLocation = $fileBrowser.FileName
    }

    $columnHeaders = @("WBS", "Type", "Summary", "Time Budget", "Priority", "TicketURL")
    $projectWorkPlan | Select-Object $columnHeaders | Export-Csv -Path $fileLocation -NoTypeInformation

    Write-Host "Project Work Plan #$($parentProjectID) saved successfully!" -ForegroundColor Green
    Invoke-Item -Path $fileLocation

    $response = Read-Host "Would you like to run again? (Y|N)"
} until ($response -eq 'n')

# Remove sensitive variables
$varsToClear = @('privateAPIKey', 'clientID')
Remove-Variable $varsToClear -ErrorAction SilentlyContinue
