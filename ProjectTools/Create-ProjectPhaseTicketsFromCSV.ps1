# S. Moody 2/16/2021
# This script creates a set of sequentially numbered project sub-tickets from user contact information provided in a CSV. Tickets under the designated phase given the provided inputs.
# Each ticket will use the provided Displayname and email address and are suitable for creating a custom TimeZest invite for each user.
#currently does NOT populate phone numbers, so make sure to use a TimeZest appt. type that requires user to provide a callback number and populates it into the ticket!

using namespace System.Runtime.InteropServices
Add-Type -AssemblyName System.Windows.Forms

$fileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
    #using the Desktop location as a starting point
    InitialDirectory = [Environment]::GetFolderPath('Desktop')
    Filter = 'CSV (*.csv)|*.csv'
}

#get user input
$userAPIcompany = Read-Host "Provide the Manage company name"
$userAPIpublickey = Read-Host "Provide your public API key for Manage"
$userAPIprivatekey = Read-Host "Provide your private API key for Manage" -AsSecureString

do {
    $clientID = Read-Host "Provide your Manage Client ID"
    $validGuid = [guid]::TryParse($clientID, $([ref][guid]::Empty))
    if (-not $validGuid) {
        Write-Warning "Not a valid client ID (GUID). Try again."
    }
} until ($validGuid)

$manageServerFqdn = Read-Host "Provide the FQDN of your Manage server. Do not use https:// or a trailing slash!"
$ClientNameString = Read-Host "Please Enter the first part of the company name. No wildcard required."
$ProjectNameString = Read-Host "Please Enter the Project Name, exactly as it appears in Manage."
$projectPhase = Read-Host "Enter the 'WBS' project phase (i.e. '2' or '3.2') to use for these tickets. It needs to already exist!"

Write-Host "Browse to and select the CSV file. It should have headers with columns for DisplayName and EmailAddress."
$fileBrowser.Title = "Select file"
$null = $fileBrowser.ShowDialog()
if ([string]::IsNullOrEmpty($fileBrowser.FileName)) {
    throw "File path is empty. Exiting script."
} else {
    $csvFilePath = $fileBrowser.FileName
}

$ticketSummary = Read-Host "Provide a brief description for these tickets, to go in the summary line."
$ticketTemplateID = Read-Host "Enter a project ticket ID to serve as the template"
$manage_base_url = "https://$($manageServerFqdn)/v4_6_release/apis/3.0/"

#set up the Basic auth header
$userAPIprivateKey = [Marshal]::PtrToStringAuto([Marshal]::SecureStringToBSTR($userAPIprivatekey))
$pair = "$($userAPIcompany)+$($userAPIpublickey):$($userAPIprivatekey)"
$encodedCreds = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($pair))
$basicAuthValue = "Basic $encodedCreds"

#load other headers
$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Authorization", $basicAuthValue)
$headers.Add("Content-Type", "application/json")
$headers.Add("clientid", $clientID)

#build functions
function Get-CWProjectTicketNote {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [int]$TicketID
    )

    $endpoint = "project/tickets/$($TicketID)/notes"
    Invoke-RestMethod ($manage_base_url + $endpoint) -Method 'GET' -Headers $headers
}

function Get-CWProjectTicketTask {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [int]$TicketID
    )

    $endpoint = "project/tickets/$($TicketID)/tasks"
    Invoke-RestMethod ($manage_base_url + $endpoint) -Method 'GET' -Headers $headers
}

function Set-CWProjectTicketTask {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param (
        [Parameter(Mandatory)]
        [int]$TicketID,

        [Parameter()]
        [array[]]$TaskNotes,

        [Parameter()]
        [switch]$ClearExisting
    )

    $endpoint = "project/tickets/$($TicketID)/tasks"

    if ($ClearExisting.IsPresent) {
        if ($PSCmdlet.ShouldProcess($TicketID, "Clear existing tasks")) {
            $tasksToDelete = (Get-CWProjectTicketTask -TicketID $TicketID).id
            foreach ($id in $tasksToDelete) {
                Invoke-RestMethod ($manage_base_url + $endpoint + "/$id" ) -Method 'DELETE' -Headers $headers
            }
        }
    }

    $body = @{notes = ""; priority = $null}
    foreach ($n in $TaskNotes) {
        $body.notes = $n.notes
        $body.priority = $n.priority
        Invoke-RestMethod ($manage_base_url + $endpoint) -Method 'POST' -Headers $headers -Body ($body | ConvertTo-Json)
    }
}

function Set-CWProjectTicketNote {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param (
        [Parameter(Mandatory)]
        [int]$TicketID,

        [Parameter()]
        [string[]]$InternalNotes,

        [Parameter()]
        [switch]$ClearExisting
    )

    $endpoint = "project/tickets/$TicketID/notes"

    if ($ClearExisting.IsPresent) {
        if ($PSCmdlet.ShouldProcess($TicketID, "Clear existing notes")) {
            $notesToDelete = (Get-CWProjectTicketNote -TicketID $TicketID).id
            foreach ($id in $notesToDelete) {
                Invoke-RestMethod ($manage_base_url + $endpoint + "/$id" ) -Method 'DELETE' -Headers $headers
            }
        }
    }

    $body = @{
        internalAnalysisFlag = $true
        text                 = ''
    }
    foreach ($note in $InternalNotes) {
        $body.text = $note
        Invoke-RestMethod ($manage_base_url + $endpoint) -Method 'POST' -Headers $headers -Body ($body | ConvertTo-Json)
    }
}

function Get-CWProjectTicket {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [int]$TicketID
    )

    $endpoint = "project/tickets/$($TicketID)"
    Invoke-RestMethod ($manage_base_url + $endpoint) -Method 'GET' -Headers $headers
}

function New-CWProjectTicket {
    [CmdletBinding()]
    param (
        [int]$TicketID,

        [Parameter(Mandatory)]
        [ValidateLength(1,100)]
        [string]$Summary,

        [hashtable]$Project,

        [hashtable]$Phase,

        [ValidateLength(1,62)]
        [string]$ContactName,

        [ValidateLength(1,250)]
        [string]$ContactEmailAddress,

        [double]$BudgetHours,

        [string]$InitialDescription,

        [string]$InitialInternalAnalysis
    )

    $request_url = $manage_base_url + "project/tickets"
    $body = @{
        summary                 = $Summary
        project                 = $Project
        phase                   = $Phase
        contactName             = $username
        contactEmailAddress     = $ContactEmailAddress
        budgetHours             = $BudgetHours
        initialDescription      = $InitialDescription
        initialInternalAnalysis = $InitialInternalAnalysis
    } | ConvertTo-Json

    Invoke-RestMethod $request_url -Method 'POST' -Headers $headers -Body $body
}

#get the description & internal notes from the ticket ID
$ticketNoteInfo = Get-CWProjectTicketNote -TicketID $ticketTemplateID
$descriptionNotes = ($ticketNoteInfo | Where-Object {$_.detailDescriptionFlag -eq $true}).text
$internalNotes = ($ticketNoteInfo | Where-Object {$_.internalAnalysisFlag -eq $true}).text

#get the task notes
$taskNotes = Get-CWProjectTicketTask -TicketID $ticketTemplateID

#get budget hours
$budgetHours = (Get-CWProjectTicket -TicketID $ticketTemplateID).budgetHours

#find the project
$request_url = $manage_base_url + "project/projects?conditions=company/name LIKE `'$ClientNameString*`' AND name = `'$ProjectNameString`'"
$response = Invoke-RestMethod $request_url -Method 'GET' -Headers $headers

#only proceed if we get exactly one result
if ($response.Count -eq 1) {
    $projectId = $response.id

    #get the project phase
    $request_url = $manage_base_url + "project/projects/$($projectId)/phases?conditions=wbsCode = `'$projectPhase`'"
    $response = Invoke-RestMethod $request_url -Method 'GET' -Headers $headers

    if ($response.Count -eq 1) {
        $phaseId = $response.id
        $users = Import-Csv -Path $csvFilePath

        $counter = 1
        #create a ticket for each user
        foreach ($user in $users) {
            $username = $user.DisplayName
            $summary = "$($projectPhase).$($counter) $ticketSummary | $username"

            $projectTicketParams = @{
                Summary                 = $summary
                Project                 = @{id = $projectId}
                Phase                   = @{id = $phaseId}
                ContactName             = $username
                ContactEmailAddress     = $user.EmailAddress
                BudgetHours             = $budgetHours
                InitialDescription      = $descriptionNotes
            }

            $ticketID = (New-CWProjectTicket @projectTicketParams).id
            Set-CWProjectTicketNote -TicketID $ticketID -InternalNotes $internalNotes | Out-Null
            Set-CWProjectTicketTask -TicketID $ticketID -TaskNotes $taskNotes | Out-Null

            Write-Host "Ticket $ticketID created for $username."
            $counter++
        }
    } else {
        Write-Host "We got $($response.Count) results, which is unexpected. Please adjust your Project Phase code and try again."
    }
} else {
    Write-Host "We got $($response.Count) results, which is unexpected. Please adjust your search terms and try again."
}

#remove sensitive variables
$varsToClear = @('userAPIprivatekey', 'pair', 'encodedCreds', 'headers', 'basicAuthValue')
Remove-Variable $varsToClear -ErrorAction SilentlyContinue
