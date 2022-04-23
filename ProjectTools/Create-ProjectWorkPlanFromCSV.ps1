<#
.SYNOPSIS
    Creates a Manage project from a CSV file

.DESCRIPTION
    This script ingests a CSV file and then fills in a project Work Plan with phases and tickets (including tasks & notes)

.NOTES
    Company: GreenLoop IT Solutions
    Version 1.0 - Initial release
#>

using namespace System.Runtime.InteropServices

$parentProjectID = Read-Host "Provide the Project ID you want to populate the work plan for. This project needs to already exist!"
$companyName = Read-Host "Manage Company Name"
$manageServerFqdn = (Read-Host "Manage Server FQDN") -replace 'https://|http://|/'
$manageBaseUrl = "https://$($manageServerFqdn)/v4_6_release/apis/3.0/"

# Confirm the $clientID is a valid GUID format (36 characters including hyphens)
do {
    $clientID = Read-Host "Provide your ConnectWise Manage Client ID. Go to https://developer.connectwise.com/ if you need to acquire one."
    $validGuid = [guid]::TryParse($clientID, $([ref][guid]::Empty))
    if (-not $validGuid) {
        Write-Warning "Not a valid GUID format. Try again."
    }
} until ($validGuid)

# Create the basic auth value and headers
$publicAPIKey = Read-Host "PublicAPIKey"
$privateAPIKey = Read-Host "PrivateAPIKey" -AsSecureString
$clearPrivateAPIKey= [Marshal]::PtrToStringAuto([Marshal]::SecureStringToBSTR($privateAPIKey))
$pair = "$($companyName)+$($publicAPIKey):$($clearPrivateAPIKey)"
$encodedCreds = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($pair))
$basicAuthValue = "Basic $encodedCreds"

$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Authorization", $basicAuthValue)
$headers.Add("Content-Type", "application/json")
$headers.Add("clientid", $clientID)

# Build functions
function Get-CWProjectTicketPhase {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [int]$TicketID
    )

    $request_url = $manageBaseUrl + "project/projects/$TicketID/phases"
    Invoke-RestMethod $request_url -Method 'GET' -Headers $headers
}
function Get-CWProjectTicket {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [int]$TicketID
    )

    $request_url = $manageBaseUrl + "project/tickets/$TicketID"
    Invoke-RestMethod $request_url -Method 'GET' -Headers $headers
}
function Get-CWProjectTicketNote {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [int]$TicketID
    )

    $request_url = $manageBaseUrl + "project/tickets/$TicketID/notes"
    Invoke-RestMethod $request_url -Method 'GET' -Headers $headers
}
function Get-CWProjectTicketTask {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [int]$TicketID
    )

    $request_url = $manageBaseUrl + "project/tickets/$TicketID/tasks"
    Invoke-RestMethod $request_url -Method 'GET' -Headers $headers
}
function Set-CWProjectTicketTask {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param (
        [Parameter(Mandatory)]
        [int]$TicketID,

        [Parameter()]
        [string]$TaskNotes,

        [Parameter()]
        [int]$Priority,

        [Parameter()]
        [switch]$ClearExisting
    )

    $request_url = $manageBaseUrl + "project/tickets/$TicketID/tasks"

    if ($ClearExisting.IsPresent) {
        if ($PSCmdlet.ShouldProcess($TicketID, "Clear existing tasks")) {
            $tasksToDelete = (Get-CWProjectTicketTask -TicketID $TicketID).id
            foreach ($id in $tasksToDelete) {
                Invoke-RestMethod ($request_url + "/$id" ) -Method 'DELETE' -Headers $headers
            }
        }
    }

    $body = @{
        notes    = $TaskNotes
        priority = $Priority
    } | ConvertTo-Json

    Invoke-RestMethod $request_url -Method 'POST' -Headers $headers -Body $body
}
function Set-CWProjectTicketNote {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param (
        [Parameter(Mandatory)]
        [int]$TicketID,

        [Parameter(ParameterSetName='Internal')]
        [string]$InternalNotes,

        [Parameter(ParameterSetName='Description')]
        [string]$DescriptionNotes,

        [Parameter()]
        [switch]$ClearExisting
    )

    $request_url = $manageBaseUrl + "project/tickets/$TicketID/notes"

    if ($ClearExisting.IsPresent) {
        if ($PSCmdlet.ShouldProcess($TicketID, "Clear existing notes")) {
            $notesToDelete = (Get-CWProjectTicketNote -TicketID $TicketID).id
            foreach ($id in $notesToDelete) {
                Invoke-RestMethod ($request_url + "/$id" ) -Method 'DELETE' -Headers $headers
            }
        }
    }

    if ($InternalNotes) {
        $body = @{
            internalAnalysisFlag = $true
            text                 = $InternalNotes
        } | ConvertTo-Json
    }

    if ($DescriptionNotes) {
        $body = @{
            detailDescriptionFlag = $true
            text                  = $DescriptionNotes
        } | ConvertTo-Json
    }

    Invoke-RestMethod $request_url -Method 'POST' -Headers $headers -Body $body
}
function New-CWProjectTicketPhase {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [int]$TicketID,

        [Parameter(Mandatory)]
        [ValidateLength(1, 100)]
        [string]$Description,

        [int]$ParentPhaseID
    )

    $request_url = $manageBaseUrl + "project/projects/$TicketID/phases"

    $body = @{
        description = $Description
        parentPhase = @{id = $ParentPhaseID}
    } | ConvertTo-Json

    Invoke-RestMethod $request_url -Method 'POST' -Headers $headers -Body $body
}
function New-CWProjectTicket {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
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

    $request_url = $manageBaseUrl + "project/tickets"
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

# Create the Windows form object
Add-Type -AssemblyName System.Windows.Forms
$fileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
    #using the Desktop location as a starting point
    InitialDirectory = [Environment]::GetFolderPath('Desktop')
    Filter = 'CSV (*.csv)|*.csv'
}

# Prompt for the CSV file
Write-Host ""
Write-Host "Browse to and select the CSV file to build the project from...`n"
$fileBrowser.Title = "Select file"
$null = $fileBrowser.ShowDialog()
if ([string]::IsNullOrEmpty($fileBrowser.FileName)) {
    throw "File path is empty. Exiting script."
} else {
    $csvFilePath = $fileBrowser.FileName
    $project = Import-Csv -Path $csvFilePath
}

# Create the project work plan
foreach ($item in $project) {
    switch ($item.Type) {
        'Phase' {
            Write-Host "Creating Phase $($item.WBS) - $($item.Summary)`n" -ForegroundColor Yellow
            if ($item.WBS.Length -eq 3) {
                # found a sub-phase
                $phaseID = ((Get-CWProjectTicketPhase -TicketID $parentProjectID).GetEnumerator() | Where-Object {$_.wbsCode -eq $item.WBS[0]}).id
                New-CWProjectTicketPhase -TicketID $parentProjectID -Description $item.Summary -ParentPhaseID $phaseID
            } else {
                New-CWProjectTicketPhase -TicketID $parentProjectID -Description $item.Summary
            }
        }
        'Ticket' {
            Write-Host "Creating Ticket $($item.WBS) - $($item.Summary) - $($item.'Time Budget')`n" -ForegroundColor Yellow
            if ($item.WBS.Length -eq 5) {
                # found a ticket belonging to a sub-phase
                $phaseID = ((Get-CWProjectTicketPhase -TicketID $parentProjectID).GetEnumerator() | Where-Object {$_.wbsCode -eq $item.WBS.Substring(0, 3)}).id
            } else {
                $phaseID = ((Get-CWProjectTicketPhase -TicketID $parentProjectID).GetEnumerator() | Where-Object {$_.wbsCode -eq $item.WBS[0]}).id
            }
            $recentTicketID = (New-CWProjectTicket -TicketID $parentProjectID -Summary $item.Summary -BudgetHours $item.'Time Budget' -Phase @{id = $phaseID}).id
        }
        'Task' {
            Set-CWProjectTicketTask -TicketID $recentTicketID -TaskNotes $item.Summary -Priority $item.Priority
        }
        'Internal Note' {
            Set-CWProjectTicketNote -TicketID $recentTicketID -InternalNotes $item.Summary
        }
        'Description Note' {
            Set-CWProjectTicketNote -TicketID $recentTicketID -DescriptionNotes $item.Summary
        }
    }
}

# Remove sensitive variables
$varsToClear = @('privateAPIKey', 'clearPrivateAPIkey', 'pair', 'encodedCreds', 'headers', 'basicAuthValue', 'clientID')
Remove-Variable $varsToClear -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "Project created successfully! (#$($parentProjectID))" -ForegroundColor Green
