# Authenticate to ConnectWise Manage
using namespace System.Runtime.InteropServices

Install-Module 'ConnectWiseManageAPI'

$connectWiseFQDN = Read-Host "Provide your ConnectWise Manage Server FQDN"
$connectWiseCompany = Read-Host "Provide your ConnectWise Company ID"
$publicKey = Read-Host "API Public Key"
$privateKey = Read-Host "API Private Key" -AsSecureString
$sourceTicket = Read-Host "Provide the ticket ID of a ticket to serve as a template"

# Confirm the $clientID is a valid GUID format (36 characters including hyphens)
do {
    $clientID = Read-Host "Provide your ConnectWise Manage Client ID. Go to https://developer.connectwise.com/ if you need to acquire one."
    $validGuid = [guid]::TryParse($clientID, $([ref][guid]::Empty))
    if (-not $validGuid) {
        Write-Warning "Not a valid GUID format. Try again."
    }
} until ($validGuid)

$connectionParams = @{
    Server     = "$($connectWiseFQDN)/v4_6_release/apis/3.0"
    ClientID   = $clientID
    Company    = $connectWiseCompany
    PubKey     = $publicKey
    PrivateKey = [Marshal]::PtrToStringAuto([Marshal]::SecureStringToBSTR($privateKey))
}
Connect-CWM @connectionParams

$ticketNotes = Get-CWMTicketNote -ticketId $sourceTicket
$taskNotes = (Get-CWMTicketTask -ticketId $sourceTicket).notes

foreach ($ticket in $tickets) {
    # Set task notes
    foreach ($task in $taskNotes) {
        New-CWMTicketTask -ticketId $($ticket.id) -notes $task
    }

    foreach ($note in $ticketnotes) {
        New-CWMTicketNote -parentId $($ticket.id) -text $($note.text) -detailDescriptionFlag $($note.detailDescriptionFlag) -internalAnalysisFlag $($note.internalAnalysisFlag) -resolutionFlag $($note.resolutionFlag) -dateCreated $($note.dateCreated) -createdBy $($note.createdBy)
    }

}
# Disconnect (optional)
#Disconnect-CWM
