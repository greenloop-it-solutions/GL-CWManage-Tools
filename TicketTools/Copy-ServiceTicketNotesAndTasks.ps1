# Authenticate to ConnectWise Manage
using namespace System.Runtime.InteropServices

Install-Module 'ConnectWiseManageAPI'

$connectWiseFQDN = Read-Host "Provide your ConnectWise Manage Server FQDN"
$connectWiseCompany = Read-Host "Provide your ConnectWise Company ID"
$publicKey = Read-Host "API Public Key"
$privateKey = Read-Host "API Private Key" -AsSecureString
$sourceTicket = Read-Host "Provide the ticket ID of a ticket to serve as a template. This assumes the variable '$tickets' will already be populated with target tickets."
#$targetTicket = Read-Host "Target Ticket ID"
$connectionParams = @{
Server = "$($connectWiseFQDN)/v4_6_release/apis/3.0"
ClientID = '2628d75d-c2ea-43c4-bac4-33ea7181d0d6'
Company = $connectWiseCompany
PubKey = $publicKey
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
       new-cwmticketnote -parentId $($ticket.id) -text $($note.text) -detailDescriptionFlag $($note.detailDescriptionFlag) -internalAnalysisFlag $($note.internalAnalysisFlag) -resolutionFlag $($note.resolutionFlag) -dateCreated $($note.dateCreated) -createdBy $($note.createdBy)
    }
   
}
# Disconnect (optional)
#Disconnect-CWM
