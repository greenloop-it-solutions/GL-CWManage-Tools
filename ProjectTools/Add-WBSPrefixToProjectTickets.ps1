#Script: This script finds all tickets for a specified project and updates them so that the WBS Code is prefixed to the Summary.
# Stephen Moody
# GreenLoop IT Solutions
# 2022-02-15 version 1.1
using namespace System.Runtime.InteropServices

$cwCompanyName = Read-Host "ConnectWise Company ID"
$cwAPIPublicKey = Read-Host "Please provide your API public key"
$cwAPIPrivateKey = Read-Host "Please provide your API private key" -AsSecureString

$companyNameMatch = Read-Host "Provide just the FIRST PART of company name"
$projectNameMatch = Read-Host "Provide the EXACT Project Name"

$BasicKey = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes(
    $cwCompanyName + "+" + $cwAPIPublicKey + ":" + [Marshal]::PtrToStringAuto([Marshal]::SecureStringToBSTR($cwAPIPrivateKey))
))

$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Authorization", "Basic $BasicKey")
$headers.Add("Content-Type", "application/json")

# Confirm the clientID is a valid GUID format (36 characters including hyphens)
do {
    $clientID = Read-Host "ConnectWise Manage Client ID (Go to https://developer.connectwise.com/ if you need to acquire one)"
    $validGuid = [guid]::TryParse($clientID, $([ref][guid]::Empty))
    if (-not $validGuid) {
        Write-Warning "Not a valid GUID format. Try again."
    }
} until ($validGuid)
$headers.Add("clientid", $clientID)

$tickets = Invoke-RestMethod "https://connect.greenloopsolutions.com/v4_6_release/apis/3.0/project/tickets?conditions=company/name LIKE '$companyNameMatch*' AND project/name = '$projectNameMatch'" -Method 'GET' -Headers $headers

foreach ($ticket in $tickets) {
    $wbsCode = $ticket.wbsCode
    $summary = $ticket.summary

    #insert regex to match for current WBS Code here
    if ($ticket.summary -match "^[0-9].[0-9]*") {
        if ($matches[0] -eq $wbsCode) {
            continue;
        }
        $newsummary = $summary -replace '^[0-9].[0-9]*', "$wbsCode"
    } else {
        $newsummary = $wbsCode + " " + $summary
    }

    $body = @()
    $body += (@{
        op    = "replace"
        path  = "summary"
        value = $newsummary
    })
    $body += (@{
        op    = "replace"
        path  = "type"
        value = ""
    })

    $body = $body | ConvertTo-Json

    Invoke-RestMethod "https://connect.greenloopsolutions.com/v4_6_release/apis/3.0/project/tickets/$($ticket.id)" -Method Patch -Headers $headers -Body $body
}
