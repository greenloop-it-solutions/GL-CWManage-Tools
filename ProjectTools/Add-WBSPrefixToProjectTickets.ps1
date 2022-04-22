#Script: This script finds all tickets for a specified project and updates them so that the WBS Code is prefixed to the Summary.
# Stephen Moody
# GreenLoop IT Solutions
# 2022-02-15 version 1.0

$companyNameMatch = Read-Host "Provide just the FIRST PART of company name"
$projectNameMatch = Read-Host "Provide the EXACT Project Name"

$cwCompanyName = "greenloop"
$cwAPIPublicKey = Read-Host "Please provide your API public key"
$cwAPIPrivateKey = Read-Host "Please provide your API private key"

$BasicKey = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($cwCompanyName + "+" + $cwAPIPublicKey + ":" + $cwAPIPrivateKey))

$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Authorization", "Basic $BasicKey")
$headers.Add("Content-Type", "application/json")
$headers.Add("clientid", "2628d75d-c2ea-43c4-bac4-33ea7181d0d6")

$tickets = Invoke-RestMethod "https://connect.greenloopsolutions.com/v4_6_release/apis/3.0/project/tickets?conditions=company/name LIKE '$companyNameMatch*' AND project/name = '$projectNameMatch'" -Method 'GET' -Headers $headers

foreach ($ticket in $tickets) {
    $wbsCode = $ticket.wbsCode
    $summary = $ticket.summary
    
    #insert regex to match for current WBS Code here
    if ($ticket.summary -match "^[0-9].[0-9]*") {
            if ($matches[0] -eq $wbsCode) {
            continue;
        }
        $newsummary = $summary -replace '^[0-9].[0-9]*',"$wbsCode"
    } else {
        $newsummary = $wbsCode + " " + $summary
    }


    $body =  @{
        op = "replace"
        path = "summary"
        value = $newsummary
    } |ConvertTo-Json

    $body ="[$body]"

    Invoke-RestMethod "https://connect.greenloopsolutions.com/v4_6_release/apis/3.0/project/tickets/$($ticket.id)" -Method Patch -Headers $headers -Body $body
}
