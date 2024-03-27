$baseUri = "https://connect.greenloopsolutions.com/v4_6_release/apis/3.0/"

$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Authorization", "Basic Z3JlZW5sb29wKzFWaGVpMFZCdEdKMDNxNjA6dnFmTmNQa0tVSW5tbjNWdQ==")
$headers.Add("clientId", "bb74caf5-89f4-4d2a-ba3e-5bb003843b13")
$headers.Add("Content-Type", "application/json")

$agrBody1 = @(
    @{
    op ="replace"
    path = "prorateFlag"
    value = $true
    }
)
$agrBody1 = ConvertTo-Json -InputObject $agrBody1

$agrBody2 = @(

    @{
    op ="replace"
    path = "prorateFlag"
    value = $false
    }
)
$agrBody2 = ConvertTo-Json -InputObject $agrBody2


<#
IV_Item_RecID	Product ID	Description	ColumnName
2319	INFO-MSA-Antivirus	Antivirus	INFO-MSA-Antivirus
2342	INFO-MSA-BYOD-Support-60	Support for Non-Client Owned Equipment 	INFO-MSA-BYOD-Support-60
2345	INFO-MSA-BYOD-Support-CONTACTFIRST	Support for Non-Client Owned Equipment 	INFO-MSA-BYOD-Support-CONTACTFIRST
2334	INFO-MSA-BYOD-Support-NONE	Support for Non-Client Owned Equipment 	INFO-MSA-BYOD-Support-NONE
2343	INFO-MSA-CoverageArea-Bend	Coverage Area Bend	INFO-MSA-CoverageArea-Bend
2333	INFO-MSA-CoverageArea-Phoenix	Coverage Area Phoenix	INFO-MSA-CoverageArea-Phoenix
2344	INFO-MSA-CoverageArea-Portland	Coverage Area Portland	INFO-MSA-CoverageArea-Portland
2327	INFO-MSA-Darkweb-Monitoring	Darkweb monitoring	INFO-MSA-Darkweb-Monitoring
2328	INFO-MSA-E-mail-Hygiene	E-mail Hygiene and Securty	INFO-MSA-E-mail-Hygiene
2320	INFO-MSA-Endpoint-Security	Endpoint-Security	INFO-MSA-Endpoint-Security
2322	INFO-MSA-Machine-Security-EnhancedPolicy	Machine Security - Enhanced Policy and Monitoring	INFO-MSA-Machine-Security-EnhancedPolicy
2318	INFO-MSA-MFA-PolicyEnforcement	MFA	INFO-MSA-MFA-PolicyEnforcement
2317	INFO-MSA-MS-O365Management	MS O365 management	INFO-MSA-MS-O365Management
2430	INFO-MSA-No-Endpoint-Security	No-Endpoint-Security	INFO-MSA-Endpoint-Security
2300	INFO-MSA-OnCall-247Included	On-call support w/24x7 Emergency Support (included)	INFO-MSA-OnCall-247Included
2299	INFO-MSA-OnCall-ExtdHrs	On-call support (Extended-hours, billable) 	INFO-MSA-OnCall-ExtdHrs
2301	INFO-MSA-OnCall-NotAvail	On-call support NONE	INFO-MSA-OnCall-NotAvail
2298	INFO-MSA-OnCall-StdHrs_8-5	On-call support (billable)	INFO-MSA-OnCall-StdHrs_8-5
2323	INFO-MSA-PasswordManager	Password Manager	INFO-MSA-PasswordManager
2321	INFO-MSA-RansomwareProtection	Ransomware	INFO-MSA-RansomwareProtection
2332	INFO-MSA-NetworkMaintenance	Network Equipment Maintenance	INFO-MSA-NetworkMaintenance
2329	INFO-MSA-SecurityAwarenessTraining	Security Awareness Training	INFO-MSA-SecurityAwarenessTraining
2324	INFO-MSA-SSLManagement	SSL Management	INFO-MSA-SSLManagement
2326	INFO-MSA-Vulnerability scanning	Vulnerability scanning	INFO-MSA-Vulnerability scanning
#>


# Import CSV file with client's agreement and product details
$companies = Import-CSV '.\ServiceAgreementDetails.csv'

# Import CSV file with Connectwise product IDs, names and descriptions
$productCatalog = Import-CSV ".\ProductSetupList.csv"

# Get Non-TW Additions for now
$productCatalog = $productCatalog | ? {$_.'Product ID' -notmatch "INFO-MSA-TW*"}

foreach ($company in $companies) {
    # Get agreement ID of current company
    $endpoint = $baseUri + "finance/agreements?conditions=company/name=`"$($companies.CompanyName)`"%20and%20type/name=`"$($companies.AgreementType)`"&fields=id"
    $agreementId = (Invoke-RestMethod -Uri $endpoint -Method 'GET' -Headers $headers).id

    # Find the matching agreement in Connectwise
    $endpoint = $baseUri + "finance/agreements/$($agreementId)/additions?fields=product/identifier&Max%20Size=100"
    $currentAdditions = (Invoke-RestMethod -Uri $endpoint -Method 'GET' -Headers $headers).product.identifier
    
    # Check to see if:
        # Current column is blank
        # Contains "Not Implemented" 
        # Starts with "TW" (ThirdWall Policies)
        # Starts with "INFO" (Informational policies)
    # Add additions that pass all conditions to $activeAdditions
    $activeAddtions = @()
    foreach ($column in $company.PSObject.Properties) {
        if($column.value -and $column.value -ne "Not Implemented" -and ($column.name).StartsWith("TW") -eq $false -and ($column.name).StartsWith("INFO") -eq $true) {
            $activeAddtions += $column.name
        }
    }

    # Compare list of implemented addtions in CSV file to currently configured additions in Connectwise
    # For any addtions that are not configured in Connectwise, create them with value and description
    foreach ($addition in $activeAddtions) {
        if ($currentAdditions -contains $addition) {
            Write-Host "$($addition) already exists" -ForegroundColor DarkGreen
        } else {
            Write-Host "Creating $($addition)" -ForegroundColor DarkRed

            # Enable prorate flag
            $endpoint = $baseUri + "finance/agreements/$($agreementId)"
            Invoke-RestMethod -Uri $endpoint -Method Patch -Headers $headers -Body $agrBody1
            
            # Create new addition
            $newProductId = $productCatalog | Where-Object "Product ID" -eq $addition | Select-Object -ExpandProperty "IV_Item_RecID"
            
            $additionBody = @{
                product = @{
                    id = $newProductId
                }
                quantity = 1
                billCustomer = "DoNotBill"
            } | ConvertTo-Json

            $additionBody 

            $endpoint = $baseUri + "finance/agreements/$($agreementId)/additions/"
            Invoke-RestMethod -Uri $endpoint -Method 'Post' -Headers $headers -Body $additionBody

            # Disable prorate flag
            $endpoint = $baseUri + "finance/agreements/$($agreementId)"
            Invoke-RestMethod -Uri $endpoint -Method Patch -Headers $headers -Body $agrBody2
        }
    }
}
