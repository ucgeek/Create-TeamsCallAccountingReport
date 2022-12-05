#region Settings
$fromDateTimeString = "2020-06-01" #"yyyy-MM-dd" "2020-08-08"
$toDateTimeString = "2020-08-09" #"yyyy-MM-dd"
$currency = "NZD"
$type = "getDirectRoutingCalls"
#Client (application) ID, tenant (directory) ID and secret - created in Azure see https://github.com/leeford/Get-TeamsPSTNCallRecords/blob/master/README.md
$clientId = "7ad7c91a-xxxx-xxxx-xxxx-fe9eb86f9a6f"
$tenantId = "94a2a205-xxxx-xxxx-xxxx-8805dc057d4a"
$clientSecret = 'xxxxxxx'

# Import config from Csv files
# RateCard / Gateway mappings (not implemented yet)
$gateways = Import-Csv ".\Gateways.csv"
if ($gateways.Count -gt 0) {
    Write-Host "Found gateway rate card mappings for the following gateways:"
    foreach ($gateway in $gateways) {        
        $gateway.PstnGateway
    }
}
else {
    Write-Host "Using default rate card RateCardDefault.csv"
    $gateways = $null
}

$rateCard = Import-Csv ".\RateCardDefault.csv"
$rateCard = $rateCard | sort -Property NumberPrefix -Descending #| select -Unique
$specialNumbers = Import-Csv ".\RateCardOveridesDefault.csv"

$isRateInboundCallsEnabled = $false
$inboundCallRate = 0.01
$localCallRate = 0.02 #Used when the caller and callee are in the same calling region, based on the rate card
$isIncludeNonRateableRecordsEnabled = $false #Call forwards and transfers include a duplicate inbound call record that shares a correlationId to the actual forward or transfer record - we bill/rate calls on the latter.
#endregion Settings

#region Classes
class callRating {
    $calleeNumberPrefix
    $calleeRate
    $calleeDescription
    $calleeSubDescription
    $callerNumberPrefix
    $callerRate
    $callerDescription
    $callerSubDescription
    $isLocalCall
}

class RatedCall {
    $UserPrincipalName
    $StartDateTime
    $EndDateTime   
    $CallerNumber
    $CalleeNumber            
    $FromDescription
    $ToDescription           
    $PstnGateway     	
    $CallDirection         
    $Duration
    $Rate
    $CallCharge
    $ChargeTo
    $ChargeToCompany
    $ChargeToDepartment
    $Id 
    $CorrelationId
    $SuccessfulCall
    $Debug
    $ReferredByNumberRegion
}
#endregion Classes

#region Functions
function GetCallRates ($callerNumber, $calleeNumber) { 
    <#
    Write-Host "caller: $callerNumber"
    Write-Host "Called: $calleeNumber"
    $callerNumber = "1580592"
    $calledNumber = "1580592"
    #>

    $callerRegion = $rateCard | Where-Object { $callerNumber -like "$($_.NumberPrefix)*" -or $callerNumber -like "+$($_.NumberPrefix)*" } | Select-Object -First 1
    $calleeRegion = $rateCard | Where-Object { $calleeNumber -like "$($_.NumberPrefix)*" -or $calleeNumber -like "+$($_.NumberPrefix)*" } | Select-Object -First 1

    if ($callerRegion -eq $calleeRegion) {
        Write-Host "Local call - Caller and callee are in the same calling region"
        return [callRating]@{                                      
            calleeNumberPrefix = $calleeRegion.NumberPrefix
            calleeRate        = $localCallRate
            calleeDescription     = $calleeRegion.Description
            calleeSubDescription = $calleeRegion.SubDescription
            callerNumberPrefix = $callerRegion.NumberPrefix
            callerRate        = $callerRegion.Rate
            callerDescription     = $callerRegion.Description
            callerSubDescription = $callerRegion.SubDescription
            isLocalCall       = $true
        }
    }
    elseif ($callerRegion -ne $calleeRegion) {
        return [callRating]@{                                      
            calleeNumberPrefix = $calleeRegion.NumberPrefix
            calleeRate        = $calleeRegion.Rate
            calleeDescription     = $calleeRegion.Description
            calleeSubDescription = $calleeRegion.SubDescription
            callerNumberPrefix = $callerRegion.NumberPrefix
            callerRate        = $callerRegion.Rate
            callerDescription     = $callerRegion.Description
            callerSubDescription = $callerRegion.SubDescription
            isLocalCall       = $false
        }
    }
}
#rateCall -calledNumber $item.calleeNumber -callerNumber $item.callerNumber

function CalculateCallCost ($callRate, $durationSecs) {
    #$callRate = [math]::Round(($ratedCall.calleeRate),4)
    #$durationSecs = [math]::Round(($item.duration),4)
    if ($durationSecs -le 60) {    
        #Charge 1 min minimum
        return [math]::Round(($callRate * 1), 4)
    }
    else {
        #Charge actual duration
        return [math]::Round(($callRate * ($durationSecs / 60)), 4)
    }
}
#endregion Functions

#region Main Script
$uri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
$body = @{
    client_id     = $clientId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $clientSecret
    grant_type    = "client_credentials"
}

#Get OAuth Access Token
$tokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing
#Set Access Token
$token = ($tokenRequest.Content | ConvertFrom-Json).access_token
#Build API call URI
$currentUri = "https://graph.microsoft.com/beta/communications/callRecords/$type(fromDateTime=$fromDateTimeString,toDateTime=$toDateTimeString)"

Write-Host "Checking for call records between $fromDateTimeString and $toDateTimeString..." -ForegroundColor Cyan
$content = @()
$content += while (-not [string]::IsNullOrEmpty($currentUri)) {
    $apiCall = Invoke-RestMethod -Method "GET" -Uri $currentUri -ContentType "application/json" -Headers @{Authorization = "Bearer $token" } -ErrorAction Stop    
    $currentUri = $null

    if ($apiCall) {
        #Check if any data is left
        $currentUri = $apiCall.'@odata.nextLink'

        #Count total records so far
        $totalRecords += $apiCall.'@odata.count'

        $apiCall.value
    }
}

Write-Host "Processing $($content.Count) record(s) retrived via API" -ForegroundColor Cyan
$ratedCalls = @()
foreach ($item in $content) {
    $isCorrelatedRecord = ($content | where { $_.correlationId -eq $item.correlationId } | Measure-Object).Count -gt 1

    if ($item.calltype -eq 'ByotOut') {

        $callDirection = "Outbound Call"
    }
    elseif ($item.calltype -eq 'ByotIn') {
        # If more than two correlation IDs (two linked calls)
        if ($isCorrelatedRecord) {
            Write-Host "Call has related records: $($item.callerNumber) -> $($item.calleeNumber) - Call charge will not be billed on this line. Instead the charge will apply to the transfer or forwarding line"
            if (!$isIncludeNonRateableRecordsEnabled){
                continue
            }
        }
        else {
            Write-Host "Inbound call: $($item.callerNumber) -> $($item.calleeNumber)"
        }

        $callDirection = "Inbound Call"
    }
    elseif ($item.calltype -eq 'ByotOutUserForwarding') {
        Write-Host "Forwarded call: $($item.callerNumber) -> $($item.calleeNumber)"

        $callDirection = "Outbound Forward"

    }
    elseif ($item.calltype -eq 'ByotOutUserTransfer') {
        Write-Host "Transfered call: $($item.callerNumber) -> $($item.calleeNumber)"
        $callDirection = "Outbound Transfer"

    }

    # Don't charge inbound calls
    if ($item.callType -notcontains "ByotIn") {
        $ratedCall = GetCallRates -calleeNumber $item.calleeNumber -callerNumber $item.callerNumber
        $callRate = [math]::Round(($ratedCall.calleeRate), 4)
        $durationSecs = [math]::Round(($item.duration), 4)
        $calculatedCharge = CalculateCallCost $callRate $durationSecs
    } else {
        $ratedCall = GetCallRates -calleeNumber $item.calleeNumber -callerNumber $item.callerNumber
        $callRate = 0
        $durationSecs = [math]::Round(($item.duration), 4)

        if ($isRateInboundCallsEnabled -and !$isCorrelatedRecord){
            $calculatedCharge = $inboundCallRate
        } else {
            $calculatedCharge = 0
        }      
    }

    $ratedCalls += [RatedCall]@{
        UserPrincipalName      = $item.userPrincipalName
        StartDateTime          = $item.startDateTime
        EndDateTime            = $item.endDateTime
        CallerNumber           = $item.callerNumber
        CalleeNumber           = $item.calleeNumber
        FromDescription        = "$($ratedCall.callerDescription) :: $($ratedCall.callerSubDescription)"
        ToDescription          = "$($ratedCall.calleeDescription) :: $($ratedCall.calleeSubDescription)"
        PstnGateway            = $item.trunkFullyQualifiedDomainName
        CallDirection          = $callDirection
        Duration               = $durationSecs
        Rate                   = $callRate
        CallCharge             = $calculatedCharge
        ChargeTo               = ("$($item.userDisplayName) ($($item.userPrincipalName))")
        ChargeToCompany        = "**Need to do an AD lookup**"
        ChargeToDepartment     = "**Need to do an AD lookup**"
        Id                     = $item.Id
        CorrelationId          = $item.correlationId
        SuccessfulCall         = $item.successfulCall
        Debug                  = "SipCode: $($item.finalSipCode) MsCode: $($item.callEndSubReason) Info: $($item.finalSipCodePhrase)"
        ReferredByNumberRegion = "**Need to find a way to related records - so far looks difficult**"
    }

}
#endregion Main Script
#$ratedCalls = $null
Write-Host "Rated $($ratedCalls.Count) call records" -ForegroundColor Cyan
$ratedCalls | Select-Object userPrincipalName, CallerNumber, CalleeNumber, CallDirection, Duration, Rate, CallCharge, FromDescription, ToDescription, Id | Format-Table
$ratedCalls | Export-Csv -Path ".\Output.csv"

