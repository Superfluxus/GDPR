#Function to check how many API calls we have left in our 1 hour period
function Get-APIThreshold-PROD {
    $Config = Get-Content -Path "C:\Windows\#REDACTED#\GDPR\config.json" -Raw | ConvertFrom-Json

    $Apiurl = "https://webservices.autotask.net/atservices/1.6/atws.wsdl"
    $ApiUsername = $($Config.PsaUser)
    $ApiPassword = $($Config.PsaPassword) | ConvertTo-SecureString -AsPlainText -Force
    $ApiCreds = New-Object System.Management.Automation.PSCredential($ApiUsername, $ApiPassword)
    
    $Integration#REDACTED# = "$($Config.PsaIntegration#REDACTED#)"
    $UserZone = "https://webservices17.autotask.net/ATServices/1.6/atws.asmx"
    
    #Set all connection to use TLS 1.2 (I think)
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    
    $Api = New-WebServiceProxy -Uri $ApiUrl -Credential $ApiCreds
    
    $Query = '<?xml version="1.0" encoding="utf-8"?>
    <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
      <soap:Header>
        <AutotaskIntegrations xmlns="http://autotask.net/ATWS/v1_6/">
          <Integration#REDACTED#>#REDACTED#</Integration#REDACTED#>
        </AutotaskIntegrations>
      </soap:Header>
      <soap:Body>
        <getThresholdAndUsageInfo xmlns="http://autotask.net/ATWS/v1_6/">
        </getThresholdAndUsageInfo>
      </soap:Body>
    </soap:Envelope>'
    
    $Request = Invoke-WebRequest -Uri $UserZone -Credential $ApiCreds -Method Post -Body $Query -ContentType 'text/xml; charset=utf-8'
    
    [xml]$Response = $Request.Content
    
    $XML = $response.Envelope.Body.getThresholdAndUsageInfoResponse.getThresholdAndUsageInfoResult.EntityReturnInfoResults.EntityReturnInfo.Message
    
    $XML.numberOfExternalRequest

    $Count1 = $XML.Split(" ")[-1]
    $Count1 = $Count1.TrimEnd(";")
    $Count2 = [int]$Count1
    $Count3 = 10000 - $Count2
    
    #Write-Host "You have currently used $Count2 requests in the last hour"
    #Write-host "You have $Count3 remaining"

    Return $Count3


}
#Function end, script start


function Write-Log {
    #This function is used for all log writing to a set format
    Param (
        [Parameter(Mandatory = $True)]
        $Text,

        [Parameter(Mandatory = $True)]
        [ValidateSet('API', 'Ticket', 'Maintenance', 'Warning', 'Wait', 'Info', 'Error')]
        $Type
    )
    #Sets Log File Location
    $Date = Get-Date -Format yyyy-MM-dd
    $TIME = Get-Date -Format HH:mm:ss
    $Path = "C:\Windows\#REDACTED#\GDPR\\log\$($Date) - Full_Log.csv"

    New-Object -TypeName psobject -Property @{Date = $DATE
        Time                                       = $TIME
        Type                                       = $Type
        Information                                = $Text
        Computer                                   = $env:COMPUTERNAME
    } | 
        Select-Object Date, Time, Type, Information | 
        Export-Csv -Append -NoTypeInformation -Path $Path

    #Enable the below line for debug
    #Write-host "$(Get-Date) - $Type - $Text"
}

################# Setup global variables
$Redacted = 0
$Date = [DateTime](Get-Date).AddDays(-28)

#Connect to API
$Config = Get-Content -Path "C:\Windows\#REDACTED#\GDPR\config.json" -Raw | ConvertFrom-Json
$Password = $Config.PsaPassword | ConvertTo-SecureString -AsPlainText -Force
$Creds = New-Object System.Management.Automation.PSCredential ($Config.PsaUser, $Password)

If ( $(Get-InstalledModule -Name Autotask) -ne $None) {
    try {
        Write-Log -Type Info -Text "Importing PSA module..."

        #Set-Variable AtwsNoDiskCache -Scope Global -Value $True
        #Import-Module Autotask -ArgumentList $Creds, $Config.PsaIntegration#REDACTED#
  
        Import-Module -Name Autotask -DisableNameChecking

        Connect-AtwsWebAPI -Credential $Creds -ApiTrackingIdentifier $Config.PsaIntegration#REDACTED# -NoDiskCache

        #Run this to test if the connection to Datto was successful. Account 0 is #REDACTED#
        Get-AtwsAccount -id 0 | Out-Null

        Write-Log -Type Info -Text "Successfully imported PSA module"    
    }
    catch {
        Write-Log -Type Error -Text "Connection to Datto failed. Are your credentials and integration #REDACTED# in config.json correct?"

        Start-Sleep -Seconds 3
        Exit 1
    }    
}
#Check for API limit and gather all tickets to search through

$APICheck = Get-APIThreshold-PROD
$APICheck = [int]$APICheck[1]
Write-Log -Type API -Text "We have $APICheck API Calls remaining for this hour"
$Counter = 0
$Found = 0
Write-Log -Type Info -Text "Gathering tickets... (This may take a while)"
$Tickets = get-atwsticket -TicketCategory "#REDACTED#" -status #REDACTED# -QueueID "#REDACTED#" | where-object { $_.ResolvedDateTime -GE $Date }
Write-Log -Type Info -Text "We found $($Tickets.Count) tickets"

$Ticket1 = $Tickets | Where-Object {$_.Status -eq "Refund - Processed"}
$Ticket2 = $Tickets | Where-Object {$_.Status -eq "Replacement - Processed"}

if ($($Tickets.Count) -eq 0) {
    Write-Log -Type Info -Text "No tickets met the criteria for cleanup! All clear."
    Break
}

$APICheck = Get-APIThreshold-PROD
$APICheck = [int]$APICheck[1]
Write-Log -Type API -Text "We have $APICheck API Calls remaining for this hour"
if ($APICheck -le 2100) {

    Write-Log -Type API "API Limit hit, sleeping..."

    $Now = (get-date).tostring("HH:mm:ss")
    $1HourPlus = (get-date).AddHours(1).tostring("HH:mm:ss")
    Write-Log -Type API -Text "Sleep started at $Now"
    Write-Log -Type API -Text "Operation will resume at $1HourPlus"

    Start-Sleep -Seconds 900

    Write-Log -Type API -Text "Ready to restart"
}
    
#Begin searching through gathered tickets
Write-Log -Type Ticket -Text "Working with $($Tickets.Count) tickets"
Write-Log -Type Ticket -Text "$($Ticket1.Count) REFUND tickets"
Write-Log -Type Ticket -Text "$($Ticket2.Count) REPLACEMENT tickets"

foreach ($Ticket in $tickets) {
    #if Current ticket processed is a number divisible exactly by 1000, check our API limit. If it's below X (currently 1100), sleep for an hour
    if (($Counter % 1000) -eq $True) {
        Write-Log -Type API -Text "Checking API Calls remaining..."
        $APICheck = Get-APIThreshold-PROD
        $APICheck = [int]$APICheck[1]
        if ($APICheck -le 1100) {
            Write-Log -Type API -Text "API Limit hit, sleeping..."
            Write-Log -Type API -Text "We've found $Found matches so far!"        
            $Now = (get-date).tostring("HH:mm:ss")
            $1HourPlus = (get-date).AddHours(1).tostring("HH:mm:ss")
            Write-Log -Type API -Text "Sleep started at $Now"
            Write-Log -Type API -Text "Operation will resume at $1HourPlus"
            Start-Sleep -Seconds 900
            Write-Log -Type API -Text "Ready to restart"
            Write-Log -Type API -Text "Checking API calls"
            $APICheck2 = Get-APIThreshold-PROD
            $APICheck2 = [int]$APICheck[1]
            $Gained = [int]$APICheck2 - [int]$APICheck
            Write-Log -Type API -Text "$Gained API Calls freed up by sleeping, continuing operation"

        }
        
    }
    else { Write-Log -Type Info -Text "We have $Found matches so far" }
    
    $Counter ++
    Write-Log -Type Ticket -Text "Ticket #$Counter"
    Write-Log -Type Ticket -Text "Looking at $($Ticket.TicketNumber)..."
    $Title = Get-ATWSTicket -id $($Ticket.ID)
    $Title = $($Title.Title)
    $RedactMe = Get-ATWSTicketNote -TicketID $($Ticket.ID)
    Write-Log -Type Ticket -Text "$($Ticket.TicketNumber) has $($RedactMe.Count) notes on it"
    Write-Log -Type Ticket -Text "Begin overwrite"
    $NoteCount = 0
    foreach ($Note in $RedactMe) {       
        $NoteCount ++
        Write-Host -ForegroundColor Yellow "Redact Note #$NoteCount"
        Set-ATWSTicketNote -Id $($Note.ID) -Title "REDACTED" -Description "REDACTED" -erroraction SilentlyContinue -warningaction SilentlyContinue
        $Redacted ++
            
    }  
    Set-ATWSTicket -ID $($ticket.ID) -Description #REDACTED#
                 
    else {
        Write-Log -Type Error -Text "$($Ticket.TicketNumber) does not match"
    }
    
}
   
Write-Log -Type Info -Text "$Redacted notes purged"
Write-Log -Type Info -Text "$($Ticket1.Count) REFUND tickets"
Write-Log -Type Info -Text "$($Ticket2.Count) REPLACEMENT tickets"
Write-Log -Type Info -Text "$($Tickets.Count) total tickets"
