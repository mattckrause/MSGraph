<#
Script to purge messages from mailboxes using MS Graph API/PowerShell
Application Premissions are used for Graph Access
#>

#Variable declaration
$dataFile = "C:\temp\messages.csv" # <- CSV Import File
$url = "https://graph.microsoft.com/v1.0/"

#----FUNCTIONS----
#Import CSV File
Function Import-Data
{
	Param(
            [Parameter(Mandatory=$true,
                    ValueFromPipeline=$true,
                    ValueFromPipeLineByPropertyName=$true,
                    ValueFromRemainingArguments=$false,
                    Position=0)]
            [ValidateNotNullOrEmpty()]
                [String]$File
    )
    $returnValues = import-csv $File
    return $returnValues
}

#Get access token from Graph API
Function Get-AccessToken
{
    $AppId = "XXXXXXXXXX" # <- Application ID from AAD App Registration
    $AppSecret = "XXXXXXXXXX" # <- App secret from AAD App Registration
    $Scope = "https://graph.microsoft.com/.default"
    $Url = "https://login.microsoftonline.com/XXXXXXXXXX/oauth2/v2.0/token" # < - OAuth 2.0 token endpoint (v2) from AAD App Registration

    # Add System.Web for urlencode
    Add-Type -AssemblyName System.Web

    # Create body
    $Body = @{
    	client_id = $AppId
        client_secret = $AppSecret
        scope = $Scope
        grant_type = 'client_credentials'
    }

    # Splat the Hash Table
    $Splat = @{
        ContentType = 'application/x-www-form-urlencoded'
        Method = 'POST'
        Body = $Body
        Uri = $Url
    }
    # Submit Request for the Token
    $Request = Invoke-RestMethod @Splat
    return $Request
}

#Message Purge using Graph
Function MessagePurge
{
    Param(
            [Parameter(Mandatory=$true,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true,
                    ValueFromRemainingArguments=$false,
                    Position=0)]
            [ValidateNotNullOrEmpty()]
            [array]$Data,
            [Parameter(Mandatory=$true,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true,
                    ValueFromRemainingArguments=$false,
                    Position=1)]
            [ValidateNotNullOrEmpty()]
            [String]$Token
        )

	foreach ($record in $Data)
	{
        # Submit Request for the Token
        $user = $record.Recipient
        $rawmID = $record.message_id

        #call function to encode search string
        write-host "Encoding URL Search String..." -ForegroundColor Green
        $mID = urlencode -ImID $rawmID

        write-host "User $user"
        Write-Host "MID $mID"
        $queryURL = $url + "users/" + $user + "/messages?`$filter=internetMessageId eq " +  "`'$($mID)`'"
        Write-Host "The queryURL is $queryURL" -ForegroundColor Red
        $queryResponse = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token)"} -Uri $queryURL -Method Get
                
        #Collect Message Info for Logging
        $mSender = $queryResponse.value.sender.emailaddress.address
        $mRecipient = $queryResponse.value.torecipients.emailaddress.address
        $mSubject = $queryResponse.value.subject
                
        #Create URL for delete using ID collected above
        $deleteURL = $url + "users/" + $user + "/messages/" + $queryResponse.value.ID
        Write-Host "The delete url is $deleteURL" -ForegroundColor Yellow
        #delete message
        Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token)"} -Uri $deleteURL -Method Delete
        #Create message delete logging here
        Write-Host "The message was deleted. It was From: $mSender, To: $mRecipient, with the Subject: $mSubject." -ForegroundColor Green
    }
}

Function urlencode
{
        Param(
                [Parameter(Mandatory=$true,
                        ValueFromPipeline=$true,
                        ValueFromPipeLineByPropertyName=$true,
                        ValueFromRemainingArguments=$false,
                        Position=0)]
                [ValidateNotNullOrEmpty()]
                [String]$ImID
        )
        #split Message ID at '@'
        $a = $ImID.Split('@')
        #Trim leading '<' from id value
        $b = $a[0].trim("<")
        #build encoded MID value
        $enCodeString = "<"+[uri]::EscapeDataString($b) + "@" + $a[1]
        
        return $enCodeString 
}
#----MAIN----
#Get User/Message Data from CSV File
$userData = Import-Data -File $dataFile

#Get OAuth Token
write-host "Getting Token..."
$TokenCache = Get-AccessToken
$Token = $TokenCache.access_token

#Make Requests to the Graph API to delete the messages
Write-host "Deleting Message from Input File..."
MessagePurge -Data $userData -Token $Token
