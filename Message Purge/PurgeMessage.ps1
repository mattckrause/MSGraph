<#
Script to purge messages from mailboxes using MS Graph API/PowerShell
Application Premissions and client credential auth flow used for Graph Access
#>

#get/set Archive parameter
Param(
    [bool]$archive = $false
)

#Variable declaration
$dataFile = "C:\MessagePurge\messages.csv" # <- CSV Import File
$url = "https://graph.microsoft.com/v1.0/"
$archiveMailbox = "archivemailbox"

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
    $AppId = "XXXXX" # <- Application ID from AAD App Registration
    $AppSecret = "XXXXX" # <- App secret from AAD App Registration
    $Scope = "https://graph.microsoft.com/.default"
	$Url = "https://login.microsoftonline.com/XXXXX/oauth2/v2.0/token" # < - OAuth 2.0 token endpoint (v2) from AAD App Registration

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

#Message Purge using Graph API
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
            [bool]$arch,
            [Parameter(Mandatory=$true,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true,
                    ValueFromRemainingArguments=$false,
                    Position=2)]
            [ValidateNotNullOrEmpty()]
            [String]$Token
        )

	foreach ($record in $Data)
	{
        #set user and message ID
        $user = $record.Recipient
        $rawmID = $record.message_id

        #call function to encode search string
        $mID = urlencode -ImID $rawmID

        $queryURL = $url + "users/" + $user + "/messages?`$filter=internetMessageId eq " +  "`'$($mID)`'"
        $queryResponse = Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token)"} -Uri $queryURL -Method Get
        
        #is archiving the message necessary?
        if ($arch -eq $true)
        {
            #forward message to "archive" mailbox
            $archURL = $url + "users/" + $user + "/messages/" + $queryResponse.value.ID + "/forward"
            $archBuildBody = [ordered]@{
                "comment" = "Forwarded as archive from message purge script.";
                "toRecipients" = @(
                    @{"emailAddress" = @{"address" = $archiveMailbox; "name" = "Archive Mailbox"}}
                )
            }

            $archBody = $archBuildBody | convertto-json -Depth 3

            Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token)"; "Content-Type" = "application/json"} -Body $archBody -Uri $archURL -Method POST
        }

        #Create URL for delete using ID collected above
        $deleteURL = $url + "users/" + $user + "/messages/" + $queryResponse.value.ID
        #delete message
        Invoke-RestMethod -Headers @{Authorization = "Bearer $($Token)"} -Uri $deleteURL -Method Delete
        #Create message delete logging here
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
$TokenCache = Get-AccessToken
$Token = $TokenCache.access_token

#Make Requests to the Graph API to delete the messages
MessagePurge -Data $userData -arch $archive -Token $Token
