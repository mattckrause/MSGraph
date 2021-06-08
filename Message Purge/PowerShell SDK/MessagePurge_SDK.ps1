<#
Script to lookup and purge a message using Graph API PowerShell Module
Authentication using application permissions
#>

#get/set Archive parameter
Param(
    [bool]$archive = $false
)

#Variable Declaration
$dataFile = "C:\GitHub\MyScripts\GraphAPI\PSModule\MessagePurge\messages.csv" # <- CSV Import File
$archiveMailbox = "" #<- Mailbox to be used as archive.

#----FUNCTIONS----
#Import CSV File
Function Import-Data
{
    Param(
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipeLineByPropertyName = $true,
            ValueFromRemainingArguments = $false,
            Position = 0)]
        [ValidateNotNullOrEmpty()]
        [String]$File
    )
    $returnValues = import-csv $File
    return $returnValues
}

#Connect to Graph using appID and certificate
Function psmConnectToGraph
{
    #$authCert = "GraphPSModule1"
    $authCertThumb = "" #<- Certificate Thumbprint from App registration
    $tenantID = "" #<- Tenant ID from App registration
    $appID = "" #<- Application ID from App registration

    Connect-MgGraph -ClientId $appID -TenantId $tenantID -CertificateThumbprint $authCertThumb
}

Function purgeMessage
{
    Param(
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            ValueFromRemainingArguments = $false,
            Position = 0)]
        [ValidateNotNullOrEmpty()]
        [array]$Data,
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            ValueFromRemainingArguments = $false,
            Position = 1)]
        [ValidateNotNullOrEmpty()]
        [bool]$arch
    )

    foreach ($record in $Data)
    {
        #Lookup ID of message based on internet Message ID
        $message = get-mgusermessage -UserId $record.Recipient -Filter "internetMessageId eq '$($record.Message_id)'"

        #Determine if Archive is needed
        If ($arch -eq "true")
        {
            #archive message before delete
            $archBuildBody = [ordered]@{
                "comment"      = "Forwarded as archive from message purge script.";
                "toRecipients" = @(
                    @{"emailAddress" = @{"address" = $archiveMailbox; "name" = "Archive Mailbox" } }
                )
            }
            Move-MgUserMessage -MessageId $message.Values.id -UserId $record.Recipient -BodyParameter $archBuildBody
        }

        #Remove message from mailbox
        remove-mgusermessage -MessageId $message.Id -UserId $record.Recipient
        #Report results
        write-Host "The message was found and removed. The sender is: $($message.from.emailaddress.Name). The subject is: $($message.Subject), and the ID is: $($message.Id)." -ForegroundColor Green
    }
}

#Main Script
#Get User/Message Data from CSV File
$userData = Import-Data -File $dataFile
#Connect to Graph API
psmConnectToGraph
#Purge messages from mailboxes
purgeMessage -Data $userData -arch $archive

#Disconnect from Graph when complete!
Disconnect-MgGraph
