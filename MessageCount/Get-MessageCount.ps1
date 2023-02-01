<#
PS Script to use Graph API get email message count for all users in a group/dl
requires date in format "2023-01-30"
Matt Krause - Microsoft Graph CPX
#>

#get parameters for Script
Param(
    [Parameter(Mandatory=$true, Position=0)]
    [String]$groupName,
    [Parameter(Mandatory=$true, Position=1)]
    [String]$startDate,
    [Parameter(Mandatory=$true, Position=2)]
    [String]$endDate
)

#Variable Declaration
$reportFile = "C:\github\MyScripts\GraphAPI\PS_SDK\Mail\MessageCount\messages.csv" # <- CSV Report
[System.Collections.ArrayList]$MBmessageCount = @()
$totalMessageCount = 0

#----FUNCTIONS----
#Connect to Graph using appID and certificate
Function psmConnectToGraph
{
    $authCertThumb = "DB807E7E156FB01919A517971208A3EEE3CB0087"
    $tenantID = "7b2828b9-89a3-4507-9e1a-05ff46d1192d"
    $appID = "4b6a9db7-646c-4c4e-b083-1cce4169f424"

    Connect-MGGraph -ClientId $appID -TenantId $tenantID -CertificateThumbprint $authCertThumb
}

Function Get-GroupMember
{
    Param(
        [Parameter(Mandatory=$true, Position=0)]
        [ValidateNotNullOrEmpty()]
        [string]$groupName
    )
    $groupID = Get-MgGroup -Search "displayName:$groupName" -Select Id -ConsistencyLevel eventual
    $groupMembers = Get-MgGroupMember -GroupId $groupID.Id
    return $groupMembers
}

Function Count-Messages
{
    Param(
        [Parameter(Mandatory=$true, Position=0)]
        [ValidateNotNullOrEmpty()]
        [String]$mailbox,
        [Parameter(Mandatory=$true, Position=1)]
        [ValidateNotNullOrEmpty()]
        [String]$start,
        [Parameter(Mandatory=$true, Position=2)]
        [ValidateNotNullOrEmpty()]
        [String]$end
        )
    $temp = Get-MgUserMessage -UserId $mailbox -Filter "ReceivedDateTime ge $start and ReceivedDateTime le $end" -Select "Id" -All -CountVariable MessageCount
    return $MessageCount
}

#Main Script
#authenticate
psmConnectToGraph

#Lookup Group for members and then lookup messages for each member.
$groupUsers =  Get-GroupMember -groupName $groupName
$groupUsers.Id | ForEach-Object {$count = Count-Messages -mailbox $_ -start (get-date $startDate).ToString('yyyy-MM-dd') -end (get-date $endDate).ToString('yyyy-MM-dd');
    $val = [PSCustomObject]@{ 'MailboxID' = $_; 'MessageCount' = $count}
    $totalMessageCount += $count
    $MBmessageCount.Add($val) | out-Null; $val = $null}

#
$MBmessageCount |  Export-Csv $reportFile -NoTypeInformation
write-host "$totalMessageCount total messages for MGDC extraction."
Disconnect-MgGraph
