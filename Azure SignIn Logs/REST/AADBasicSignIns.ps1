<#
Script to get Basic Auth sign-ins using Microsoft Graph API (REST) and export to .CSV
https://docs.microsoft.com/en-us/graph/api/signin-list?view=graph-rest-1.0&tabs=http

Required Permissions:AuditLog.Read.All AND Directory.Read.All (There is a know issue and currently BOTH permissions are required)
#>

#Variables
$length = -7 #<--Last 7 days
$date = (get-date).AddDays($length)
$sDate = $date.ToString("yyyy-MM-dd")
[System.Collections.ArrayList]$reportData = @()
$fileLocation = "C:\temp\temp.csv"

#----FUNCTIONS----
#Get access token from Graph API
Function Get-AccessToken
{
    $AppId = "" # <- Application ID from AAD App Registration
    $AppSecret = "" # <- App secret from AAD App Registration
    $tenantID = "" # <-- tenantID from AAD App Registration
    $Scope = "https://graph.microsoft.com/.default"
    $Url = "https://login.microsoftonline.com/$tenantID/oauth2/v2.0/token" # < - OAuth 2.0 token endpoint (v2) from AAD App Registration

    # Create body
    $Body = @{
        client_id     = $AppId
        client_secret = $AppSecret
        scope         = $Scope
        grant_type    = 'client_credentials'
    }

    # Splat the hashtable
    $Splat = @{
        ContentType = 'application/x-www-form-urlencoded'
        Method      = 'POST'
        Body        = $Body
        Uri         = $Url
    }
    # Submit Request for the Token
    $Request = Invoke-RestMethod @Splat
    return $Request
}

#----MAIN----
#Get OAuth Token
$TokenCache = Get-AccessToken
$Token = $TokenCache.access_token

#Make Requests to the Graph API to get Azure Sign-In logs
$uri = "https://graph.microsoft.com/v1.0/auditLogs/SignIns?`$filter=CreatedDateTime ge $sDate and clientAppUsed ne 'Browser' and clientAppUsed ne 'Mobile Apps and Desktop clients'"
$method = "GET"

do
{
    $siLogs = Invoke-RestMethod -Method $method -Headers @{Authorization = "Bearer $($Token)"; "Content-Type" = "application/json" } -Uri $uri
    $reportData +=$siLogs.value
    $uri = $siLogs.'@odata.nextlink'    
}until([String]::IsNullOrEmpty($uri))
$reportData | export-csv $fileLocation -NoTypeInformation
