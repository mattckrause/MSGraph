<#
Script to get the last login time from all users based on Azure Sign-In logs
Using Microsoft Graph PowerShell SDK
#>

#Connect to Graph
$appID = "" #<- Application ID from app registration
$tenantID = "" #<- Tenant ID from app registration
$certThumb = "" #<- Certificate tumbprint of cert uploaded to app registration

Connect-MgGraph -ClientId $appID -TenantId $tenantID -CertificateThumbprint $certThumb

#connect to beta endpoint (functionality currently only exists here)
select-mgprofile -Name beta

#setting up an arraylist for saving data
[System.Collections.ArrayList] $output =  @()
$output.add("DispalyName,LastSignInDateTime (UTC)")

#Getting lastsignindatetime from a MS Graph requests
Get-MgUser -all -Property SignInActivity| foreach-object{ `
        if($null -ne $_.SignInActivity.LastSignInDateTime) `
            {$output.add($_.DisplayName +","+ $_.SignInActivity.LastSignInDateTime) | Out-Null} `
        else `
            {$output.add($_.DisplayName +",No Logon") | Out-Null}}

$output | out-file -FilePath "C:\temp\logins.csv" #<- Update this to change the report file location

Disconnect-MgGraph #<-- Remember to disconnect from MS Graph API when the app is done!
