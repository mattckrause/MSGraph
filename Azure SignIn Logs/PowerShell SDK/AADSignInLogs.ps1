<#
Microsoft Graph PowerShell SDK Script to get aad sign-ins using Microsoft Graph API and export to .CSV.
https://docs.microsoft.com/en-us/graph/api/signin-list?view=graph-rest-1.0&tabs=http

Required Permissions:AuditLog.Read.All AND Directory.Read.All (There is a know issue and currently BOTH permissions are required)
#>
#Variables
$length = -7 #<--Last 7 days
$date = (get-date).AddDays($length)
$sDate = $date.ToString("yyyy-MM-dd")
$fileLocation = "C:\temp\temp.csv" #<- Report File Location

#connecto to Microsoft Graph
Function psmConnectToGraph
{
    $authCertThumb = "" #<- Certificate Thumbprint from App registration
    $tenantID = "" #<- Tenant ID from App registration
    $appID = "" #<- Application ID from App registration

    Connect-MgGraph -ClientId $appID -TenantId $tenantID -CertificateThumbprint $authCertThumb
}

#Main Script
psmConnectToGraph

$auditInfo = Get-MgAuditLogSignIn -Filter "CreatedDateTime ge $sDate" -All

$auditInfo | where-object {$_.clientAppUsed -ne "Browser" -and $_.ClientAppUsed -ne "Mobile Apps and Desktop clients"} #| export-csv $fileLocation -NoTypeInformation

#Disconnect-MgGraph