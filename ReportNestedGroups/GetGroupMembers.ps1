<#
Report all members of a group. Processing users in nested groups as well.
Requires Microsoft Graph API PowerShell SDK (https://docs.microsoft.com/en-us/graph/powershell/installation)
#>

#Connect to Graph API using App Registration
$appID = "" #<- Application Id from App registration
$tenantID = "" #<- Tenant Id from App registration
$certThumb = "" #<- Certificate Thumbpring of cert uploaded to App registration

Connect-MgGraph -ClientId $appID -TenantId $tenantID -CertificateThumbprint $certThumb

#variables
$report = "C:\temp\GroupMembers.csv"
$groupName = "testgroup1"
[System.Collections.ArrayList] $output =  @()

#lookup group
$group = Get-MgGroup -Filter "DisplayName eq '$groupName'"

#Get all transitive members
$groupMembers = Get-MgGroupTransitiveMember -GroupId $group.Id

#Add only User objects to the List for export
$groupMembers | foreach-object{ `
    if($_.AdditionalProperties.'@odata.type' -like "*user") `
        {write-host $_.AdditionalProperties.displayName `
        $output.Add($_.AdditionalProperties.displayName) | Out-Null }}

#Write data to .csv file
$output | Out-File $report

#Print summary to console
Write-Host $output.count "Users exported to $report"

#Disconnect from Graph
Disconnect-MgGraph #<-- Remember to disconnect from MS Graph API when the app is done!
