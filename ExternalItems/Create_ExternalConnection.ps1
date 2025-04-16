<#
Create_ExternalConnection.ps1
Description:
    This script creates an external connection in Microsoft Graph using the Microsoft Graph PowerShell SDK.
    It retrieves authentication details from a .env file and uses them to connect to Microsoft Graph.
    The script defines the connection parameters and creates a new external connection with the specified name and description.
#>


#auth
$data = get-content -Path .env
$appID = ($data[0].split("="))[1]
$tenantID = ($data[1].split("="))[1]
$authCertThumb = ($data[2].split("="))[1]

Connect-MGGraph -ClientId $appID -TenantId $tenantID -CertificateThumbprint $authCertThumb

$ConnectionName = "DemoGraphConnector"

$connectionParams = @{
    id = $ConnectionName
    name = $ConnectionName
    description = "Test connector called $ConnectionName. Containing a list of company names."
}

New-MgExternalConnection -BodyParameter $connectionParams

Disconnect-MGGraph