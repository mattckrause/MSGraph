<#
Register_Schema.ps1
Description:
    This script registers a schema for an external connection in Microsoft Graph using the Microsoft Graph PowerShell SDK.
    It retrieves authentication details from a .env file and uses them to connect to Microsoft Graph.
    The script defines the schema parameters and updates the external connection with the specified schema.
#>

#auth
$data = get-content -Path .env
$appID = ($data[0].split("="))[1]
$tenantID = ($data[1].split("="))[1]
$authCertThumb = ($data[2].split("="))[1]

Connect-MGGraph -ClientId $appID -TenantId $tenantID -CertificateThumbprint $authCertThumb -NoWelcome


$ConnectionName = "DemoGraphConnector"

$schemaParams = @{
    baseType = "microsoft.graph.externalItem"
    properties = @(
        @{
            name = "CompanyName"
            type = "String"
            isSearchable = "true"
            isRetrievable = "true"
            labels = @(
            "title"
        )
    }
)
}

Update-MgExternalConnectionSchema -ExternalConnectionId $ConnectionName -BodyParameter $schemaParams

Disconnect-MGGraph