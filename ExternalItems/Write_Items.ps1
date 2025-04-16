<#
write_item.ps1
Description:
    This script writes items to an external connection in Microsoft Graph using the Microsoft Graph PowerShell SDK.
    It retrieves authentication details from a .env file and uses them to connect to Microsoft Graph.
    The script reads a CSV file containing fictitious company names and descriptions, and writes each item to the external connection.
#>

#auth
$data = get-content -Path .env
$appID = ($data[0].split("="))[1]
$tenantID = ($data[1].split("="))[1]
$authCertThumb = ($data[2].split("="))[1]

Connect-MGGraph -ClientId $appID -TenantId $tenantID -CertificateThumbprint $authCertThumb -NoWelcome

$ConnectionName = "DemoGraphConnector"

#write items
Import-Csv -Path "C:\github\MyScripts\GraphAPI\PS_SDK\ExternalItems\fictitious_companies.csv" | ForEach-Object {
    $params = @{
        acl = @(
            @{
                type = "everyone"
                value = "everyone"
                accessType = "grant"
            }
        )
        properties = @{
            CompanyName = $_.name
        }
        content = @{
            value = $_.description
            type = "text"
        }
    }

    Set-MgExternalConnectionItem -ExternalConnectionId $ConnectionName -ExternalItemId $id -BodyParameter $params
}

Disconnect-MGGraph
