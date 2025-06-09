<#
New-GraphConnector.ps1
Description:
    This script creates a Microsoft Graph external connection and writes items to it.
    It requires the Microsoft.Graph module to be installed and imported.
    The script provides functions to create and remove external connections, as well as write items to the connection.
    The script uses a .env file for authentication details.
    The .env file should contain the following lines:
        APP_ID=<your_app_id>
        TENANT_ID=<your_tenant_id>
        AUTH_CERT_THUMB=<your_auth_cert_thumbprint>
Matt Krause
#>

param(
    [Parameter(Mandatory = $true,
        ValueFromPipeline = $true,
        ValueFromPipeLineByPropertyName = $true,
        ValueFromRemainingArguments = $false,
        Position = 0)]
    [ValidateNotNullOrEmpty()]
    [String]$Process
)

Function Connect-ToGraph
{
    $data = get-content -Path .env
    $appID = ($data[0].split("="))[1]
    $tenantID = ($data[1].split("="))[1]
    $authCertThumb = ($data[2].split("="))[1]

    Connect-MGGraph -ClientId $appID -TenantId $tenantID -CertificateThumbprint $authCertThumb -nowelcome
}

Function New-ExternalConnection
{
    Param(
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipeLineByPropertyName = $true,
            ValueFromRemainingArguments = $false,
            Position = 0)]
        [ValidateNotNullOrEmpty()]
        [String]$ConnectionName
    )
    $connectionParams = @{
        id = $ConnectionName
        name = $ConnectionName
        description = "Test connector called $ConnectionName created using PowerShell. Contains a list of company names and discriptions for each."
    }
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
    try{
        write-host "Creating connection $ConnectionName"
        New-MgExternalConnection -BodyParameter $connectionParams
    }
    catch{
        write-host "Error creating connection $ConnectionName"
    }
    start-sleep -s 5
    try{
        Update-MgExternalConnectionSchema -ExternalConnectionId $ConnectionName -BodyParameter $schemaParams
    }
    catch{
        write-host "Error creating schema for connection $ConnectionName"
    }
}

Function Remove-ExternalConnection
{
    Param(
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipeLineByPropertyName = $true,
            ValueFromRemainingArguments = $false,
            Position = 0)]
        [ValidateNotNullOrEmpty()]
        [String]$ConnectionName
    )
    Remove-MgExternalConnection -ExternalConnectionId $ConnectionName
}

Function Write-Object
{
    param (
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipeLineByPropertyName = $true,
            ValueFromRemainingArguments = $false,
            Position = 0)]
        [string]$externalConnectionId,
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipeLineByPropertyName = $true,
            ValueFromRemainingArguments = $false,
            Position = 1)]
            [string]$item,
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipeLineByPropertyName = $true,
            ValueFromRemainingArguments = $false,
            Position = 2)]
            [string]$externalItemId,
        [Parameter(Mandatory = $true,
            ValueFromPipeline = $true,
            ValueFromPipeLineByPropertyName = $true,
            ValueFromRemainingArguments = $false,
            Position = 3)]
            [string]$content
    )

    $params = @{
        acl = @(
            @{
                type = "everyone"
                value = "everyone"
                accessType = "grant"
            }
        )
        properties = @{
            CompanyName = $item
        }
        content = @{
            value = $content
            type = "text"
        }
}

Set-MgExternalConnectionItem -ExternalConnectionId $externalConnectionId -ExternalItemId $externalItemId -BodyParameter $params
}



#Main Script
Connect-ToGraph

$ConnectionName = "PowerShellGraphConnector"

if ($Process.ToLower() -eq "install")
{
    New-ExternalConnection -ConnectionName $ConnectionName

}
elseif ($Process.ToLower() -eq "uninstall")
{
    Remove-ExternalConnection -ConnectionName $ConnectionName
}
elseif ($Process.ToLower() -eq "writeitems")
{
    Import-Csv -Path "C:\github\MyScripts\GraphAPI\PS_SDK\ExternalItems\fictitious_companies.csv" | ForEach-Object {
        Write-Object -externalConnectionId $ConnectionName -item $_.name -externalItemId $_.name -content $_.description
    } 
}
else
{
    Write-Host "Invalid process specified. Use 'install','uninstall', or 'writeitems."
}

#Disconnect from Graph when complete!
Disconnect-MGGraph
