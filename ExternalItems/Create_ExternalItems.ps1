<#
Very simple script to use MS PowerShell Graph API SDK to create external connections and write objects to them. 
It uses dummy data created from copilot to populate the external items.
#>

#connect to graph
Function psmConnectToGraph
{
    $data = get-content -Path .env
    $appID = ($data[0].split("="))[1]
    $tenantID = ($data[1].split("="))[1]
    $authCertThumb = ($data[2].split("="))[1]

    Connect-MGGraph -ClientId $appID -TenantId $tenantID -CertificateThumbprint $authCertThumb -nowelcome
}

Function Create-ExternalConnection
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
        description = "Test connector called $ConnectionName. Containing a list of company names."
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
    start-sleep -s 20
    try{
        Update-MgExternalConnectionSchema -ExternalConnectionId $ConnectionName -BodyParameter $schemaParams
    }
    catch{
        write-host "Error creating schema for connection $ConnectionName"
    }
}

Function delete-ExternalConnection
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
#Invoke-MgGraphRequest -Method put -Uri "https://graph.microsoft.com/v1.0/external/connections/$externalConnectionId/items/$externalItemId" -Body @params
}



#Main Script
#Connect to Graph API
psmConnectToGraph

#Create External Connection and schema manually adjusting for smaller batches
for ($i = 1; $i -lt 26; $i++) {
    #create external item and schema
    Create-ExternalConnection -ConnectionName "Connection$i"
    #delete-ExternalConnection -ConnectionName "Connection$i"
    #sleep
    start-sleep -s 20
    }
#>

#write objects to external connections after creation manually adjusting for batches of 25
#<#
$objects = Import-Csv -Path "C:\github\MyScripts\GraphAPI\PS_SDK\ExternalItems\companies.csv" 
$content = Import-Csv -Path "C:\github\MyScripts\GraphAPI\PS_SDK\ExternalItems\content.csv"
$l = 0
#start with 132 for data ingestion
for($j = 1; $j -lt 26; $j++)
    {
        write-host "Writing objects to Connection$j"
        for ($k = 0; $k -lt 5; $k += 5)
            { $data = $objects[$l..($l+4)]
                $id = 1
                $data | ForEach-Object {
                    $index = get-random -Minimum 0 -Maximum 49
                    Write-Object -externalConnectionId "Connection$j" -item $_.name -externalItemId $id -content $content[$index].description
                    $id++
                }
                $l += 5
            }
            start-sleep -Seconds 10
    }
write-host "start next write object run at $l" #manually update $l with this value
#>

#Disconnect from Graph when complete!
Disconnect-MGGraph
