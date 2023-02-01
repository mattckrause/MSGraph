# Get-MessageCount.ps1

## Description

This is a script that uses Microsoft Graph API to report the number of messages in a group of mailboxes. Pass in a group name, and the start/end dates for collection.
 It uses the Microsoft Graph PowerShell SDK (https://learn.microsoft.com/en-us/powershell/microsoftgraph/installation?view=graph-powershell-1.0)

## Syntax for use

```powershell
Get-MessageCount.ps1 -groupName "aadGroupName" -startDate "2023-01-01" -endDate "2023-02-01"
```

## Requirements for use

1. Application Registraion in Azure AD configured with certificate authentication and the following Application permissions for Microsoft Graph:
   1. Group.Read.All
   2. Mail.ReadBasic
2. Update the Get-MessageCount.ps1 script to include:
   1. Location of .CSV file report (line 17)
   2. Your certificate thumbprint for authentication (line 25)
   3. Your tenant ID (line 26)
   4. Your Application ID (line 27)

## Output

The script will output the total count of messages in the console. It also creates a .CSV file with a message breakdown per mailbox
[CSV Example File](https://github.com/mattckrause/MSGraph/blob/Main/MessageCount/messages.csv)

## Disclaimer

This is a work in progress
