# Message Purge Sample Script using Microsoft Graph PowerShell SDK

## Description

A PowerShell script I created to purge messages from mailboxes using Microsoft Graph PowerShell SDK.
This is the same functionality as the RESTful request script, just using the PowerShell SDK.

The script reads in a list of mailboxes and messages from a .CSV file.

The script is using Application Permissions and a certificate for authentication. Your AAD app registrations should be done accordingly.

To archive messages add the archive switch and set to $true.

## Updates


## Examples

Script with Archiving:
````PS C:\PurgeMessage_SDK.ps1 -Archive:$true````
