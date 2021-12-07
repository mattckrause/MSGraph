# Message Purge Sample Script using Invoke-RestMethod for Graph API Calls

## Description

A sample PowerShell script I created to purge messages from mailboxes using Microsoft Graph API.
The sample reads in a list of mailboxes and messages from a .CSV file.

The script is expecting Application Permissions (Client Secret/App ID). So AAD app registrations should be done accordingly.

To archive messages add the archive switch and set to $true.

## Updates
6.23.20 - Created function to encode URL in search query. This solves a bug where internet message ids that included '+' would return a null value.

8.12.20 - Added functionality for archiving the purged messages. Controlable using the archive switch.

## Examples

Script with Archiving:
````PS C:\PurgeMessage.ps1 -Archive:$true````