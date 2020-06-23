# Message Purge Sample Script

## Description

A sample PowerShell script I created to purge messages from mailboxes using Microsoft Graph API.  
The sample reads in a list of mailboxes and messages from a .CSV file.

The script is expecting Application Permissions (Client Secret/App ID). So AAD app registrations should be done accordingly.

## Updates
6.23.20 - Created function to encode URL in search query. This solves a bug where internet message ids that included '+' would return a null value.