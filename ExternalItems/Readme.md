# PowerShell Graph Connector sample

This is a sample PowerShell Graph connector script. I wrote it as a quick sample for what the steps are to create a custom Graph connector.

1. Create and external connection.
2. Create a Schema in that external connection.
3. Write objects to the external connection.

This is a very simple script. There is very little error checking and the process is completely manual.

You will need to have an app registration created with the below permissions in your environment:

Required permissions:
ExternalItem.ReadWrite.Ownedby
ExternalConnection.ReadWrite.Ownedby

The authentication uses certificate authentication. For testing, I used a self-signed cert. But that's not recommended for security purposes.
additionally, you can use a client secret if necessary. Update the Connect-tograph function to use the following Gist:

[Client Secret Auth](https://gist.github.com/mattckrause/61873597a90716a169961a37eee7728b#file-secretauth-ps1)

To run the script:

``` PowerShell
New-GraphConnector.ps1 -process Install #this will create the external connection and register the schema
New-GraphConnector.ps1 -process uninstall #this will remove the Graph connector if necessary
New-graphConnector.ps1 -process WriteItems #this will read the external items from the ficticious_companies.csv file and write them to the Graph connector.
```

I have also included each of the steps in it's own script file to run individually.

Create_ExternalConnection.ps1
Register-Schema.ps1
Write-Items.ps1
