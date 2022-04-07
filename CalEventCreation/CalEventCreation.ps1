<#
Create Calendar events using batching.
Getting token via Client Credentials Flow
#>

#Access Token Function
Function Get-AccessToken
{
    $AppId = '' # <- Application ID (AAD)
    $AppSecret = '' # <- App Registration Secret
    $TenantID = '' # <- Tenant ID
    $Scope = "https://graph.microsoft.com/.default" # <- Scope of the application
    $Url = "https://login.microsoftonline.com/$tenantID/oauth2/v2.0/token" # <- V2 Token Endpoint

    # Create body
    $Body = @{
        client_id     = $AppId
        client_secret = $AppSecret
        scope         = $Scope
        grant_type    = 'client_credentials'
    }

    # Splat the parameters for Invoke-Restmethod for cleaner code
    $PostSplat = @{
        ContentType = 'application/x-www-form-urlencoded'

        Method = 'POST'
        Body = $Body
        Uri = $Url
    }
    # Request the token!
    $Request = Invoke-RestMethod @PostSplat
    return $Request
}

Function SetEvent
{
    Param(
        [Parameter(Mandatory = $true,
            Position = 0)]
        [ValidateNotNullOrEmpty()]
        [String]$User,
        [Parameter(Mandatory = $true,
            Position = 1)]
        [ValidateNotNullOrEmpty()]
        [String]$Token
    )
    #batch URI
    $queryURI = "https://graph.microsoft.com/v1.0/`$batch"

    $bodyCreation = @{
        "requests" = @(
            @{
                "id"      = "1";
                "method"  = "POST";
                "url"     = "/users/$User/calendar/events";
                "body"    = @{
                    "subject"  = "New Years Day";
                    "body"     = @{
                        "contentType" = "HTML";
                        "content"     = "New Years"
                    };
                    "start"    = @{
                        "dateTime" = "2022-01-01T00:00:00";
                        "timeZone" = "America/Denver"
                    };
                    "end"      = @{
                        "dateTime" = "2022-01-02T00:00:00";
                        "timeZone" = "America/Denver"
                    };
                    "isAllDay" = $true
                };
                "headers" = @{
                    "Content-Type" = "application/json"
                }
            };
            @{
                "id"      = "2";
                "method"  = "POST";
                "url"     = "/users/$User/calendar/events";
                "body"    = @{
                    "subject"  = "Thanksgiving";
                    "body"     = @{
                        "contentType" = "HTML";
                        "content"     = "Thanksgiving"
                    };
                    "start"    = @{
                        "dateTime" = "2022-11-27T00:00:00";
                        "timeZone" = "America/Denver"
                    };
                    "end"      = @{
                        "dateTime" = "2022-11-28T00:00:00";
                        "timeZone" = "America/Denver"
                    };
                    "isAllDay" = $true
                };
                "headers" = @{
                    "Content-Type" = "application/json"
                }
            };
            @{
                "id"      = "3";
                "method"  = "POST";
                "url"     = "/users/$User/calendar/events";
                "body"    = @{
                    "subject"  = "Martin Luther King Jr. Day";
                    "body"     = @{
                        "contentType" = "HTML";
                        "content"     = "Martin Luther King Jr. Day"
                    };
                    "start"    = @{
                        "dateTime" = "2022-01-18T00:00:00";
                        "timeZone" = "America/Denver"
                    };
                    "end"      = @{
                        "dateTime" = "2022-01-19T00:00:00";
                        "timeZone" = "America/Denver"
                    };
                    "isAllDay" = $true
                };
                "headers" = @{
                    "Content-Type" = "application/json"
                }
            };
            @{
                "id"      = "4";
                "method"  = "POST";
                "url"     = "/users/$User/calendar/events";
                "body"    = @{
                    "subject"  = "Memorial Day";
                    "body"     = @{
                        "contentType" = "HTML";
                        "content"     = "Memorial Day"
                    };
                    "start"    = @{
                        "dateTime" = "2022-05-31T00:00:00";
                        "timeZone" = "America/Denver"
                    };
                    "end"      = @{
                        "dateTime" = "2022-06-1T00:00:00";
                        "timeZone" = "America/Denver"
                    };
                    "isAllDay" = $true
                };
                "headers" = @{
                    "Content-Type" = "application/json"
                }
            }
        )
    }

    $body = $bodyCreation | ConvertTo-Json -Depth 4

    $CreateResults = Invoke-RestMethod -Body $body -Uri $queryURI -Method "POST" -Headers @{Authorization = "Bearer $($Token)"; "Content-Type" = "application/json" }
    Return $CreateResults
}

#Main Script
$user = "user@domain.com" # <- Update this to the desired user.
#get access token
$TokenCache = Get-AccessToken
$Token = $TokenCache.access_token

#Create Appointments
SetEvent -User $user -Token $token
