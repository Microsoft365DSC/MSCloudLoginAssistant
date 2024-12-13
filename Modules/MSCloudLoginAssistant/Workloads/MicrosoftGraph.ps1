function Connect-MSCloudLoginMicrosoftGraph
{
    [CmdletBinding()]
    param()

    $ProgressPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginMicrosoftGraph'

    # If the current profile is not the same we expect, make the switch.
    if ($Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected)
    {
        if (($Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AuthenticationType -eq 'ServicePrincipalWithSecret' `
                    -or $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AuthenticationType -eq 'Identity') `
                -and (Get-Date -Date $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ConnectedDateTime) -lt [System.DateTime]::Now.AddMinutes(-50))
        {
            Add-MSCloudLoginAssistantEvent -Message 'Token is about to expire, renewing' -Source $source

            $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected = $false
        }
        elseif ($null -eq (Get-MgContext))
        {
            $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected = $false
        }
        else
        {
            return
        }
    }

    if ($Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AuthenticationType -eq 'CredentialsWithApplicationId' -or
        $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AuthenticationType -eq 'Credentials')
    {
        Add-MSCloudLoginAssistantEvent -Message 'Will try connecting with user credentials' -Source $source
        Connect-MSCloudLoginMSGraphWithUser
    }
    elseif ($Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AuthenticationType -eq 'CredentialsWithTenantId')
    {
        Add-MSCloudLoginAssistantEvent -Message 'Will try connecting with user credentials and Tenant Id' -Source $source
        Connect-MSCloudLoginMSGraphWithUser -TenantId $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.TenantId
    }
    elseif ($Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AuthenticationType -eq 'Identity')
    {
        Add-MSCloudLoginAssistantEvent -Message 'Connecting with managed identity' -Source $source

        $resourceEndpoint = ($Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ResourceUrl -split '/')[2]
        if ($env:AZUREPS_HOST_ENVIRONMENT -like 'AzureAutomation*')
        {
            $url = $env:IDENTITY_ENDPOINT
            $headers = New-Object 'System.Collections.Generic.Dictionary[[String],[String]]'
            $headers.Add('X-IDENTITY-HEADER', $env:IDENTITY_HEADER)
            $headers.Add('Metadata', 'True')
            $body = @{resource = "https://$resourceEndPoint/" }
            $oauth2 = Invoke-RestMethod $url -Method 'POST' -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body $body
            $accessToken = $oauth2.access_token
        }
        elseif ('http://localhost:40342' -eq $env:IMDS_ENDPOINT)
        {
            #Get endpoint for Azure Arc Connected Device
            $apiVersion = '2020-06-01'
            $resource = "https://$resourceEndpoint"
            $endpoint = '{0}?resource={1}&api-version={2}' -f $env:IDENTITY_ENDPOINT, $resource, $apiVersion
            $secretFile = ''
            try
            {
                Invoke-WebRequest -Method GET -Uri $endpoint -Headers @{Metadata = 'True' } -UseBasicParsing
            }
            catch
            {
                $wwwAuthHeader = $_.Exception.Response.Headers['WWW-Authenticate']
                if ($wwwAuthHeader -match 'Basic realm=.+')
                {
                    $secretFile = ($wwwAuthHeader -split 'Basic realm=')[1]
                }
            }
            $secret = Get-Content -Raw $secretFile
            $response = Invoke-WebRequest -Method GET -Uri $endpoint -Headers @{Metadata = 'True'; Authorization = "Basic $secret" } -UseBasicParsing
            if ($response)
            {
                $accessToken = (ConvertFrom-Json -InputObject $response.Content).access_token
            }
        }
        else
        {
            # Get correct endopint for AzureVM
            $oauth2 = Invoke-RestMethod -Uri "http://169.254.169.254/metadata/identity/oauth2/token?api-version=2018-02-01&resource=https%3A%2F%2F$($resourceEndpoint)%2F" -Headers @{Metadata = 'true' }
            $accessToken = $oauth2.access_token
        }

        $accessToken = $accessToken | ConvertTo-SecureString -AsPlainText -Force
        Connect-MgGraph -AccessToken $accessToken -Environment $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.GraphEnvironment
        $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ConnectedDateTime = [System.DateTime]::Now.ToString()
        $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.MultiFactorAuthentication = $false
        $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected = $true
        $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.TenantId = (Get-MgContext).TenantId
    }
    else
    {
        try
        {
            if ($Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AuthenticationType -eq 'ServicePrincipalWithThumbprint')
            {
                if ($null -ne $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Endpoints -and `
                    $null -ne $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Endpoints.ConnectionUri -and `
                    $null -ne $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Endpoints.AzureADAuthorizationEndpointUri)
                {
                    $accessToken = Get-MSCloudLoginAccessToken -ConnectionUri $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Endpoints.ConnectionUri `
                        -AzureADAuthorizationEndpointUri $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Endpoints.AzureADAuthorizationEndpointUri `
                        -ApplicationId $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ApplicationId `
                        -TenantId $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.TenantId `
                        -CertificateThumbprint $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.CertificateThumbprint
                    $accessToken = ConvertTo-SecureString $accessToken -AsPlainText -Force
                    Connect-MgGraph -AccessToken $accessToken
                    Add-MSCloudLoginAssistantEvent -Message 'Successfully connected to the Microsoft Graph API using Certificate Thumbprint' -Source $source
                }
                else
                {
                    Add-MSCloudLoginAssistantEvent -Message 'Connecting by Environment Name' -Source $source
                    try
                    {
                        Connect-MgGraph -ClientId $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ApplicationId `
                            -TenantId $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.TenantId `
                            -CertificateThumbprint $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.CertificateThumbprint `
                            -Environment $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.GraphEnvironment `
                            -ErrorAction Stop | Out-Null
                    }
                    catch
                    {
                        # Check into the localmachine store
                        $cert = Get-ChildItem "Cert:\LocalMachine\My\$($Script:MSCloudLoginConnectionProfile.MicrosoftGraph.CertificateThumbprint)"
                        Connect-MgGraph -ClientId $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ApplicationId `
                            -TenantId $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.TenantId `
                            -Environment $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.GraphEnvironment `
                            -Certificate $cert | Out-Null
                    }
                }

                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ConnectedDateTime = [System.DateTime]::Now.ToString()
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.MultiFactorAuthentication = $false
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected = $true
            }
            elseif ($Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AuthenticationType -eq 'ServicePrincipalWithSecret')
            {
                Add-MSCloudLoginAssistantEvent -Message 'Connecting to Microsoft Graph with ApplicationSecret' -Source $source
                $secStringPassword = ConvertTo-SecureString -String $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ApplicationSecret -AsPlainText -Force
                $userName = $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ApplicationId
                [pscredential]$credObject = New-Object System.Management.Automation.PSCredential ($userName, $secStringPassword)
                Connect-MgGraph -TenantId $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.TenantId `
                    -ClientSecretCredential $credObject | Out-Null
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ConnectedDateTime = [System.DateTime]::Now.ToString()
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.MultiFactorAuthentication = $false
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected = $true
            }
            elseif ($Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AuthenticationType -eq 'AccessTokens')
            {
                Add-MSCloudLoginAssistantEvent -Message 'Connecting to Microsoft Graph with AccessToken' -Source $source
                $secStringAccessToken = ConvertTo-SecureString -String $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AccessTokens[0] -AsPlainText -Force
                Connect-MgGraph -Environment $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.GraphEnvironment `
                    -AccessToken $secStringAccessToken | Out-Null
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ConnectedDateTime = [System.DateTime]::Now.ToString()
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.MultiFactorAuthentication = $false
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected = $true
            }
            Add-MSCloudLoginAssistantEvent -Message 'Connected' -Source $source
        }
        catch
        {
            Add-MSCloudLoginAssistantEvent -Message $_ -Source $source -EntryType 'Error'
            throw $_
        }
    }
}

function Connect-MSCloudLoginMSGraphWithUser
{
    [CmdletBinding()]
    param(
        [Parameter()]
        [System.String]
        $TenantId
    )

    $source = 'Connect-MSCloudLoginMSGraphWithUser'

    if ($Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Credentials.UserName -ne (Get-MgContext).Account)
    {
        Add-MSCloudLoginAssistantEvent -Message "The current account that is connect doesn't match the one we're trying to authenticate with. Disconnecting from Graph." -Source $source
        try
        {
            Disconnect-MgGraph -ErrorAction Stop | Out-Null
        }
        catch
        {
            Add-MSCloudLoginAssistantEvent -Message 'No connections to Microsoft Graph were found.' -Source $source
        }
    }

    if ([System.String]::IsNullOrEmpty($Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ApplicationId))
    {
        $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ApplicationId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'
    }

    $TenantId = $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Credentials.Username.Split('@')[1]
    $url = $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.TokenUrl
    $body = @{
        scope      = $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Scope
        grant_type = 'password'
        username   = $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Credentials.Username
        password   = $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Credentials.GetNetworkCredential().Password
        client_id  = $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ApplicationId
    }
    Add-MSCloudLoginAssistantEvent -Message 'Requesting Access Token for Microsoft Graph' -Source $source

    try
    {
        $OAuthReq = Invoke-RestMethod -Uri $url -Method Post -Body $body
        $AccessToken = ConvertTo-SecureString $OAuthReq.access_token -AsPlainText -Force

        Add-MSCloudLoginAssistantEvent -Message "Connecting to Microsoft Graph - Environment {$($Script:MSCloudLoginConnectionProfile.MicrosoftGraph.GraphEnvironment)}" -Source $source

        # Domain.Read.All permission Scope is required to get the domain name for the SPO Admin Center.
        if ([System.String]::IsNullOrEmpty($Script:MSCloudLoginConnectionProfile.MicrosoftGraph.TenantId))
        {
            Connect-MgGraph -AccessToken $AccessToken `
                -Environment $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.GraphEnvironment | Out-Null
        }
        else
        {
            Connect-MgGraph -AccessToken $AccessToken `
                -TenantId $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.TenantId `
                -Environment $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.GraphEnvironment | Out-Null
        }
        $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ConnectedDateTime = [System.DateTime]::Now.ToString()
        $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.MultiFactorAuthentication = $false
        $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected = $true
        $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AccessTokens = $AccessToken
    }
    catch
    {
        if ($_.Exception -like 'System.Net.WebException: The remote server returned an error: (400) Bad Request.*' -and `
            (Assert-IsNonInteractiveShell) -eq $true)
        {
            Write-Warning -Message "Unable to retrieve AccessToken. Have you registered the 'Microsoft Graph PowerShell' application already? Please run 'Connect-MgGraph -Scopes Domain.Read.All' and logon using '$($Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Credentials.Username)'"
            return
        }

        try
        {
            Add-MSCloudLoginAssistantEvent -Message 'Attempting to connect without specifying the Environment' -Source $source
            Connect-MgGraph -AccessToken $AccessToken | Out-Null
            $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ConnectedDateTime = [System.DateTime]::Now.ToString()
            $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.MultiFactorAuthentication = $false
            $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected = $true
            $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AccessTokens = $AccessToken
        }
        catch
        {
            Add-MSCloudLoginAssistantEvent -Message "Error connecting - $_" -Source $source -EntryType 'Error'
            Add-MSCloudLoginAssistantEvent -Message 'Connecting to Microsoft Graph interactively' -Source $source

            try
            {
                Connect-MgGraph -Environment $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.GraphEnvironment `
                    -TenantId $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.TenantId `
                    -ClientId $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ApplicationId `
                    -Scopes 'Domain.Read.All' -ErrorAction 'Stop' | Out-Null
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected = $true
            }
            catch
            {
                $err = $_
                if ($err -like '*\.graph\GraphContext.json*')
                {
                    $pathStart = $err.ToString().IndexOf("to file at '", 0) + 12
                    $pathEnd = $err.ToString().IndexOf("'", $pathStart)
                    $path = $err.ToString().Substring($pathStart, $pathEnd - $pathStart)

                    New-Item $path -Force | Out-Null
                    Connect-MgGraph -Environment $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.GraphEnvironment `
                        -TenantId $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.TenantId `
                        -ClientId $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ApplicationId `
                        -Scopes 'Domain.Read.All' | Out-Null
                    $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected = $true
                }

                if ($err.Exception.Message -eq 'Device code terminal timed-out after 120 seconds. Please try again.')
                {
                    throw 'Unable to connect to the Microsoft Graph. Please make sure the app permissions are setup correctly. Please run Update-M365DSCAllowedGraphScopes.'
                }
            }
        }
    }
}

function Disconnect-MSCloudLoginMicrosoftGraph
{
    [CmdletBinding()]
    param()

    $source = 'Disconnect-MSCloudLoginMicrosoftGraph'

    if ($Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected)
    {
        Add-MSCloudLoginAssistantEvent -Message 'Attempting to disconnect from Microsoft Graph' -Source $source
        Disconnect-MgGraph | Out-Null
        $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected = $false
        Add-MSCloudLoginAssistantEvent -Message 'Successfully disconnected from Microsoft Graph' -Source $source
    }
    else
    {
        Add-MSCloudLoginAssistantEvent -Message 'No connections to Microsoft Graph were found' -Source $source
    }
}
