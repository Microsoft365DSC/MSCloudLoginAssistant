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

    if ($Script:CustomEnvConfig.CustomEnvironment)
    {
        $customEnv = Get-MgEnvironment | Where-Object { $_.Name -eq 'Custom' }
        if ($null -eq $customEnv)
        {
            Add-MgEnvironment -Name 'Custom' -GraphEndpoint $Script:CustomEnvConfig.CustomGraphResourceUrl -AzureADEndPoint $Script:CustomEnvConfig.CustomGraphTokenUrl
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

        $resourceEndpoint = $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ResourceUrl.TrimEnd('/')
        $accessToken = Get-AuthToken -Resource $resourceEndpoint -Identity

        $accessToken = $accessToken | ConvertTo-SecureString -AsPlainText -Force
        Connect-MgGraph -AccessToken $accessToken -Environment $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.GraphEnvironment -NoWelcome
        $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.CompleteConnection()
        $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.TenantId = (Get-MgContext).TenantId
    }
    else
    {
        try
        {
            if ($Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AuthenticationType -eq 'ServicePrincipalWithThumbprint')
            {
                if ($Script:CustomEnvConfig.CustomEnvironment -and `
                    $null -ne $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Scope -and `
                    $null -ne $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.TokenUrl)
                {
                    $accessToken = Get-MSCloudLoginAccessToken -ConnectionUri $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Scope `
                        -AzureADAuthorizationEndpointUri $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.TokenUrl `
                        -ApplicationId $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ApplicationId `
                        -TenantId $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.TenantId `
                        -CertificateThumbprint $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.CertificateThumbprint
                    $accessToken = ConvertTo-SecureString $accessToken -AsPlainText -Force
                    Connect-MgGraph -AccessToken $accessToken -NoWelcome -Environment $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.GraphEnvironment
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
                            -NoWelcome `
                            -ErrorAction Stop
                    }
                    catch
                    {
                        # Check into the localmachine store
                        $cert = Get-ChildItem "Cert:\LocalMachine\My\$($Script:MSCloudLoginConnectionProfile.MicrosoftGraph.CertificateThumbprint)"
                        Connect-MgGraph -ClientId $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ApplicationId `
                            -TenantId $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.TenantId `
                            -Environment $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.GraphEnvironment `
                            -Certificate $cert `
                            -NoWelcome
                    }
                }

                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.CompleteConnection()
            }
            elseif ($Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AuthenticationType -eq 'ServicePrincipalWithSecret')
            {
                Add-MSCloudLoginAssistantEvent -Message 'Connecting to Microsoft Graph with ApplicationSecret' -Source $source
                $secStringPassword = ConvertTo-SecureString -String $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ApplicationSecret -AsPlainText -Force
                $userName = $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ApplicationId
                [pscredential]$credObject = New-Object System.Management.Automation.PSCredential ($userName, $secStringPassword)
                Connect-MgGraph -TenantId $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.TenantId `
                    -ClientSecretCredential $credObject `
                    -Environment $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.GraphEnvironment `
                    -NoWelcome
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.CompleteConnection()
            }
            elseif ($Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AuthenticationType -eq 'ServicePrincipalWithPath')
            {
                Add-MSCloudLoginAssistantEvent -Message 'Connecting to Microsoft Graph with Certificate Path' -Source $source
                $certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new((Resolve-Path $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.CertificatePath), $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.CertificatePassword, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::UserKeySet)
                Connect-MgGraph -TenantId $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.TenantId `
                    -ClientId $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ApplicationId `
                    -Certificate $certificate `
                    -Environment $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.GraphEnvironment `
                    -NoWelcome
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.CompleteConnection()
            }
            elseif ($Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AuthenticationType -eq 'AccessTokens')
            {
                Add-MSCloudLoginAssistantEvent -Message 'Connecting to Microsoft Graph with AccessToken' -Source $source
                $secStringAccessToken = ConvertTo-SecureString -String $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AccessTokens[0] -AsPlainText -Force
                Connect-MgGraph -Environment $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.GraphEnvironment `
                    -AccessToken $secStringAccessToken `
                    -NoWelcome
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.CompleteConnection()
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

    if ($Script:CustomEnvConfig.CustomEnvironment)
    {
        $customEnv = Get-MgEnvironment | Where-Object { $_.Name -eq 'Custom' }
        if ($null -eq $customEnv)
        {
            Add-MgEnvironment -Name 'Custom' -GraphEndpoint $Script:CustomEnvConfig.CustomGraphResourceUrl -AzureADEndPoint $Script:CustomEnvConfig.CustomGraphTokenUrl
        }
    }

    $source = 'Connect-MSCloudLoginMSGraphWithUser'

    if ($Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Credentials.UserName -ne (Get-MgContext).Account)
    {
        Add-MSCloudLoginAssistantEvent -Message "The currently connected account doesn't match the one we're trying to authenticate with. Disconnecting from Graph." -Source $source
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

    Add-MSCloudLoginAssistantEvent -Message 'Requesting Access Token for Microsoft Graph' -Source $source -Verbose

    try
    {
        try
        {
            $request = Get-AuthToken -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AuthorizationUrl `
                -Credential $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Credentials `
                -ClientId $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ApplicationId `
                -Scope $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Scope `
                -TenantId $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Credentials.Username.Split('@')[1]

            $AccessToken = ConvertTo-SecureString $request.access_token -AsPlainText -Force

            Add-MSCloudLoginAssistantEvent -Message "Connecting to Microsoft Graph - Environment {$($Script:MSCloudLoginConnectionProfile.MicrosoftGraph.GraphEnvironment)}" -Source $source

            # Domain.Read.All permission Scope is required to get the domain name for the SPO Admin Center.
            Connect-MgGraph -AccessToken $AccessToken `
                    -Environment $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.GraphEnvironment `
                    -NoWelcome
            $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.CompleteConnection()
            $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AccessTokens = $AccessToken
        }
        catch
        {
            if ($_.ErrorDetails.Message -like '*AADSTS50076*')
            {
                Add-MSCloudLoginAssistantEvent -Message 'Account used required MFA' -Source $source
                Connect-MSCloudLoginMSGraphWithUserMFA
            }
        }
    }
    catch
    {
        if ($_.Exception.Message -like 'System.Net.WebException: The remote server returned an error: (400) Bad Request.*' -and `
            (Assert-IsNonInteractiveShell) -eq $true)
        {
            Write-Warning -Message "Unable to retrieve AccessToken. Have you registered the 'Microsoft Graph PowerShell' application already? Please run 'Connect-MgGraph -Scopes Domain.Read.All' and logon using '$($Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Credentials.Username)'"
            return
        }

        try
        {
            Add-MSCloudLoginAssistantEvent -Message 'Attempting to connect without specifying the Environment' -Source $source
            Connect-MgGraph -AccessToken $AccessToken -NoWelcome
            $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.CompleteConnection()
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
                    -Scopes 'Domain.Read.All' -ErrorAction 'Stop' `
                    -NoWelcome
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.CompleteConnection()
            }
            catch
            {
                $err = $_
                if ($err.ToString() -like '*\.graph\GraphContext.json*')
                {
                    $pathStart = $err.ToString().IndexOf("to file at '", 0) + 12
                    $pathEnd = $err.ToString().IndexOf("'", $pathStart)
                    $path = $err.ToString().Substring($pathStart, $pathEnd - $pathStart)

                    New-Item $path -Force | Out-Null
                    Connect-MgGraph -Environment $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.GraphEnvironment `
                        -TenantId $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.TenantId `
                        -ClientId $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ApplicationId `
                        -Scopes 'Domain.Read.All' `
                        -NoWelcome
                    $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.CompleteConnection()
                }

                if ($err.Exception.Message -eq 'Device code terminal timed-out after 120 seconds. Please try again.')
                {
                    throw 'Unable to connect to the Microsoft Graph. Please make sure the app permissions are setup correctly. Please run Update-M365DSCAllowedGraphScopes.'
                }
            }
        }
    }
}

function Connect-MSCloudLoginMSGraphWithUserMFA
{
    [CmdletBinding()]
    param()

    $source = 'Connect-MSCloudLoginMSGraphWithUserMFA'
    if ([System.String]::IsNullOrEmpty($Script:MSCloudLoginConnectionProfile.MicrosoftGraph.TenantId))
    {
        $tenantId = $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Credentials.UserName.Split('@')[1]
    }
    else
    {
        $tenantId = $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.TenantId
    }

    Add-MSCloudLoginAssistantEvent -Message 'Getting access token from Microsoft Graph using device code' -Source $source

    $request = Get-AuthToken -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AuthorizationUrl `
        -Credentials $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Credentials `
        -TenantId $tenantId `
        -ClientId $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ApplicationId `
        -Scope $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Scope `
        -DeviceCode

    $AccessToken = ConvertTo-SecureString $request.access_token -AsPlainText -Force

    Add-MSCloudLoginAssistantEvent -Message "Connecting to Microsoft Graph with MFA - Environment {$($Script:MSCloudLoginConnectionProfile.MicrosoftGraph.GraphEnvironment)}" -Source $source
    Connect-MgGraph -AccessToken $AccessToken `
        -Environment $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.GraphEnvironment `
        -NoWelcome

    Add-MSCloudLoginAssistantEvent -Message 'Successfully connected to Microsoft Graph with MFA' -Source $source

    $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AccessToken = $AccessToken
    $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.CompleteConnection($true)
}

function Disconnect-MSCloudLoginMicrosoftGraph
{
    [CmdletBinding()]
    param()

    $source = 'Disconnect-MSCloudLoginMicrosoftGraph'

    if ($Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected)
    {
        Add-MSCloudLoginAssistantEvent -Message 'Attempting to disconnect from Microsoft Graph' -Source $source
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected = $false
        Add-MSCloudLoginAssistantEvent -Message 'Successfully disconnected from Microsoft Graph' -Source $source
    }
    else
    {
        Add-MSCloudLoginAssistantEvent -Message 'No connections to Microsoft Graph were found' -Source $source
    }
}
