function Connect-MSCloudLoginPnP
{
    [CmdletBinding()]
    param(
        [boolean]
        $ForceRefreshConnection = $false
    )

    $ProgressPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginPnP'

    if ($Script:MSCloudLoginConnectionProfile.PnP.Connected)
    {
        Add-MSCloudLoginAssistantEvent -Message 'Already connected to PnP, not attempting to authenticate.' -Source $source
        return
    }

    # Check if Graph-module is loaded and, if not, explicitly load before PnP
    # Workaround to fix: https://github.com/microsoft/Microsoft365DSC/issues/4746
    if (-not (Get-Module Microsoft.Graph.Authentication -ErrorAction SilentlyContinue))
    {
        Add-MSCloudLoginAssistantEvent -Message 'Explicit import of PS-module Microsoft.Graph.Authentication' -Source $source
        Import-Module Microsoft.Graph.Authentication -ErrorAction SilentlyContinue
    }

    $requiresWindowsPowerShell = $false
    if ($psversiontable.PSVersion.Major -ge 7)
    {
        try
        {
            Get-PnPAlert -ErrorAction 'Stop' | Out-Null
            Add-MSCloudLoginAssistantEvent -Message 'Retrieved results from the command. Not re-connecting to PnP.' -Source $source
            $Script:MSCloudLoginConnectionProfile.PnP.Connected = $true
            return
        }
        catch
        {
            Add-MSCloudLoginAssistantEvent -Message "Couldn't get results back from the command" -Source $source -EntryType 'Warning'
            Add-MSCloudLoginAssistantEvent -Message 'Using PowerShell 7 or above. Loading the PnP.PowerShell module using Windows PowerShell.' -Source $source
            try
            {
                $currentLoadedModule = Get-Module PnP.PowerShell
                if ($null -eq $currentLoadedModule)
                {
                    Import-Module PnP.PowerShell -UseWindowsPowerShell -Global -Force -ErrorAction Stop | Out-Null
                }
            }
            catch
            {
                $requiresWindowsPowerShell = $true
            }
        }
    }

    if ($requiresWindowsPowerShell)
    {
        throw "Powershell 7+ was detected. We need to load the PnP.PowerShell module using the -UseWindowsPowerShell switch which requires the module to be installed under C:\Program Files\WindowsPowerShell\Modules. You can either move the module to that location or use PowerShell 5.1 to install the modules using 'Install-Module Pnp.PowerShell -Force -Scope AllUsers'."
    }

    if ([string]::IsNullOrEmpty($Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl))
    {
        if (-not [string]::IsNullOrEmpty($Script:MSCloudLoginConnectionProfile.PnP.AdminUrl))
        {
            $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl = $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl
        }
        else
        {
            if ($Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType -eq 'Credentials' -and `
                    -not $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl)
            {
                $adminUrl = Get-SPOAdminUrl -Credential $Script:MSCloudLoginConnectionProfile.PnP.Credentials
                if ([String]::IsNullOrEmpty($adminUrl) -eq $false)
                {
                    $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl = $adminUrl
                    $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl = $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl
                }
                else
                {
                    throw 'Unable to retrieve SharePoint Admin Url. Check if the Graph can be contacted successfully.'
                }
            }
            else
            {
                if ($Script:MSCloudLoginConnectionProfile.PnP.TenantId.Contains('onmicrosoft'))
                {
                    $domain = $Script:MSCloudLoginConnectionProfile.PnP.TenantId.Replace('.onmicrosoft.', '-admin.sharepoint.')
                    if (-not $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl)
                    {
                        $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl = "https://$domain"
                    }
                    $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl = ("https://$domain").Replace('-admin', '')
                }
                elseif ($Script:MSCloudLoginConnectionProfile.PnP.TenantId.Contains('.onmschina.'))
                {
                    $domain = $Script:MSCloudLoginConnectionProfile.PnP.TenantId.Replace('.partner.onmschina.', '-admin.sharepoint.')
                    if (-not $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl)
                    {
                        $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl = "https://$domain"
                    }
                    $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl = ("https://$domain").Replace('-admin', '')
                }
                else
                {
                    throw 'TenantId must be in format contoso.onmicrosoft.com'
                }
            }
        }
    }
    elseif ([string]::IsNullOrEmpty($Script:MSCloudLoginConnectionProfile.PnP.AdminUrl))
    {
        $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl = $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl
    }

    try
    {
        if (-not $Script:MSCloudLoginConnectionProfile.PnP.Connected)
        {
            if ($Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType -eq 'ServicePrincipalWithThumbprint')
            {
                if ($Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl)
                {
                    if ($null -ne $Script:MSCloudLoginConnectionProfile.PnP.Endpoints -and `
                        $null -ne $Script:MSCloudLoginConnectionProfile.PnP.Endpoints.ConnectionUri -and `
                        $null -ne $Script:MSCloudLoginConnectionProfile.PnP.Endpoints.AzureADAuthorizationEndpointUri)
                    {
                        $accessToken = Get-MSCloudLoginAccessToken -ConnectionUri $Script:MSCloudLoginConnectionProfile.PnP.Endpoints.ConnectionUri `
                            -AzureADAuthorizationEndpointUri $Script:MSCloudLoginConnectionProfile.PnP.Endpoints.AzureADAuthorizationEndpointUri `
                            -ApplicationId $Script:MSCloudLoginConnectionProfile.PnP.ApplicationId `
                            -TenantId $Script:MSCloudLoginConnectionProfile.PnP.TenantId `
                            -CertificateThumbprint $Script:MSCloudLoginConnectionProfile.PnP.CertificateThumbprint
                        $Script:MSCloudLoginConnectionProfile.PnP.AccessTokens += $accessToken

                        Add-MSCloudLoginAssistantEvent -Message 'Connecting with Service Principal - Thumbprint' -Source $source
                        Add-MSCloudLoginAssistantEvent -Message "URL: $($Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl)" -Source $source
                        Add-MSCloudLoginAssistantEvent -Message "ConnectionUrl: $($Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl)" -Source $source
                        Connect-PnPOnline -Url $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl `
                            -AccessToken $accessToken | Out-Null
                    }
                    else
                    {
                        Add-MSCloudLoginAssistantEvent -Message 'Connecting with Service Principal - Thumbprint' -Source $source
                        Add-MSCloudLoginAssistantEvent -Message "URL: $($Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl)" -Source $source
                        Add-MSCloudLoginAssistantEvent -Message "ConnectionUrl: $($Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl)" -Source $source

                        if ($Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment -ne 'Custom')
                        {
                            Connect-PnPOnline -Url $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl `
                                -ClientId $Script:MSCloudLoginConnectionProfile.PnP.ApplicationId `
                                -Tenant $Script:MSCloudLoginConnectionProfile.PnP.TenantId `
                                -Thumbprint $Script:MSCloudLoginConnectionProfile.PnP.CertificateThumbprint `
                                -AzureEnvironment $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment | Out-Null
                        }
                        else
                        {
                            Connect-PnPOnline -Url $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl `
                                -ClientId $Script:MSCloudLoginConnectionProfile.PnP.ApplicationId `
                                -Tenant $Script:MSCloudLoginConnectionProfile.PnP.TenantId `
                                -Thumbprint $Script:MSCloudLoginConnectionProfile.PnP.CertificateThumbprint `
                                -AzureEnvironment $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment `
                                -AzureADLoginEndPoint $Script:MSCloudLoginConnectionProfile.PnP.EndPoints.AzureADLoginEndPoint `
                                -MicrosoftGraphEndPoint $Script:MSCloudLoginConnectionProfile.PnP.EndPoints.MicrosoftGraphEndPoint | Out-Null
                        }
                    }
                }
                elseif ($Script:MSCloudLoginConnectionProfile.PnP.AdminUrl)
                {
                    Add-MSCloudLoginAssistantEvent -Message 'Connecting with Service Principal - Thumbprint' -Source $source
                    Add-MSCloudLoginAssistantEvent -Message "URL: $($Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl)" -Source $source
                    Add-MSCloudLoginAssistantEvent -Message "AdminUrl: $($Script:MSCloudLoginConnectionProfile.PnP.AdminUrl)" -Source $source

                    $tenantIdValue = $Script:MSCloudLoginConnectionProfile.PnP.TenantId
                    if ($Script:MSCloudLoginConnectionProfile.PnP.EnvironmentName -eq 'AzureChinaCloud')
                    {
                        $tenantIdValue = $Script:MSCloudLoginConnectionProfile.PnP.TenantGUID
                    }

                    if ($Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment -ne 'Custom')
                    {
                        Connect-PnPOnline -Url $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl `
                            -ClientId $Script:MSCloudLoginConnectionProfile.PnP.ApplicationId `
                            -Tenant $tenantIdValue `
                            -Thumbprint $Script:MSCloudLoginConnectionProfile.PnP.CertificateThumbprint `
                            -AzureEnvironment $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment | Out-Null
                    }
                    else
                    {
                        Connect-PnPOnline -Url $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl `
                            -ClientId $Script:MSCloudLoginConnectionProfile.PnP.ApplicationId `
                            -Tenant $Script:MSCloudLoginConnectionProfile.PnP.TenantId `
                            -Thumbprint $Script:MSCloudLoginConnectionProfile.PnP.CertificateThumbprint `
                            -AzureEnvironment $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment `
                            -AzureADLoginEndPoint $Script:MSCloudLoginConnectionProfile.PnP.AzureADLoginEndPoint `
                            -MicrosoftGraphEndPoint $Script:MSCloudLoginConnectionProfile.PnP.MicrosoftGraphEndPoint | Out-Null
                    }
                }

                $Script:MSCloudLoginConnectionProfile.PnP.ConnectedDateTime = [System.DateTime]::Now.ToString()
                $Script:MSCloudLoginConnectionProfile.PnP.MultiFactorAuthentication = $false
                $Script:MSCloudLoginConnectionProfile.PnP.Connected = $true
            }
            elseif ($Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType -eq 'ServicePrincipalWithPath')
            {
                if ($Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl)
                {
                    Add-MSCloudLoginAssistantEvent -Message 'Connecting with Service Principal - Path' -Source $source
                    Add-MSCloudLoginAssistantEvent -Message "URL: $($Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl)" -Source $source
                    Add-MSCloudLoginAssistantEvent -Message "ConnectionUrl: $($Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl)" -Source $source
                    Connect-PnPOnline -Url $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl `
                        -ClientId $Script:MSCloudLoginConnectionProfile.PnP.ApplicationId `
                        -Tenant $Script:MSCloudLoginConnectionProfile.PnP.TenantId `
                        -CertificatePassword $Script:MSCloudLoginConnectionProfile.PnP.CertificatePassword `
                        -CertificatePath $Script:MSCloudLoginConnectionProfile.PnP.CertificatePath `
                        -AzureEnvironment $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment
                }
                else
                {
                    Add-MSCloudLoginAssistantEvent -Message 'Connecting with Service Principal - Path' -Source $source
                    Add-MSCloudLoginAssistantEvent -Message "URL: $($Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl)" -Source $source
                    Add-MSCloudLoginAssistantEvent -Message "AdminUrl: $($Script:MSCloudLoginConnectionProfile.PnP.AdminUrl)" -Source $source
                    Connect-PnPOnline -Url $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl `
                        -ClientId $Script:MSCloudLoginConnectionProfile.PnP.ApplicationId `
                        -Tenant $Script:MSCloudLoginConnectionProfile.PnP.TenantId `
                        -CertificatePassword $Script:MSCloudLoginConnectionProfile.PnP.CertificatePassword `
                        -CertificatePath $Script:MSCloudLoginConnectionProfile.PnP.CertificatePath `
                        -AzureEnvironment $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment
                }

                $Script:MSCloudLoginConnectionProfile.PnP.ConnectedDateTime = [System.DateTime]::Now.ToString()
                $Script:MSCloudLoginConnectionProfile.PnP.MultiFactorAuthentication = $false
                $Script:MSCloudLoginConnectionProfile.PnP.Connected = $true
            }
            elseif ($Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType -eq 'ServicePrincipalWithSecret')
            {
                if ($Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl -or $ForceRefreshConnection)
                {
                    Add-MSCloudLoginAssistantEvent -Message 'Connecting with Service Principal - Secret' -Source $source
                    Add-MSCloudLoginAssistantEvent -Message "URL: $($Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl)" -Source $source
                    Add-MSCloudLoginAssistantEvent -Message "ConnectionUrl: $($Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl)" -Source $source
                    Connect-PnPOnline -Url $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl `
                        -ClientId $Script:MSCloudLoginConnectionProfile.PnP.ApplicationId `
                        -ClientSecret $Script:MSCloudLoginConnectionProfile.PnP.ApplicationSecret `
                        -AzureEnvironment $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment `
                        -WarningAction 'Ignore'
                }
                else
                {
                    Add-MSCloudLoginAssistantEvent -Message 'Connecting with Service Principal - Secret' -Source $source
                    Add-MSCloudLoginAssistantEvent -Message "URL: $($Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl)" -Source $source
                    Add-MSCloudLoginAssistantEvent -Message "AdminUrl: $($Script:MSCloudLoginConnectionProfile.PnP.AdminUrl)" -Source $source
                    Connect-PnPOnline -Url $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl `
                        -ClientId $Script:MSCloudLoginConnectionProfile.PnP.ApplicationId `
                        -ClientSecret $Script:MSCloudLoginConnectionProfile.PnP.ApplicationSecret `
                        -AzureEnvironment $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment `
                        -WarningAction 'Ignore'
                }
                $Script:MSCloudLoginConnectionProfile.PnP.ConnectedDateTime = [System.DateTime]::Now.ToString()
                $Script:MSCloudLoginConnectionProfile.PnP.MultiFactorAuthentication = $false
                $Script:MSCloudLoginConnectionProfile.PnP.Connected = $true
            }
            elseif ($Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType -eq 'CredentialsWithTenantId')
            {
                throw 'You cannot specify TenantId with Credentials when connecting to PnP.'
            }
            elseif ($Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType -eq 'Credentials')
            {
                if ($Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl -or $ForceRefreshConnection)
                {
                    Add-MSCloudLoginAssistantEvent -Message 'Connecting with Credentials' -Source $source
                    Add-MSCloudLoginAssistantEvent -Message "URL: $($Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl)" -Source $source
                    Add-MSCloudLoginAssistantEvent -Message "ConnectionUrl: $($Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl)" -Source $source
                    Connect-PnPOnline -Url $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl `
                        -Credentials $Script:MSCloudLoginConnectionProfile.PnP.Credentials `
                        -AzureEnvironment $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment
                }
                else
                {
                    Add-MSCloudLoginAssistantEvent -Message 'Connecting with Credentials' -Source $source
                    Add-MSCloudLoginAssistantEvent -Message "URL: $($Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl)" -Source $source
                    Add-MSCloudLoginAssistantEvent -Message "AdminUrl: $($Script:MSCloudLoginConnectionProfile.PnP.AdminUrl)" -Source $source
                    Connect-PnPOnline -Url $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl `
                        -Credentials $Script:MSCloudLoginConnectionProfile.PnP.Credentials `
                        -AzureEnvironment $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment
                }

                $Script:MSCloudLoginConnectionProfile.PnP.ConnectedDateTime = [System.DateTime]::Now.ToString()
                $Script:MSCloudLoginConnectionProfile.PnP.MultiFactorAuthentication = $false
                $Script:MSCloudLoginConnectionProfile.PnP.Connected = $true
            }
            elseif ($Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType -eq 'Identity')
            {
                if ($Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl)
                {
                    $connectionURL = $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl
                }
                else
                {
                    $connectionURL = $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl
                }

                if ('AzureAutomation/' -eq $env:AZUREPS_HOST_ENVIRONMENT)
                {
                    $url = $env:IDENTITY_ENDPOINT
                    $headers = New-Object 'System.Collections.Generic.Dictionary[[String],[String]]'
                    $headers.Add('X-IDENTITY-HEADER', $env:IDENTITY_HEADER)
                    $headers.Add('Metadata', 'True')
                    $body = @{resource = $connectionURL }
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
                    $oauth2 = Invoke-RestMethod -Uri "http://169.254.169.254/metadata/identity/oauth2/token?api-version=2018-02-01&resource=$ConnectionURL" -Headers @{Metadata = 'true' }
                    $accessToken = $oauth2.access_token

                }

                Connect-PnPOnline -Url $connectionURL `
                    -AccessToken $accessToken `
                    -AzureEnvironment $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment `
                    -WarningAction 'Ignore'

                $Script:MSCloudLoginConnectionProfile.PnP.ConnectedDateTime = [System.DateTime]::Now.ToString()
                $Script:MSCloudLoginConnectionProfile.PnP.MultiFactorAuthentication = $false
                $Script:MSCloudLoginConnectionProfile.PnP.Connected = $true
            }
            elseif ($Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType -eq 'AccessToken')
            {
                $Ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToCoTaskMemUnicode($Script:MSCloudLoginConnectionProfile.PnP.AccessTokens[0])
                $AccessTokenValue = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($Ptr)
                [System.Runtime.InteropServices.Marshal]::ZeroFreeCoTaskMemUnicode($Ptr)
                if ($Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl -or $ForceRefreshConnection)
                {
                    Add-MSCloudLoginAssistantEvent -Message 'Connecting with AccessToken' -Source $source
                    Add-MSCloudLoginAssistantEvent -Message "URL: $($Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl)" -Source $source
                    Add-MSCloudLoginAssistantEvent -Message "ConnectionUrl: $($Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl)" -Source $source
                    Connect-PnPOnline -Url $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl `
                        -AccessToken $AccessTokenValue `
                        -AzureEnvironment $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment
                }
                else
                {
                    Add-MSCloudLoginAssistantEvent -Message 'Connecting with AccessToken' -Source $source
                    Add-MSCloudLoginAssistantEvent -Message "URL: $($Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl)" -Source $source
                    Add-MSCloudLoginAssistantEvent -Message "AdminUrl: $($Script:MSCloudLoginConnectionProfile.PnP.AdminUrl)" -Source $source
                    Connect-PnPOnline -Url $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl `
                        -AccessToken $AccessTokenValue `
                        -AzureEnvironment $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment
                }

                $Script:MSCloudLoginConnectionProfile.PnP.ConnectedDateTime = [System.DateTime]::Now.ToString()
                $Script:MSCloudLoginConnectionProfile.PnP.MultiFactorAuthentication = $false
                $Script:MSCloudLoginConnectionProfile.PnP.Connected = $true
            }
        }
    }
    catch
    {
        if ($_.Exception -like '*AADSTS50076*')
        {
            try
            {
                Connect-PnPOnline -Url $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl `
                    -Interactive
                $Script:MSCloudLoginConnectionProfile.PnP.ConnectedDateTime = [System.DateTime]::Now.ToString()
                $Script:MSCloudLoginConnectionProfile.PnP.MultiFactorAuthentication = $true
                $Script:MSCloudLoginConnectionProfile.PnP.Connected = $true
            }
            catch
            {
                try
                {
                    Connect-PnPOnline -Url $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl -UseWebLogin
                    $Script:MSCloudLoginConnectionProfile.PnP.ConnectedDateTime = [System.DateTime]::Now.ToString()
                    $Script:MSCloudLoginConnectionProfile.PnP.MultiFactorAuthentication = $true
                    $Script:MSCloudLoginConnectionProfile.PnP.Connected = $true
                }
                catch
                {
                    $Script:MSCloudLoginConnectionProfile.PnP.Connected = $false
                    throw $_
                }
            }
        }
        elseif ($_.Exception -like '*The sign-in name or password does not match one in the Microsoft account system*')
        {
            # This error means that the account was trying to connect using MFA.
            try
            {
                Add-MSCloudLoginAssistantEvent 'Trying to acquire AccessToken' -Source $source
                $AuthHeader = Get-AuthHeader -UserPrincipalName $Script:MSCloudLoginConnectionProfile.PnP.Credentials.UserName `
                    -ResourceURI $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl `
                    -clientId $Script:MSCloudLoginConnectionProfile.PnP.ClientId `
                    -RedirectURI $Script:MSCloudLoginConnectionProfile.PnP.RedirectURI
                $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl.AccessToken = $AuthHeader.split(' ')[1]

                Add-MSCloudLoginAssistantEvent "Access Token = $($Script:MSCloudLoginConnectionProfile.PnP.AccessToken)" -Source $source
                if ($null -ne $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl.AccessToken)
                {
                    if ($Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl)
                    {
                        Connect-PnPOnline -Url $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl `
                            -AccessToken $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl.AccessToken
                    }
                    else
                    {
                        Connect-PnPOnline -Url $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl `
                            -AccessToken $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl.AccessToken
                    }
                }
                else
                {
                    if ($Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl)
                    {
                        Connect-PnPOnline -Url $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl `
                            -Interactive
                    }
                    else
                    {
                        Connect-PnPOnline -Url $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl `
                            -Interactive
                    }
                }
                $Script:MSCloudLoginConnectionProfile.PnP.ConnectedDateTime = [System.DateTime]::Now.ToString()
                $Script:MSCloudLoginConnectionProfile.PnP.MultiFactorAuthentication = $true
                $Script:MSCloudLoginConnectionProfile.PnP.Connected = $true
            }
            catch
            {
                Add-MSCloudLoginAssistantEvent "Error acquiring AccessToken: $($_.Exception.Message)" -Source $source -EntryType 'Error'
                try
                {
                    Connect-PnPOnline -Url $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl `
                        -Interactive
                    $Script:MSCloudLoginConnectionProfile.PnP.ConnectedDateTime = [System.DateTime]::Now.ToString()
                    $Script:MSCloudLoginConnectionProfile.PnP.MultiFactorAuthentication = $true
                    $Script:MSCloudLoginConnectionProfile.PnP.Connected = $true
                }
                catch
                {
                    $Script:MSCloudLoginConnectionProfile.PnP.Connected = $false
                    throw $_
                }
            }
        }
        elseif ($_.Exception -like '*AADSTS65001: The user or administrator has not consented to use the application with ID*')
        {
            try
            {
                Register-PnPManagementShellAccess
                Connect-PnPOnline -UseWebLogin -Url $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl
                $Script:MSCloudLoginConnectionProfile.PnP.Connected = $true
                $Script:MSCloudLoginConnectionProfile.PnP.ConnectedDateTime = [System.DateTime]::Now.ToString()
            }
            catch
            {
                throw "The PnP.PowerShell Azure AD Application has not been granted access for this tenant. Please run 'Register-PnPManagementShellAccess' to grant access and try again after."
            }
        }
        else
        {
            $Script:MSCloudLoginConnectionProfile.PnP.connected = $false

            $message = "An error has occurred $($_.Exception.Message)"
            throw $message
        }
    }
    return
}

function Disconnect-MSCloudLoginPnP
{
    [CmdletBinding()]
    param()

    $source = 'Disconnect-MSCloudLoginPnP'

    if ($Script:MSCloudLoginConnectionProfile.PnP.Connected)
    {
        Add-MSCloudLoginAssistantEvent -Message 'Attempting to disconnect from PnP' -Source $source
        Disconnect-PnPOnline | Out-Null
        $Script:MSCloudLoginConnectionProfile.PnP.Connected = $false
        Add-MSCloudLoginAssistantEvent -Message 'Successfully disconnected from PnP' -Source $source
    }
    else
    {
        Add-MSCloudLoginAssistantEvent -Message 'No connections to PnP were found' -Source $source
    }
}
