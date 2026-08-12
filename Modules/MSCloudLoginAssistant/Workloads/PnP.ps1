function Connect-MSCloudLoginPnP
{
    [CmdletBinding()]
    param(
        [boolean]
        $ForceRefreshConnection = $false
    )

    $ProgressPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginPnP'

    if (Test-MSCloudLoginConnectionReusable -WorkloadProfile $Script:MSCloudLoginConnectionProfile.PnP -Source $source)
    {
        Add-MSCloudLoginAssistantEvent -Message 'Already connected to PnP, not attempting to authenticate.' -Source $source
        return
    }

    # Check if Graph-module is loaded and, if not, explicitly load before PnP
    # Workaround to fix: https://github.com/microsoft/Microsoft365DSC/issues/4746
    if (-not (Get-Module Microsoft.Graph.Authentication -ErrorAction SilentlyContinue))
    {
        Add-MSCloudLoginAssistantEvent -Message 'Explicit import of PS-module Microsoft.Graph.Authentication' -Source $source
        try
        {
            Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
        }
        catch
        {
            Add-MSCloudLoginAssistantEvent -Message "Failed to import Microsoft.Graph.Authentication (workaround for Microsoft365DSC#4746 not applied): $($_.Exception.Message)" -Source $source -EntryType 'Warning'
        }
    }

    $currentLoadedModule = Get-Module PnP.PowerShell
    if ($null -eq $currentLoadedModule)
    {
        $availablePnPModule = Get-Module -Name PnP.PowerShell -ListAvailable | Sort-Object -Property Version -Descending | Select-Object -First 1
        if ($PSEdition -ne 'Desktop' -and $IsWindows -and $availablePnPModule.Version.Major -eq 1)
        {
            Add-MSCloudLoginAssistantEvent -Message 'Using PowerShell Core on Windows.' -Source $source
            try
            {
                Add-MSCloudLoginAssistantEvent -Message 'Loading the PnP.PowerShell module using Windows PowerShell.' -Source $source
                $pnpModule = Get-Module -Name PnP.PowerShell -ListAvailable | Where-Object CompatiblePSEditions -Contains 'Desktop' | Sort-Object -Property Version -Descending | Select-Object -First 1
                if ($null -eq $pnpModule)
                {
                    throw 'PnP.PowerShell module is not installed for Windows PowerShell. Please install the module using PowerShell 5.1 and try again.'
                }
                Import-Module -Name PnP.PowerShell -RequiredVersion $pnpModule.Version -UseWindowsPowerShell -Global -Force -ErrorAction Stop | Out-Null
            }
            catch
            {
                throw "Powershell 7+ was detected. We need to load the PnP.PowerShell module using the -UseWindowsPowerShell switch which requires the module to be installed under C:\Program Files\WindowsPowerShell\Modules. You can either move the module to that location or use PowerShell 5.1 to install the modules using 'Install-Module Pnp.PowerShell -Force -Scope AllUsers'. Original error: $($_.Exception.Message)"
            }
        }
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
                    throw 'Unable to retrieve SharePoint Admin Url. Check if Microsoft Graph can be contacted successfully.'
                }
            }
            else
            {
                $spoUrls = Get-MSCloudLoginSPOUrlFromTenantId -TenantId $Script:MSCloudLoginConnectionProfile.PnP.TenantId `
                    -EnvironmentName $Script:MSCloudLoginConnectionProfile.PnP.EnvironmentName
                if (-not $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl)
                {
                    $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl = $spoUrls.AdminUrl
                }
                $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl = $spoUrls.ConnectionUrl
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
                    if ($null -ne $Script:MSCloudLoginConnectionProfile.PnP.Scope -and `
                        $null -ne $Script:MSCloudLoginConnectionProfile.PnP.TokenUrl)
                    {
                        $accessToken = Get-MSCloudLoginAccessToken -ConnectionUri $Script:MSCloudLoginConnectionProfile.PnP.Scope `
                            -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.PnP.AuthorizationUrl `
                            -AzureADAuthorizationEndpointUri $Script:MSCloudLoginConnectionProfile.PnP.TokenUrl `
                            -ApplicationId $Script:MSCloudLoginConnectionProfile.PnP.ApplicationId `
                            -TenantId $Script:MSCloudLoginConnectionProfile.PnP.TenantId `
                            -CertificateThumbprint $Script:MSCloudLoginConnectionProfile.PnP.CertificateThumbprint
                        $Script:MSCloudLoginConnectionProfile.PnP.AccessTokens += $accessToken

                        Add-MSCloudLoginAssistantEvent -Message 'Connecting with Service Principal - Thumbprint' -Source $source
                        Add-MSCloudLoginAssistantEvent -Message "URL: $($Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl)" -Source $source
                        Add-MSCloudLoginAssistantEvent -Message "ConnectionUrl: $($Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl)" -Source $source
                        Connect-PnPOnline -Url $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl `
                            -AccessToken $($accessToken) | Out-Null
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
                            -AzureADLoginEndPoint $Script:MSCloudLoginConnectionProfile.PnP.EndPoints.AzureADLoginEndPoint `
                            -MicrosoftGraphEndPoint $Script:MSCloudLoginConnectionProfile.PnP.EndPoints.MicrosoftGraphEndPoint | Out-Null
                    }
                }

                $Script:MSCloudLoginConnectionProfile.PnP.CompleteConnection()
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

                $Script:MSCloudLoginConnectionProfile.PnP.CompleteConnection()
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
                $Script:MSCloudLoginConnectionProfile.PnP.CompleteConnection()
            }
            elseif ($Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType -eq 'CredentialsWithTenantId')
            {
                throw 'You cannot specify TenantId with Credentials when connecting to PnP.'
            }
            elseif ($Script:MSCloudLoginConnectionProfile.Pnp.AuthenticationType -eq 'CredentialsWithApplicationId')
            {
                if ($Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl -or $ForceRefreshConnection)
                {
                    Add-MSCloudLoginAssistantEvent -Message 'Connecting with Credentials and Application Id' -Source $source
                    Add-MSCloudLoginAssistantEvent -Message "URL: $($Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl)" -Source $source
                    Add-MSCloudLoginAssistantEvent -Message "ConnectionUrl: $($Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl)" -Source $source
                    Add-MSCloudLoginAssistantEvent -Message "ApplicationId: $($Script:MSCloudLoginConnectionProfile.PnP.ApplicationId)" -Source $source
                    Connect-PnPOnline -Url $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl `
                        -ClientId $Script:MSCloudLoginConnectionProfile.PnP.ApplicationId `
                        -Credentials $Script:MSCloudLoginConnectionProfile.PnP.Credentials `
                        -AzureEnvironment $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment
                }
                else
                {
                    Add-MSCloudLoginAssistantEvent -Message 'Connecting with Credentials and Application Id' -Source $source
                    Add-MSCloudLoginAssistantEvent -Message "URL: $($Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl)" -Source $source
                    Add-MSCloudLoginAssistantEvent -Message "AdminUrl: $($Script:MSCloudLoginConnectionProfile.PnP.AdminUrl)" -Source $source
                    Add-MSCloudLoginAssistantEvent -Message "ApplicationId: $($Script:MSCloudLoginConnectionProfile.PnP.ApplicationId)" -Source $source
                    Connect-PnPOnline -Url $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl `
                        -ClientId $Script:MSCloudLoginConnectionProfile.PnP.ApplicationId `
                        -Credentials $Script:MSCloudLoginConnectionProfile.PnP.Credentials `
                        -AzureEnvironment $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment
                }

                $Script:MSCloudLoginConnectionProfile.PnP.CompleteConnection()
            }
            elseif ($Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType -eq 'Credentials')
            {
                if ($Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl -or $ForceRefreshConnection)
                {
                    Add-MSCloudLoginAssistantEvent -Message 'Connecting with Credentials using SPOManagementShell' -Source $source
                    Add-MSCloudLoginAssistantEvent -Message "URL: $($Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl)" -Source $source
                    Add-MSCloudLoginAssistantEvent -Message "ConnectionUrl: $($Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl)" -Source $source
                    Connect-PnPOnline -Url $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl `
                        -Credentials $Script:MSCloudLoginConnectionProfile.PnP.Credentials `
                        -AzureEnvironment $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment `
                        -ClientId $Script:MSCloudLoginConnectionProfile.PnP.ClientId
                }
                else
                {
                    Add-MSCloudLoginAssistantEvent -Message 'Connecting with Credentials using SPOManagementShell and AdminUrl' -Source $source
                    Add-MSCloudLoginAssistantEvent -Message "URL: $($Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl)" -Source $source
                    Add-MSCloudLoginAssistantEvent -Message "AdminUrl: $($Script:MSCloudLoginConnectionProfile.PnP.AdminUrl)" -Source $source
                    Connect-PnPOnline -Url $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl `
                        -Credentials $Script:MSCloudLoginConnectionProfile.PnP.Credentials `
                        -AzureEnvironment $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment `
                        -ClientId $Script:MSCloudLoginConnectionProfile.PnP.ClientId
                }

                $Script:MSCloudLoginConnectionProfile.PnP.CompleteConnection()
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

                $accessToken = Get-AuthToken -Resource $connectionURL -Identity
                Connect-PnPOnline -Url $connectionURL `
                    -AccessToken $accessToken `
                    -AzureEnvironment $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment `
                    -WarningAction 'Ignore'

                $Script:MSCloudLoginConnectionProfile.PnP.CompleteConnection()
            }
            elseif ($Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType -eq 'AccessTokens')
            {
                $AccessTokenValue = Get-MSCloudLoginAccessTokenValue -Token $Script:MSCloudLoginConnectionProfile.PnP.AccessTokens[0]
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

                $Script:MSCloudLoginConnectionProfile.PnP.CompleteConnection()
            }
            else
            {
                throw "Authentication type '$($Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType)' is not supported for workload 'PnP'."
            }
        }
    }
    catch
    {
        if ((Test-MSCloudLoginMFARequiredError -ErrorRecord $_) -and -not (Assert-IsNonInteractiveShell))
        {
            try
            {
                Connect-PnPOnline -Url $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl `
                    -Interactive `
                    -ClientId $Script:MSCloudLoginConnectionProfile.PnP.ApplicationId
                $Script:MSCloudLoginConnectionProfile.PnP.CompleteConnection($true)
            }
            catch
            {
                try
                {
                    Connect-PnPOnline -Url $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl -UseWebLogin
                    $Script:MSCloudLoginConnectionProfile.PnP.CompleteConnection($true)
                }
                catch
                {
                    $Script:MSCloudLoginConnectionProfile.PnP.Connected = $false
                    Add-MSCloudLoginAssistantEvent -Message "Failed to connect to PnP interactively after MFA-required error: $($_.Exception.Message)" -Source $source -EntryType 'Error'
                    throw
                }
            }
        }
        elseif ($_.Exception.Message -like '*The sign-in name or password does not match one in the Microsoft account system*' -and `
                -not (Assert-IsNonInteractiveShell))
        {
            # This error means that the account was trying to connect using MFA.
            try
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
                $Script:MSCloudLoginConnectionProfile.PnP.CompleteConnection($true)
            }
            catch
            {
                $Script:MSCloudLoginConnectionProfile.PnP.Connected = $false
                Add-MSCloudLoginAssistantEvent -Message "Failed to connect to PnP interactively: $($_.Exception.Message)" -Source $source -EntryType 'Error'
                throw
            }
        }
        elseif ($_.Exception.Message -like '*AADSTS65001: The user or administrator has not consented to use the application with ID*' -and `
                -not (Assert-IsNonInteractiveShell))
        {
            try
            {
                Register-PnPManagementShellAccess
                Connect-PnPOnline -UseWebLogin -Url $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl
                $Script:MSCloudLoginConnectionProfile.PnP.CompleteConnection()
            }
            catch
            {
                throw "The PnP.PowerShell Azure AD Application has not been granted access for this tenant. Please run 'Register-PnPManagementShellAccess' to grant access and try again after. Original error: $($_.Exception.Message)"
            }
        }
        else
        {
            $Script:MSCloudLoginConnectionProfile.PnP.Connected = $false
            Add-MSCloudLoginAssistantEvent -Message "Failed to connect to PnP: $($_.Exception.Message)" -Source $source -EntryType 'Error'
            throw
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
