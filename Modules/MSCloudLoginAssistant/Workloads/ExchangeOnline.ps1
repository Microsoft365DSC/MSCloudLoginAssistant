function Connect-MSCloudLoginExchangeOnline
{
    [CmdletBinding()]
    param(
        [Parameter()]
        [switch]
        $SkipPSSessionEvaluation
    )

    $InformationPreference = 'SilentlyContinue'
    $ProgressPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginExchangeOnline'

    Add-MSCloudLoginAssistantEvent -Message 'Trying to get the Get-AcceptedDomain command from within MSCloudLoginAssistant' -Source $source

    if ($Script:MSCloudLoginConnectionProfile.ExchangeOnline.CmdletsToLoad.Count -eq 0)
    {
        $loadAllCmdlets = $true
    }

    Add-MSCloudLoginAssistantEvent "Current loaded module: $($Script:MSCloudLoginCurrentLoadedModule)" -Source $source
    if ($Script:MSCloudLoginCurrentLoadedModule -eq 'EXO')
    {
        try
        {
            $null = Get-Command -Name Get-AcceptedDomain -ErrorAction Stop

            if (-not $loadAllCmdlets)
            {
                Add-MSCloudLoginAssistantEvent 'Checking for missing commands' -Source $source
                Add-MSCloudLoginAssistantEvent "Cmdlets to load: $($Script:MSCloudLoginConnectionProfile.ExchangeOnline.CmdletsToLoad -join ',')" -Source $source
                Add-MSCloudLoginAssistantEvent "Loaded Cmdlets: $($Script:MSCloudLoginConnectionProfile.ExchangeOnline.LoadedCmdlets -join ',')" -Source $source
                $missingCommands = $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CmdletsToLoad | Where-Object -FilterScript {
                    $Script:MSCloudLoginConnectionProfile.ExchangeOnline.LoadedCmdlets -notcontains $_
                }
                Add-MSCloudLoginAssistantEvent "Missing commands: $($missingCommands -join ',')" -Source $source
            }

            # $missingCommands is null if no missing commands are found
            Add-MSCloudLoginAssistantEvent "Loaded all cmdlets: $($Script:MSCloudLoginConnectionProfile.ExchangeOnline.LoadedAllCmdlets)" -Source $source
            if ($Script:MSCloudLoginConnectionProfile.ExchangeOnline.LoadedAllCmdlets -or (-not $loadAllCmdlets -and $null -eq $missingCommands))
            {
                Add-MSCloudLoginAssistantEvent -Message 'Succeeded' -Source $source
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected = $true
                return
            }
        }
        catch
        {
            Add-MSCloudLoginAssistantEvent -Message 'Failed' -Source $source
        }
    }

    if ($Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected -and `
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.SkipModuleReload)
    {
        $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected = $true
        return
    }

    Add-MSCloudLoginAssistantEvent -Message "Loaded Modules: $(Get-Module | Select-Object -ExpandProperty Name)" -Source $source
    $alreadyLoadedEXOProxyModules = Get-Module | Where-Object -FilterScript { $_.ExportedCommands.Keys.Contains('Get-AcceptedDomain') }
    foreach ($loadedModule in $alreadyLoadedEXOProxyModules)
    {
        Add-MSCloudLoginAssistantEvent -Message "Removing module {$($loadedModule.Name)} from current EXO session" -Source $source
        Remove-Module $loadedModule.Name -Force -Verbose:$false | Out-Null
    }

    [array]$activeSessions = Get-PSSession | Where-Object -FilterScript { $_.ComputerName -like '*outlook.office*' -and $_.State -eq 'Opened' }
    Add-MSCloudLoginAssistantEvent -Message "Active Sessions: $($activeSessions | Out-String)" -Source $source
    if (-not $SkipPSSessionEvaluation -and $activeSessions.Length -ge 1)
    {
        Add-MSCloudLoginAssistantEvent -Message "Found {$($activeSessions.Length)} existing Exchange Online Session" -Source $source
        $ProxyModule = Import-PSSession $activeSessions[0] `
            -DisableNameChecking `
            -AllowClobber
        Add-MSCloudLoginAssistantEvent -Message "Imported session into $ProxyModule" -Source $source
        Import-Module $ProxyModule -Global `
            -Verbose:$false | Out-Null
        $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected = $true
        Add-MSCloudLoginAssistantEvent -Message 'Reloaded the Exchange Module' -Source $source

        # Rerun the function to make sure we have all the necessary commands loaded
        # but prevent an infinite loop by skipping the PSSession evaluation
        Connect-MSCloudLoginExchangeOnline -SkipPSSessionEvaluation
        return
    }
    Add-MSCloudLoginAssistantEvent -Message 'No active Exchange Online session found.' -Source $source

    # Make sure we disconnect from any existing connections
    Disconnect-ExchangeOnline -Confirm:$false
    $CommandName = @{}
    if ($Script:MSCloudLoginConnectionProfile.ExchangeOnline.CmdletsToLoad.Count -gt 0)
    {
        # Make sure we have the Get-AcceptedDomain command available
        if ($Script:MSCloudLoginConnectionProfile.ExchangeOnline.CmdletsToLoad -notcontains 'Get-AcceptedDomain')
        {
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CmdletsToLoad += 'Get-AcceptedDomain'
        }
        # Include the previously loaded commands, if available
        $combinedCmdlets = ($Script:MSCloudLoginConnectionProfile.ExchangeOnline.CmdletsToLoad + $Script:MSCloudLoginConnectionProfile.ExchangeOnline.LoadedCmdlets) | Select-Object -Unique
        $CommandName.Add('CommandName', $combinedCmdlets)
        Add-MSCloudLoginAssistantEvent -Message "Commands to load: $($CommandName.CommandName -join ',')" -Source $source
    }

    if ($Script:MSCloudLoginConnectionProfile.ExchangeOnline.AuthenticationType -eq 'ServicePrincipalWithThumbprint')
    {
        Add-MSCloudLoginAssistantEvent -Message "Attempting to connect to Exchange Online using AAD App {$($Script:MSCloudLoginConnectionProfile.ExchangeOnline.ApplicationId)}" -Source $source
        try
        {
            if ($PSVersionTable.Platform -ne 'Win32NT')
            {
                throw 'Certificate Thumbprint authentication is only supported on the Windows platform.'
            }
            if ($null -eq $Script:MSCloudLoginConnectionProfile.OrganizationName -or `
                $Script:MSCloudLoginConnectionProfile.OrganizationName -ne $Script:MSCloudLoginConnectionProfile.ExchangeOnline.TenantId)
            {
                $Script:MSCloudLoginConnectionProfile.OrganizationName = Get-MSCloudLoginOrganizationName `
                    -ApplicationId $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ApplicationId `
                    -TenantId $Script:MSCloudLoginConnectionProfile.ExchangeOnline.TenantId `
                    -CertificateThumbprint $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CertificateThumbprint
            }

            if (($null -ne $Script:MSCloudLoginConnectionProfile.ExchangeOnline.EndPoints -or $Global:CustomEnvironment) -and `
                $null -ne $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ConnectionUri -and `
                $null -ne $Script:MSCloudLoginConnectionProfile.ExchangeOnline.AzureADAuthorizationEndpointUri)
            {
                Add-MSCloudLoginAssistantEvent -Message 'Connecting by endpoints URI' -Source $source
                Connect-ExchangeOnline -AppId $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ApplicationId `
                    -Organization $Script:MSCloudLoginConnectionProfile.OrganizationName `
                    -CertificateThumbprint $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CertificateThumbprint `
                    -ShowBanner:$false `
                    -ShowProgress:$false `
                    -ConnectionUri $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ConnectionUri `
                    -AzureADAuthorizationEndpointUri $Script:MSCloudLoginConnectionProfile.ExchangeOnline.AzureADAuthorizationEndpointUri `
                    -Verbose:$false `
                    -SkipLoadingCmdletHelp `
                    @CommandName | Out-Null
            }
            else
            {
                Add-MSCloudLoginAssistantEvent -Message 'Connecting by environment name' -Source $source
                Connect-ExchangeOnline -AppId $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ApplicationId `
                    -Organization $Script:MSCloudLoginConnectionProfile.OrganizationName `
                    -CertificateThumbprint $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CertificateThumbprint `
                    -ShowBanner:$false `
                    -ShowProgress:$false `
                    -ExchangeEnvironmentName $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ExchangeEnvironmentName `
                    -Verbose:$false `
                    -SkipLoadingCmdletHelp `
                    @CommandName | Out-Null
            }

            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ConnectedDateTime = [System.DateTime]::Now.ToString()
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected = $true
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.MultiFactorAuthentication = $false
            Add-MSCloudLoginAssistantEvent -Message "Successfully connected to Exchange Online using AAD App {$ApplicationID}" -Source $source
        }
        catch
        {
            throw $_
        }
    }
    elseif ($Script:MSCloudLoginConnectionProfile.ExchangeOnline.AuthenticationType -eq 'ServicePrincipalWithPath')
    {
        Add-MSCloudLoginAssistantEvent -Message "Attempting to connect to Exchange Online using AAD App {$($Script:MSCloudLoginConnectionProfile.ExchangeOnline.ApplicationId)} with Certificate Path" -Source $source
        try
        {
            if ($null -eq $Script:MSCloudLoginConnectionProfile.OrganizationName -or `
                $Script:MSCloudLoginConnectionProfile.OrganizationName -ne $Script:MSCloudLoginConnectionProfile.ExchangeOnline.TenantId)
            {
                # Because we are using Certificate Path, an authentication to Graph is not possible here
                if ($Script:MSCloudLoginConnectionProfile.ExchangeOnline.TenantId -notlike '*.onmicrosoft.com')
                {
                    throw 'TenantId must be specified as the primary domain in the format <domain>.onmicrosoft.com when using Certificate Path authentication.'
                }
                $Script:MSCloudLoginConnectionProfile.OrganizationName = $Script:MSCloudLoginConnectionProfile.ExchangeOnline.TenantId
            }

            if ($PSVersionTable.Platform -eq 'Win32NT' -and -not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
            {
                throw 'Certificate Path authentication on Windows requires the command to be run as Administrator.'
            }

            Connect-ExchangeOnline -AppId $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ApplicationId `
                -Organization $Script:MSCloudLoginConnectionProfile.OrganizationName `
                -CertificateFilePath $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CertificatePath `
                -CertificatePassword $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CertificatePassword `
                -ShowBanner:$false `
                -ShowProgress:$false `
                -ExchangeEnvironmentName $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ExchangeEnvironmentName `
                -Verbose:$false `
                -SkipLoadingCmdletHelp `
                @CommandName | Out-Null

            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ConnectedDateTime = [System.DateTime]::Now.ToString()
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected = $true
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.MultiFactorAuthentication = $false
            Add-MSCloudLoginAssistantEvent -Message "Successfully connected to Exchange Online using AAD App {$ApplicationID} with Certificate Path" -Source $source
        }
        catch
        {
            throw $_
        }
    }
    elseif ($Script:MSCloudLoginConnectionProfile.ExchangeOnline.AuthenticationType -eq 'Credentials')
    {
        try
        {
            Add-MSCloudLoginAssistantEvent -Message 'Attempting to connect to Exchange Online using Credentials without MFA' -Source $source

            Connect-ExchangeOnline -Credential $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Credentials `
                -ShowProgress:$false `
                -ShowBanner:$false `
                -ExchangeEnvironmentName $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ExchangeEnvironmentName `
                -Verbose:$false `
                -ErrorAction Stop `
                -SkipLoadingCmdletHelp `
                @CommandName | Out-Null
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ConnectedDateTime = [System.DateTime]::Now.ToString()
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected = $true
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.MultiFactorAuthentication = $false
            Add-MSCloudLoginAssistantEvent -Message 'Successfully connected to Exchange Online using Credentials without MFA' -Source $source
        }
        catch
        {
            if ($_.Exception -like '*you must use multi-factor authentication to access*' -or $_.Exception -like '*WAM Error*')
            {
                Connect-MSCloudLoginExchangeOnlineMFA -Credentials $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Credentials
            }
            else
            {
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected = $false
                throw $_
            }
        }
    }
    elseif ($Script:MSCloudLoginConnectionProfile.ExchangeOnline.AuthenticationType -eq 'CredentialsWithTenantId')
    {
        try
        {
            Add-MSCloudLoginAssistantEvent -Message 'Attempting to connect to Exchange Online using Credentials without MFA' -Source $source

            Connect-ExchangeOnline -Credential $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Credentials `
                -ShowProgress:$false `
                -ShowBanner:$false `
                -DelegatedOrganization $Script:MSCloudLoginConnectionProfile.ExchangeOnline.TenantId `
                -ExchangeEnvironmentName $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ExchangeEnvironmentName `
                -Verbose:$false `
                -ErrorAction Stop `
                -SkipLoadingCmdletHelp `
                @CommandName | Out-Null
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ConnectedDateTime = [System.DateTime]::Now.ToString()
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected = $true
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.MultiFactorAuthentication = $false
            Add-MSCloudLoginAssistantEvent -Message 'Successfully connected to Exchange Online using Credentials & TenantId without MFA' -Source $source
        }
        catch
        {
            if ($_.Exception -like '*you must use multi-factor authentication to access*' -or $_.Exception -like '*WAM Error*')
            {
                Connect-MSCloudLoginExchangeOnlineMFA -Credentials $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Credentials `
                    -TenantId $Script:MSCloudLoginConnectionProfile.ExchangeOnline.TenantId
            }
            else
            {
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected = $false
                throw $_
            }
        }
    }
    elseif ($Script:MSCloudLoginConnectionProfile.ExchangeOnline.AuthenticationType -eq 'Identity')
    {
        Add-MSCloudLoginAssistantEvent -Message 'Attempting to connect to Exchange Online using Managed Identity' -Source $source
        try
        {
            if ($null -eq $Script:MSCloudLoginConnectionProfile.OrganizationName)
            {
                $Script:MSCloudLoginConnectionProfile.OrganizationName = Get-MSCloudLoginOrganizationName -Identity
            }

            Connect-ExchangeOnline -AppId $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ApplicationId `
                -Organization $Script:MSCloudLoginConnectionProfile.OrganizationName `
                -ManagedIdentity `
                -ShowBanner:$false `
                -ShowProgress:$false `
                -ExchangeEnvironmentName $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ExchangeEnvironmentName `
                -Verbose:$false `
                -SkipLoadingCmdletHelp `
                @CommandName | Out-Null

            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ConnectedDateTime = [System.DateTime]::Now.ToString()
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected = $false
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.MultiFactorAuthentication = $true
            Add-MSCloudLoginAssistantEvent -Message 'Successfully connected to Exchange Online using Managed Identity' -Source $source
        }
        catch
        {
            throw $_
        }
    }
    elseif ($Script:MSCloudLoginConnectionProfile.ExchangeOnline.AuthenticationType -eq 'AccessTokens')
    {
        Add-MSCloudLoginAssistantEvent -Message 'Connecting to EXO with AccessTokens' -Source $source
        try
        {
            $AccessTokenValue = $Script:MSCloudLoginConnectionProfile.ExchangeOnline.AccessTokens[0]
            if ($AccessTokenValue.GetType().Name -eq 'PSCredential')
            {
                $Ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToCoTaskMemUnicode($AccessTokenValue.Password)
                $AccessTokenValue = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($Ptr)
                [System.Runtime.InteropServices.Marshal]::ZeroFreeCoTaskMemUnicode($Ptr)
            }
            Connect-ExchangeOnline -AccessToken $AccessTokenValue `
                -Organization $Script:MSCloudLoginConnectionProfile.ExchangeOnline.TenantId `
                -ShowBanner:$false `
                -ShowProgress:$false `
                -ExchangeEnvironmentName $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ExchangeEnvironmentName `
                -Verbose:$false `
                -SkipLoadingCmdletHelp `
                @CommandName | Out-Null

            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ConnectedDateTime = [System.DateTime]::Now.ToString()
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected = $false
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.MultiFactorAuthentication = $false
            Add-MSCloudLoginAssistantEvent -Message 'Successfully connected to Exchange Online using Access Token' -Source $source
        }
        catch
        {
            throw $_
        }
    }
    else
    {
        Add-MSCloudLoginAssistantEvent -Message 'No valid authentication type found' -Source $source
        throw 'No valid authentication type found'
    }
    $Script:MSCloudLoginCurrentLoadedModule = 'EXO'

    # Usually the tmpEXO* modules, but it might also be from another PSSession
    $loadedEXOProxyModule = Get-Module | Where-Object -FilterScript { $_.ExportedCommands.Keys.Contains('Get-AcceptedDomain') }
    $loadedEXOModule = Get-Module -Name 'ExchangeOnlineManagement'
    $Script:MSCloudLoginConnectionProfile.ExchangeOnline.LoadedCmdlets = $loadedEXOProxyModule.ExportedCommands.Keys + $loadedEXOModule.ExportedCommands.Keys
    if ($loadAllCmdlets)
    {
        $Script:MSCloudLoginConnectionProfile.ExchangeOnline.LoadedAllCmdlets = $true
    }
}

function Connect-MSCloudLoginExchangeOnlineMFA
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential]
        $Credentials,

        [Parameter()]
        [System.String]
        $TenantId
    )

    $ProgressPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginExchangeOnlineMFA'

    try
    {
        if ([System.String]::IsNullOrEmpty($TenantId))
        {
            Add-MSCloudLoginAssistantEvent -Message 'Creating a new ExchangeOnline Session using MFA' -Source $source
            Connect-ExchangeOnline -UserPrincipalName $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Credentials.UserName `
                -ShowBanner:$false `
                -ShowProgress:$false `
                -ExchangeEnvironmentName $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ExchangeEnvironmentName `
                -Verbose:$false `
                -SkipLoadingCmdletHelp `
                @CommandName | Out-Null
            Add-MSCloudLoginAssistantEvent -Message 'Successfully connected to Exchange Online using credentials with MFA' -Source $source
        }
        else
        {
            Add-MSCloudLoginAssistantEvent -Message 'Creating a new ExchangeOnline Session using MFA with Credentials and TenantId' -Source $source
            Connect-ExchangeOnline -UserPrincipalName $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Credentials.UserName `
                -ShowBanner:$false `
                -ShowProgress:$false `
                -DelegatedOrganization $TenantId `
                -ExchangeEnvironmentName $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ExchangeEnvironmentName `
                -Verbose:$false `
                -SkipLoadingCmdletHelp `
                @CommandName | Out-Null
            Add-MSCloudLoginAssistantEvent -Message 'Successfully connected to Exchange Online using credentials and tenantId with MFA' -Source $source
        }
        $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ConnectedDateTime = [System.DateTime]::Now.ToString()
        $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected = $true
        $Script:MSCloudLoginConnectionProfile.ExchangeOnline.MultiFactorAuthentication = $true

    }
    catch
    {
        throw $_
    }
}
