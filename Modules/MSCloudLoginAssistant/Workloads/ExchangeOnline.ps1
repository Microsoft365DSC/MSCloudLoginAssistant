function Connect-MSCloudLoginExchangeOnline
{
    [CmdletBinding()]
    param()

    $InformationPreference = 'SilentlyContinue'
    $ProgressPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginExchangeOnline'

    Add-MSCloudLoginAssistantEvent -Message 'Trying to get the Get-AcceptedDomain command from within MSCloudLoginAssistant' -Source $source

    $loadAllCmdlets = $false
    if ($Script:MSCloudLoginConnectionProfile.ExchangeOnline.CmdletsToLoad.Count -eq 0)
    {
        $loadAllCmdlets = $true
    }

    Add-MSCloudLoginAssistantEvent -Message "Current loaded module: $($Script:MSCloudLoginCurrentLoadedModule)" -Source $source
    if ($Script:MSCloudLoginCurrentLoadedModule -eq 'EXO')
    {
        try
        {
            $null = Get-Command -Name Get-AcceptedDomain -ErrorAction Stop

            if (-not $loadAllCmdlets)
            {
                Add-MSCloudLoginAssistantEvent -Message 'Checking for missing commands' -Source $source
                Add-MSCloudLoginAssistantEvent -Message "Cmdlets to load: $($Script:MSCloudLoginConnectionProfile.ExchangeOnline.CmdletsToLoad -join ',')" -Source $source
                Add-MSCloudLoginAssistantEvent -Message "Loaded Cmdlets: $($Script:MSCloudLoginConnectionProfile.ExchangeOnline.LoadedCmdlets -join ',')" -Source $source
                $missingCommands = $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CmdletsToLoad | Where-Object -FilterScript {
                    $Script:MSCloudLoginConnectionProfile.ExchangeOnline.LoadedCmdlets -notcontains $_
                }
                Add-MSCloudLoginAssistantEvent -Message "Missing commands: $($missingCommands -join ',')" -Source $source
            }

            # $missingCommands is null if no missing commands are found
            Add-MSCloudLoginAssistantEvent -Message "Loaded all cmdlets: $($Script:MSCloudLoginConnectionProfile.ExchangeOnline.LoadedAllCmdlets)" -Source $source
            if ($Script:MSCloudLoginConnectionProfile.ExchangeOnline.LoadedAllCmdlets -or (-not $loadAllCmdlets -and $null -eq $missingCommands))
            {
                Add-MSCloudLoginAssistantEvent -Message 'Succeeded' -Source $source
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CompleteConnection($Script:MSCloudLoginConnectionProfile.ExchangeOnline.MultiFactorAuthentication)
                return
            }
        }
        catch
        {
            Add-MSCloudLoginAssistantEvent -Message "Probe for existing Exchange Online session failed: $($_.Exception.Message)" -Source $source
        }
    }

    if ($Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected)
    {
        Add-MSCloudLoginAssistantEvent -Message 'Exchange Online is already connected' -Source $source
        return
    }

    [array]$currentSessions = Get-ConnectionInformation | Where-Object -Property Name -Like 'ExchangeOnline_*'
    if ($null -ne $currentSessions -and $currentSessions.Count -gt 0)
    {
        Add-MSCloudLoginAssistantEvent -Message "Found {$($currentSessions.Count)} active Exchange Online session(s) but not connected" -Source $source
        if (-not [System.String]::IsNullOrEmpty($Script:MSCloudLoginConnectionProfile.ExchangeOnline.ApplicationId) -and
            -not [System.String]::IsNullOrEmpty($Script:MSCloudLoginConnectionProfile.ExchangeOnline.TenantId))
        {
            $filteredSessions = $currentSessions | Where-Object { $_.AppId -eq $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ApplicationId `
                -and $_.Organization -eq $Script:MSCloudLoginConnectionProfile.ExchangeOnline.TenantId }

            if ($filteredSessions.Count -gt 0)
            {
                Add-MSCloudLoginAssistantEvent -Message "Found an active Exchange Online session for ApplicationId {$($Script:MSCloudLoginConnectionProfile.ExchangeOnline.ApplicationId)} and TenantId {$($Script:MSCloudLoginConnectionProfile.ExchangeOnline.TenantId)}" -Source $source
                Import-Module $filteredSessions.ModuleName -Force -Global -DisableNameChecking
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CompleteConnection($Script:MSCloudLoginConnectionProfile.ExchangeOnline.MultiFactorAuthentication)
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.LoadedAllCmdlets = $true
                $Script:MSCloudLoginCurrentLoadedModule = 'EXO'
                return
            }
        }

        if (-not [System.String]::IsNullOrEmpty($Script:MSCloudLoginConnectionProfile.ExchangeOnline.Credentials.UserName))
        {
            $filteredSessions = $currentSessions | Where-Object { $_.UserPrincipalName -eq $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Credentials.UserName }

            if ($filteredSessions.Count -gt 0)
            {
                Add-MSCloudLoginAssistantEvent -Message "Found an active Exchange Online session for UserPrincipalName {$($Script:MSCloudLoginConnectionProfile.ExchangeOnline.Credentials.UserName)}" -Source $source
                Import-Module $filteredSessions.ModuleName -Force -Global -DisableNameChecking
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CompleteConnection($Script:MSCloudLoginConnectionProfile.ExchangeOnline.MultiFactorAuthentication)
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.LoadedAllCmdlets = $true
                $Script:MSCloudLoginCurrentLoadedModule = 'EXO'
                return
            }
        }
    }
    Add-MSCloudLoginAssistantEvent -Message 'No active Exchange Online session found.' -Source $source

    Add-MSCloudLoginAssistantEvent -Message "Loaded Modules: $(Get-Module | Select-Object -ExpandProperty Name)" -Source $source
    Remove-MSCloudLoginProxyModule -ProbeCommand 'Get-AcceptedDomain' -Source $source

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
            if ($PSVersionTable.PSVersion -gt [Version]'6.0' -and $PSVersionTable.Platform -ne 'Win32NT')
            {
                throw 'Certificate Thumbprint authentication is only supported on the Windows platform.'
            }

            if ($null -eq $Script:MSCloudLoginConnectionProfile.OrganizationName -or `
                $Script:MSCloudLoginConnectionProfile.OrganizationName -ne $Script:MSCloudLoginConnectionProfile.ExchangeOnline.TenantId)
            {
                <#
                $Script:MSCloudLoginConnectionProfile.OrganizationName = Get-MSCloudLoginOrganizationName `
                    -ApplicationId $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ApplicationId `
                    -TenantId $Script:MSCloudLoginConnectionProfile.ExchangeOnline.TenantId `
                    -CertificateThumbprint $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CertificateThumbprint
                #>
                $Script:MSCloudLoginConnectionProfile.OrganizationName = $Script:MSCloudLoginConnectionProfile.ExchangeOnline.TenantId
            }

            if (($Script:CustomEnvConfig.CustomEnvironment -or $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ExchangeEnvironmentName -eq 'Custom') -and `
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

            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CompleteConnection()
            Add-MSCloudLoginAssistantEvent -Message "Successfully connected to Exchange Online using AAD App {$($Script:MSCloudLoginConnectionProfile.ExchangeOnline.ApplicationId)}" -Source $source
        }
        catch
        {
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected = $false
            Add-MSCloudLoginAssistantEvent -Message "Failed to connect to Exchange Online: $($_.Exception.Message)" -Source $source -EntryType 'Error'
            throw
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

            if (($IsWindows -or $PSVersionTable.PSVersion.Major -eq 5) -and -not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
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

            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CompleteConnection()
            Add-MSCloudLoginAssistantEvent -Message "Successfully connected to Exchange Online using AAD App {$($Script:MSCloudLoginConnectionProfile.ExchangeOnline.ApplicationId)} with Certificate Path" -Source $source
        }
        catch
        {
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected = $false
            Add-MSCloudLoginAssistantEvent -Message "Failed to connect to Exchange Online with Certificate Path: $($_.Exception.Message)" -Source $source -EntryType 'Error'
            throw
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
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CompleteConnection()
            Add-MSCloudLoginAssistantEvent -Message 'Successfully connected to Exchange Online using Credentials without MFA' -Source $source
        }
        catch
        {
            if ((Test-MSCloudLoginMFARequiredError -ErrorRecord $_ -AdditionalPatterns @('*WAM Error*')) -and -not (Assert-IsNonInteractiveShell))
            {
                Connect-MSCloudLoginExchangeOnlineMFA -Credentials $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Credentials -CommandName $CommandName
            }
            else
            {
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected = $false
                Add-MSCloudLoginAssistantEvent -Message "Failed to connect to Exchange Online with Credentials: $($_.Exception.Message)" -Source $source -EntryType 'Error'
                throw
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
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CompleteConnection()
            Add-MSCloudLoginAssistantEvent -Message 'Successfully connected to Exchange Online using Credentials & TenantId without MFA' -Source $source
        }
        catch
        {
            if ((Test-MSCloudLoginMFARequiredError -ErrorRecord $_ -AdditionalPatterns @('*WAM Error*')) -and -not (Assert-IsNonInteractiveShell))
            {
                Connect-MSCloudLoginExchangeOnlineMFA -Credentials $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Credentials `
                    -TenantId $Script:MSCloudLoginConnectionProfile.ExchangeOnline.TenantId `
                    -CommandName $CommandName
            }
            else
            {
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected = $false
                Add-MSCloudLoginAssistantEvent -Message "Failed to connect to Exchange Online with Credentials and TenantId: $($_.Exception.Message)" -Source $source -EntryType 'Error'
                throw
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
                #$Script:MSCloudLoginConnectionProfile.OrganizationName = Get-MSCloudLoginOrganizationName -Identity
                $Script:MSCloudLoginConnectionProfile.OrganizationName = $Script:MSCloudLoginConnectionProfile.ExchangeOnline.TenantId
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

            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CompleteConnection($true)
            Add-MSCloudLoginAssistantEvent -Message 'Successfully connected to Exchange Online using Managed Identity' -Source $source
        }
        catch
        {
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected = $false
            Add-MSCloudLoginAssistantEvent -Message "Failed to connect to Exchange Online with Managed Identity: $($_.Exception.Message)" -Source $source -EntryType 'Error'
            throw
        }
    }
    elseif ($Script:MSCloudLoginConnectionProfile.ExchangeOnline.AuthenticationType -eq 'AccessTokens')
    {
        Add-MSCloudLoginAssistantEvent -Message 'Connecting to EXO with AccessTokens' -Source $source
        try
        {
            $AccessTokenValue = Get-MSCloudLoginAccessTokenValue -Token $Script:MSCloudLoginConnectionProfile.ExchangeOnline.AccessTokens[0]
            Connect-ExchangeOnline -AccessToken $AccessTokenValue `
                -Organization $Script:MSCloudLoginConnectionProfile.ExchangeOnline.TenantId `
                -ShowBanner:$false `
                -ShowProgress:$false `
                -ExchangeEnvironmentName $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ExchangeEnvironmentName `
                -Verbose:$false `
                -SkipLoadingCmdletHelp `
                @CommandName | Out-Null

            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CompleteConnection()
            Add-MSCloudLoginAssistantEvent -Message 'Successfully connected to Exchange Online using Access Token' -Source $source
        }
        catch
        {
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected = $false
            Add-MSCloudLoginAssistantEvent -Message "Failed to connect to Exchange Online with Access Token: $($_.Exception.Message)" -Source $source -EntryType 'Error'
            throw
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
        $TenantId,

        [Parameter()]
        [System.Collections.Hashtable]
        $CommandName
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
        $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CompleteConnection($true)
    }
    catch
    {
        $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected = $false
        Add-MSCloudLoginAssistantEvent -Message "Failed to connect to Exchange Online using MFA: $($_.Exception.Message)" -Source $source -EntryType 'Error'
        throw
    }
}

function Disconnect-MSCloudLoginExchangeOnline
{
    [CmdletBinding()]
    param()

    $source = 'Disconnect-MSCloudLoginExchangeOnline'

    if ($Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected)
    {
        Add-MSCloudLoginAssistantEvent -Message 'Attempting to disconnect from Exchange Online' -Source $source
        Disconnect-ExchangeOnline -Confirm:$false
        $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected = $false
        $Script:MSCloudLoginConnectionProfile.ExchangeOnline.LoadedAllCmdlets = $false
        $Script:MSCloudLoginConnectionProfile.ExchangeOnline.LoadedCmdlets = @()
        $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CmdletsToLoad = @()
        Add-MSCloudLoginAssistantEvent -Message 'Successfully disconnected from Exchange Online' -Source $source
    }
    else
    {
        Add-MSCloudLoginAssistantEvent -Message 'No connections to Exchange Online were found' -Source $source
    }
}
