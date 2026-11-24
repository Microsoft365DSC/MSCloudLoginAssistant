<#
.SYNOPSIS
    Returns the Windows principal of the current user.

.DESCRIPTION
    Thin wrapper around the Windows identity APIs used to decide whether the
    current session is elevated. Extracted so that unit tests can substitute
    the principal without depending on the privileges of the test runner.
#>
function Get-MSCloudLoginWindowsPrincipal
{
    [CmdletBinding()]
    [OutputType([System.Security.Principal.WindowsPrincipal])]
    param()

    return [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
}

function Connect-MSCloudLoginExchangeOnline
{
    [CmdletBinding()]
    param()

    $InformationPreference = 'SilentlyContinue'
    $ProgressPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginExchangeOnline'

    Add-MSCloudLoginAssistantEvent -Message 'Trying to get the Get-OrganizationConfig command from within MSCloudLoginAssistant' -Source $source

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
            $probe = Get-Command -Name Get-OrganizationConfig -ErrorAction Stop
            $liveConnection = Get-ConnectionInformation | Where-Object -FilterScript {
                $_.ModuleName -eq $probe.Module.ModuleBase -and $_.IsEopSession -ne $true -and
                ([System.String]::IsNullOrEmpty($Script:MSCloudLoginConnectionProfile.ExchangeOnline.ApplicationId) -or
                    $_.AppId -eq $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ApplicationId)
            }

            if ($null -eq $liveConnection)
            {
                throw 'The loaded Exchange Online proxy module has no live connection for this profile.'
            }

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
        # Shared commands such as Get-Group must resolve to the Exchange Online proxy module.
        if (Restore-MSCloudLoginProxyModule -ProbeCommand 'Get-OrganizationConfig' -Source $source)
        {
            $Script:MSCloudLoginCurrentLoadedModule = 'EXO'
            Add-MSCloudLoginAssistantEvent -Message 'Exchange Online is already connected' -Source $source
            return
        }

        Add-MSCloudLoginAssistantEvent -Message 'Exchange Online proxy module is no longer loaded, reconnecting' -Source $source
        $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected = $false
    }

    # IsEopSession marks Security & Compliance connections.
    [array]$currentSessions = Get-ConnectionInformation | Where-Object -FilterScript {
        $_.Name -like 'ExchangeOnline_*' -and $_.IsEopSession -ne $true
    }
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
                try
                {
                    Import-Module $filteredSessions.ModuleName -Force -Global -DisableNameChecking -ErrorAction Stop
                }
                catch
                {
                    Add-MSCloudLoginAssistantEvent -Message "Could not import the module of the active session: $($_.Exception.Message)" -Source $source
                    $filteredSessions = $null
                }
                if ($null -ne $filteredSessions)
                {
                    $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CompleteConnection($Script:MSCloudLoginConnectionProfile.ExchangeOnline.MultiFactorAuthentication)
                    $Script:MSCloudLoginConnectionProfile.ExchangeOnline.LoadedAllCmdlets = $true
                    $Script:MSCloudLoginCurrentLoadedModule = 'EXO'
                    return
                }
            }
        }

        if (-not [System.String]::IsNullOrEmpty($Script:MSCloudLoginConnectionProfile.ExchangeOnline.Credentials.UserName))
        {
            $filteredSessions = $currentSessions | Where-Object { $_.UserPrincipalName -eq $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Credentials.UserName }

            if ($filteredSessions.Count -gt 0)
            {
                Add-MSCloudLoginAssistantEvent -Message "Found an active Exchange Online session for UserPrincipalName {$($Script:MSCloudLoginConnectionProfile.ExchangeOnline.Credentials.UserName)}" -Source $source
                try
                {
                    Import-Module $filteredSessions.ModuleName -Force -Global -DisableNameChecking -ErrorAction Stop
                }
                catch
                {
                    Add-MSCloudLoginAssistantEvent -Message "Could not import the module of the active session: $($_.Exception.Message)" -Source $source
                    $filteredSessions = $null
                }
                if ($null -ne $filteredSessions)
                {
                    $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CompleteConnection($Script:MSCloudLoginConnectionProfile.ExchangeOnline.MultiFactorAuthentication)
                    $Script:MSCloudLoginConnectionProfile.ExchangeOnline.LoadedAllCmdlets = $true
                    $Script:MSCloudLoginCurrentLoadedModule = 'EXO'
                    return
                }
            }
        }
    }
    Add-MSCloudLoginAssistantEvent -Message 'No active Exchange Online session found.' -Source $source

    Add-MSCloudLoginAssistantEvent -Message "Loaded Modules: $(Get-Module | Select-Object -ExpandProperty Name)" -Source $source
    # Security & Compliance connections of other runspaces stay connected.
    Disconnect-MSCloudLoginExchangeConnection -Source $source
    Remove-MSCloudLoginProxyModule -ProbeCommand 'Get-OrganizationConfig' -Source $source
    $CommandName = @{}
    if ($Script:MSCloudLoginConnectionProfile.ExchangeOnline.CmdletsToLoad.Count -gt 0)
    {
        # Make sure we have the Get-OrganizationConfig command available
        if ($Script:MSCloudLoginConnectionProfile.ExchangeOnline.CmdletsToLoad -notcontains 'Get-OrganizationConfig')
        {
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CmdletsToLoad += 'Get-OrganizationConfig'
        }
        # Include the previously loaded commands, if available
        $combinedCmdlets = ($Script:MSCloudLoginConnectionProfile.ExchangeOnline.CmdletsToLoad + $Script:MSCloudLoginConnectionProfile.ExchangeOnline.LoadedCmdlets) | Select-Object -Unique
        $CommandName.Add('CommandName', $combinedCmdlets)
        Add-MSCloudLoginAssistantEvent -Message "Commands to load: $($CommandName.CommandName -join ',')" -Source $source
    }

    $workloadProfile = $Script:MSCloudLoginConnectionProfile.ExchangeOnline
    $authenticationType = $workloadProfile.AuthenticationType
    $authenticationDescriptions = @{
        ServicePrincipalWithThumbprint = 'Certificate Thumbprint'
        ServicePrincipalWithPath       = 'Certificate Path'
        Credentials                    = 'Credentials'
        CredentialsWithApplicationId   = 'Credentials and ApplicationId'
        CredentialsWithTenantId        = 'Credentials and TenantId'
        Identity                       = 'Managed Identity'
        AccessTokens                   = 'Access Token'
        Interactive                    = 'Interactive sign-in'
    }

    if ([System.String]::IsNullOrEmpty($authenticationType) -or -not $authenticationDescriptions.ContainsKey($authenticationType))
    {
        Add-MSCloudLoginAssistantEvent -Message 'No valid authentication type found' -Source $source
        throw 'No valid authentication type found'
    }
    $authenticationDescription = $authenticationDescriptions[$authenticationType]

    $authenticationParameters = @{}
    $connectionParameters = @{}
    $mfaUsed = $false
    try
    {
        $connectionParameters = Get-MSCloudLoginExchangeOnlineConnectionParameters -WorkloadProfile $workloadProfile -CommandName $CommandName
        $organization = Resolve-MSCloudLoginWorkloadTenant -WorkloadProfile $workloadProfile

        switch ($authenticationType)
        {
            'ServicePrincipalWithThumbprint'
            {
                if ($PSVersionTable.PSVersion -gt [Version]'6.0' -and $PSVersionTable.Platform -ne 'Win32NT')
                {
                    throw 'Certificate Thumbprint authentication is only supported on the Windows platform.'
                }
                $authenticationParameters.AppId                 = $workloadProfile.ApplicationId
                $authenticationParameters.Organization          = $organization
                $authenticationParameters.CertificateThumbprint = $workloadProfile.CertificateThumbprint
            }
            'ServicePrincipalWithPath'
            {
                if ($organization -notmatch '\.(onmicrosoft|onsovcloud|onmschina)\.')
                {
                    throw 'TenantId must be specified as the primary domain in the format <domain>.onmicrosoft.com when using Certificate Path authentication.'
                }
                if (($IsWindows -or $PSVersionTable.PSVersion.Major -eq 5) -and `
                    -not (Get-MSCloudLoginWindowsPrincipal).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))
                {
                    throw 'Certificate Path authentication on Windows requires the command to be run as Administrator.'
                }
                $authenticationParameters.AppId               = $workloadProfile.ApplicationId
                $authenticationParameters.Organization        = $organization
                $authenticationParameters.CertificateFilePath = $workloadProfile.CertificatePath
                $authenticationParameters.CertificatePassword = $workloadProfile.CertificatePassword
            }
            'Credentials'
            {
                $authenticationParameters.Credential = $workloadProfile.Credentials
            }
            'CredentialsWithApplicationId'
            {
                $authenticationParameters.Credential = $workloadProfile.Credentials
                $authenticationParameters.AppId      = $workloadProfile.ApplicationId
            }
            'CredentialsWithTenantId'
            {
                $authenticationParameters.Credential            = $workloadProfile.Credentials
                $authenticationParameters.DelegatedOrganization = $workloadProfile.TenantId
            }
            'Identity'
            {
                $authenticationParameters.ManagedIdentity = $true
                $authenticationParameters.Organization    = $organization
                if (-not [System.String]::IsNullOrEmpty($workloadProfile.ApplicationId))
                {
                    $authenticationParameters.AppId = $workloadProfile.ApplicationId
                }
                $mfaUsed = $true
            }
            'AccessTokens'
            {
                $authenticationParameters.AccessToken  = Get-MSCloudLoginAccessTokenValue -Token $workloadProfile.AccessTokens[0]
                $authenticationParameters.Organization = $organization
            }
            'Interactive'
            {
                if (Assert-IsNonInteractiveShell)
                {
                    throw 'Interactive authentication to Exchange Online requires an interactive shell.'
                }
                $mfaUsed = $true
            }
        }

        Add-MSCloudLoginAssistantEvent -Message "Attempting to connect to Exchange Online with $authenticationDescription" -Source $source
        Connect-ExchangeOnline @authenticationParameters @connectionParameters | Out-Null
        $workloadProfile.CompleteConnection($mfaUsed)
        Add-MSCloudLoginAssistantEvent -Message "Successfully connected to Exchange Online with $authenticationDescription" -Source $source
    }
    catch
    {
        if ($authenticationType -in @('Credentials', 'CredentialsWithApplicationId', 'CredentialsWithTenantId') -and
            (Test-MSCloudLoginMFARequiredError -ErrorRecord $_ -AdditionalPatterns @('*WAM Error*')) -and
            -not (Assert-IsNonInteractiveShell))
        {
            $mfaParameters = @{
                Credentials          = $workloadProfile.Credentials
                ConnectionParameters = $connectionParameters
            }
            if ($authenticationType -eq 'CredentialsWithTenantId')
            {
                $mfaParameters.TenantId = $workloadProfile.TenantId
            }
            elseif ($authenticationType -eq 'CredentialsWithApplicationId')
            {
                $mfaParameters.ApplicationId = $workloadProfile.ApplicationId
            }
            Connect-MSCloudLoginExchangeOnlineMFA @mfaParameters
        }
        else
        {
            $workloadProfile.Connected = $false
            Add-MSCloudLoginAssistantEvent -Message "Failed to connect to Exchange Online with ${authenticationDescription}: $($_.Exception.Message)" -Source $source -EntryType 'Error'
            throw
        }
    }
    $Script:MSCloudLoginCurrentLoadedModule = 'EXO'

    # Usually the tmpEXO* modules, but it might also be from another PSSession
    $loadedEXOProxyModule = Get-Module | Where-Object -FilterScript { $_.ExportedCommands.Keys.Contains('Get-OrganizationConfig') }
    $loadedEXOModule = Get-Module -Name 'ExchangeOnlineManagement'
    $Script:MSCloudLoginConnectionProfile.ExchangeOnline.LoadedCmdlets = $loadedEXOProxyModule.ExportedCommands.Keys + $loadedEXOModule.ExportedCommands.Keys
    if ($loadAllCmdlets)
    {
        $Script:MSCloudLoginConnectionProfile.ExchangeOnline.LoadedAllCmdlets = $true
    }
}

<#
.SYNOPSIS
    Returns the parameters that every Connect-ExchangeOnline call shares.

.DESCRIPTION
    Selects the endpoint of the workload profile. A profile with a ConnectionUri connects by
    ConnectionUri and AzureADAuthorizationEndpointUri, all others by ExchangeEnvironmentName.
    The 'Custom' marker is not passed along because Connect-ExchangeOnline does not accept it.

.PARAMETER WorkloadProfile
    The Exchange Online workload profile.

.PARAMETER CommandName
    Optional hashtable with a 'CommandName' entry that limits the imported cmdlets.
#>
function Get-MSCloudLoginExchangeOnlineConnectionParameters
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.Object]
        $WorkloadProfile,

        [Parameter()]
        [System.Collections.Hashtable]
        $CommandName = @{}
    )

    $parameters = @{
        ShowBanner            = $false
        ShowProgress          = $false
        Verbose               = $false
        ErrorAction           = 'Stop'
        SkipLoadingCmdletHelp = $true
    }

    if (-not [System.String]::IsNullOrEmpty($WorkloadProfile.ConnectionUri))
    {
        $parameters.ConnectionUri = $WorkloadProfile.ConnectionUri
        if (-not [System.String]::IsNullOrEmpty($WorkloadProfile.AzureADAuthorizationEndpointUri))
        {
            $parameters.AzureADAuthorizationEndpointUri = $WorkloadProfile.AzureADAuthorizationEndpointUri
        }
    }
    elseif ($WorkloadProfile.ExchangeEnvironmentName -eq 'Custom')
    {
        throw "No Exchange Online ConnectionUri is defined for environment '$($WorkloadProfile.EnvironmentName)'. Set CustomEXOConnectionUri in the custom environment configuration."
    }
    else
    {
        $parameters.ExchangeEnvironmentName = $WorkloadProfile.ExchangeEnvironmentName
    }

    if ($null -ne $CommandName -and $CommandName.ContainsKey('CommandName'))
    {
        $parameters.CommandName = $CommandName.CommandName
    }

    return $parameters
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
        [System.String]
        $ApplicationId,

        [Parameter()]
        [System.Collections.Hashtable]
        $ConnectionParameters
    )

    $ProgressPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginExchangeOnlineMFA'

    try
    {
        if ($null -eq $ConnectionParameters -or $ConnectionParameters.Count -eq 0)
        {
            $ConnectionParameters = Get-MSCloudLoginExchangeOnlineConnectionParameters -WorkloadProfile $Script:MSCloudLoginConnectionProfile.ExchangeOnline
        }

        $parameters = $ConnectionParameters.Clone()
        $parameters.UserPrincipalName = $Credentials.UserName
        if (-not [System.String]::IsNullOrEmpty($TenantId))
        {
            $parameters.DelegatedOrganization = $TenantId
        }
        if (-not [System.String]::IsNullOrEmpty($ApplicationId))
        {
            $parameters.AppId = $ApplicationId
        }

        Add-MSCloudLoginAssistantEvent -Message 'Creating a new Exchange Online session using MFA' -Source $source
        Connect-ExchangeOnline @parameters | Out-Null
        Add-MSCloudLoginAssistantEvent -Message 'Successfully connected to Exchange Online using MFA' -Source $source
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
        Disconnect-MSCloudLoginExchangeConnection -Source $source
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
