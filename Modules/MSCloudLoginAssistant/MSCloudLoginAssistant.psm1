$Script:WriteToEventLog = ([Environment]::GetEnvironmentVariable('MSCLOUDLOGINASSISTANT_WRITETOEVENTLOG', 'Machine') -eq 'true') -or `
                          ($env:MSCLOUDLOGINASSISTANT_WRITETOEVENTLOG -eq 'true')

$Script:CustomEnvConfig = Import-PowerShellDataFile -Path "$PSScriptRoot/CustomEnvironment.psd1"
$Script:LoadedCustomEnvFileName = 'CustomEnvironment.psd1'
. "$PSScriptRoot/ConnectionProfile.ps1" -CustomEnvironmentConfig $Script:CustomEnvConfig
. "$PSScriptRoot/Helpers.ps1"

$privateModules = Get-ChildItem -Path "$PSScriptRoot/Workloads" -Filter '*.ps1' -Recurse
foreach ($module in $privateModules)
{
    Write-Verbose "Importing workload $($module.FullName)"
    . $module.FullName
}

<#
.SYNOPSIS
    This function ensures that only one connection attempt is made at a time across all runspaces.

.DESCRIPTION
    This function ensures that only one connection attempt is made at a time across all runspaces.
    It uses a named mutex to synchronize access to the connection logic, preventing concurrent connection attempts that could lead to conflicts or inconsistent states.

.PARAMETER ConnectScript
    A script block that contains the connection logic to be executed while holding the mutex.

.PARAMETER Timeout
    The maximum time to wait for the mutex to be acquired. Defaults to 5 minutes.

.OUTPUTS
    None. This function does not return any output.
#>
function Invoke-MSCloudLoginAssistantConnectionLock
{
    param
    (
        [Parameter(Mandatory = $true)]
        [scriptblock]
        $ConnectScript,

        [Parameter()]
        [TimeSpan]
        $Timeout = [TimeSpan]::FromMinutes(5)
    )

    $mutex = [System.Threading.Mutex]::new($false, "Local\MSCloudLoginAssistant.Connect")
    $acquired = $false
    try
    {
        try
        {
            $acquired = $mutex.WaitOne($Timeout)
        }
        catch [System.Threading.AbandonedMutexException]
        {
            # a runspace died while holding it — we now own it anyway
            $acquired = $true
        }

        if (-not $acquired)
        {
            throw "Timed out waiting for the M365 connection lock."
        }

        & $ConnectScript
    }
    finally
    {
        if ($acquired)
        {
            $mutex.ReleaseMutex()
        }
        $mutex.Dispose()
    }
}

<#
.SYNOPSIS
    Connects to a Microsoft 365 workload using the specified authentication method.

.DESCRIPTION
    Connect-M365Tenant establishes an authenticated session to a target Microsoft 365 workload.
    It supports multiple authentication types including credentials, service principal with certificate
    thumbprint, service principal with secret, managed identity, and access tokens.
    The function maintains a connection profile and will reuse an existing connection unless the
    authentication parameters have changed.

.PARAMETER Workload
    The Microsoft 365 workload to connect to. Valid values are: AdminAPI, Azure, AzureDevOPS,
    EngageHub, ExchangeOnline, Fabric, Licensing, O365Portal, SecurityComplianceCenter, PnP,
    PowerPlatforms, PowerPlatformREST, MicrosoftTeams, MicrosoftGraph, SharePointOnlineREST,
    Tasks, DefenderForEndpoint.

.PARAMETER Url
    The URL to connect to. Required for workloads such as PnP and SharePointOnlineREST.

.PARAMETER Credential
    The PSCredential object containing username and password for credential-based authentication.

.PARAMETER ApplicationId
    The Application (client) ID of the Azure AD app registration used for service principal authentication.

.PARAMETER TenantId
    The Tenant ID (GUID or domain) of the Azure AD tenant.

.PARAMETER ApplicationSecret
    The client secret for service principal authentication with a secret.

.PARAMETER CertificateThumbprint
    The thumbprint of the certificate used for service principal authentication.

.PARAMETER UseModernAuth
    Switch to enable modern authentication.

.PARAMETER CertificatePassword
    The password for the certificate file used in service principal authentication with a certificate path.

.PARAMETER CertificatePath
    The file system path to the certificate (.pfx) used for service principal authentication.

.PARAMETER EnableSearchOnlySession
    Switch to enable a search-only session. Applicable to the SecurityComplianceCenter workload.

.PARAMETER Identity
    Switch to authenticate using a managed identity (system-assigned or user-assigned).

.PARAMETER AccessTokens
    An array of access tokens to use for authentication.

.PARAMETER Endpoints
    A hashtable of custom endpoint URLs to override defaults.

.PARAMETER ExchangeOnlineCmdlets
    An array of Exchange Online cmdlets to load. Only applicable when Workload is 'ExchangeOnline'.

.PARAMETER CustomEnvironmentFileName
    The name of the custom environment configuration file. Defaults to 'CustomEnvironment.psd1'.
    Must be located in the same directory as the CustomEnvironment.psd1 file (root of the module).

.OUTPUTS
    None. Connect-M365Tenant does not return any output.

.EXAMPLE
    PS> Connect-M365Tenant -Workload 'MicrosoftGraph' -ApplicationId '00000000-0000-0000-0000-000000000000' `
        -TenantId 'contoso.onmicrosoft.com' -CertificateThumbprint 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD'

.EXAMPLE
    PS> Connect-M365Tenant -Workload 'ExchangeOnline' -Credential $Credential

.EXAMPLE
    PS> Connect-M365Tenant -Workload 'MicrosoftGraph' -Identity

.EXAMPLE
    PS> Connect-M365Tenant -Workload 'PnP' -Url 'https://contoso-admin.sharepoint.com' `
        -ApplicationId '00000000-0000-0000-0000-000000000000' -TenantId 'contoso.onmicrosoft.com' `
        -CertificateThumbprint 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD'
#>
function Connect-M365Tenant
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateSet('AdminAPI', 'Azure', 'AzureDevOPS', 'EngageHub', 'ExchangeOnline', 'Fabric', 'Licensing', `
                'O365Portal', 'SecurityComplianceCenter', 'PnP', 'PowerPlatforms', "PowerPlatformREST", `
                'MicrosoftTeams', 'MicrosoftGraph', 'SharePointOnlineREST', 'Tasks', 'DefenderForEndpoint')]
        [System.String]
        $Workload,

        [Parameter()]
        [System.String]
        $Url,

        [Parameter()]
        [Alias('o365Credential')]
        [System.Management.Automation.PSCredential]
        $Credential,

        [Parameter()]
        [System.String]
        $ApplicationId,

        [Parameter()]
        [System.String]
        $TenantId,

        [Parameter()]
        [System.String]
        $ApplicationSecret,

        [Parameter()]
        [System.String]
        $CertificateThumbprint,

        [Parameter()]
        [Switch]
        $UseModernAuth,

        [Parameter()]
        [SecureString]
        $CertificatePassword,

        [Parameter()]
        [System.String]
        $CertificatePath,

        [Parameter()]
        [switch]
        $EnableSearchOnlySession,

        [Parameter()]
        [Switch]
        $Identity,

        [Parameter()]
        [System.String[]]
        $AccessTokens,

        [Parameter()]
        [System.Collections.Hashtable]
        $Endpoints,

        [Parameter()]
        [System.String]
        $CustomEnvironmentFileName = 'CustomEnvironment.psd1'
    )

    dynamicparam
    {
        $paramDictionary = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()
        if ($Workload -eq 'ExchangeOnline')
        {
            $parameterAttribute = [System.Management.Automation.ParameterAttribute]@{
                Mandatory = $false
            }

            $attributeCollection = [System.Collections.ObjectModel.Collection[System.Attribute]]::new()
            $attributeCollection.Add($parameterAttribute)

            $dynamicParameter = [System.Management.Automation.RuntimeDefinedParameter]::new(
                'ExchangeOnlineCmdlets', [System.String[]], $attributeCollection
            )

            $paramDictionary.Add('ExchangeOnlineCmdlets', $dynamicParameter)
        }

        if ($Workload -eq 'Azure')
        {
            $parameterAttribute = [System.Management.Automation.ParameterAttribute]@{
                Mandatory = $false
            }

            $attributeCollection = [System.Collections.ObjectModel.Collection[System.Attribute]]::new()
            $attributeCollection.Add($parameterAttribute)

            $dynamicParameter = [System.Management.Automation.RuntimeDefinedParameter]::new(
                'SubscriptionId', [System.String], $attributeCollection
            )

            $paramDictionary.Add('SubscriptionId', $dynamicParameter)
        }

        return $paramDictionary
    }

    process
    {
        $source = 'Connect-M365Tenant'
        $workloadInternalName = $Workload

        if ($Workload -eq 'MicrosoftTeams')
        {
            $workloadInternalName = 'Teams'
        }
        elseif ($Workload -eq 'PowerPlatforms')
        {
            $workloadInternalName = 'PowerPlatform'
        }

        # Only (re)load the custom environment configuration when it changed. The classes in
        # ConnectionProfile.ps1 read $Script:CustomEnvConfig at runtime, so re-dot-sourcing
        # the file on every call is unnecessary and would re-declare all classes each time.
        if ($null -eq $Script:CustomEnvConfig -or $Script:LoadedCustomEnvFileName -ne $CustomEnvironmentFileName)
        {
            $Script:CustomEnvConfig = Import-PowerShellDataFile -Path "$PSScriptRoot/$CustomEnvironmentFileName" -ErrorAction Stop
            $Script:LoadedCustomEnvFileName = $CustomEnvironmentFileName
        }

        if ($null -eq $Script:MSCloudLoginConnectionProfile)
        {
            $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
        }

        Add-MSCloudLoginAssistantEvent -Message "Checking connection to platform {$Workload}" -Source $source
        $authenticationParameters = @{}
        foreach ($parameter in $PSBoundParameters.GetEnumerator())
        {
            # Never propagate null or empty values onto the connection profile.
            if (Test-MSCloudLoginParameterValueEmpty -Value $parameter.Value)
            {
                continue
            }
            if ($parameter.Key -eq 'Credential')
            {
                $authenticationParameters.Add('Credentials', $parameter.Value)
            }
            elseif ($parameter.Key -eq 'ExchangeOnlineCmdlets')
            {
                # The parameter and the property on the ExchangeOnline workload have different names.
                $authenticationParameters.Add('CmdletsToLoad', $parameter.Value)
            }
            else
            {
                if ($parameter.Key -in @('AccessTokens', 'ApplicationId', 'ApplicationSecret', 'CertificateThumbprint', 'CertificatePath', 'CertificatePassword', 'Identity', 'Endpoints', 'TenantId', 'TenantGUID', 'SubscriptionId'))
                {
                    $authenticationParameters.Add($parameter.Key, $parameter.Value)
                }
            }
        }

        # Internal re-entrant calls (e.g. Get-SPOAdminUrl) pass no authentication parameters at all.
        # If the workload is already connected, reuse the session instead of deriving an
        # 'Interactive' authentication type and forcing a reconnect.
        $authInfluencingKeys = @('Credentials', 'ApplicationId', 'ApplicationSecret', 'CertificateThumbprint',
            'CertificatePath', 'CertificatePassword', 'Identity', 'AccessTokens')
        $hasAuthParameters = @($authenticationParameters.Keys | Where-Object { $_ -in $authInfluencingKeys }).Count -gt 0
        if ($hasAuthParameters -or -not $Script:MSCloudLoginConnectionProfile.$workloadInternalName.Connected)
        {
            $Script:MSCloudLoginConnectionProfile.$workloadInternalName.RequestedAuthenticationType = Get-AuthenticationTypeFromParameters -AuthenticationObject $authenticationParameters
        }

        # Parameters that describe the session rather than the identity. A call that only
        # changes one of them, without repeating any identity parameter, must still be
        # detected as drift - but then only those keys may be compared, because the identity
        # keys of the profile would otherwise all report as removed.
        $sessionParameterKeys = @('SubscriptionId', 'CmdletsToLoad', 'ConnectionUrl', 'EnableSearchOnlySession')
        $hasSessionParameters = @($authenticationParameters.Keys | Where-Object { $_ -in $sessionParameterKeys }).Count -gt 0
        if (-not $hasSessionParameters)
        {
            $hasSessionParameters = @($PSBoundParameters.Keys | Where-Object { $_ -in @('Url', 'EnableSearchOnlySession') }).Count -gt 0
        }

        # Only validate the parameters if we are already connected
        if (($hasAuthParameters -or $hasSessionParameters) `
                -and $Script:MSCloudLoginConnectionProfile.$workloadInternalName.Connected)
        {
            $compareParameters = @{
                CurrentParamSet = $PSBoundParameters
            }
            if (-not $hasAuthParameters)
            {
                $compareParameters.LimitToKeys = $sessionParameterKeys
            }

            if (Compare-InputParametersForChange @compareParameters)
            {
                Add-MSCloudLoginAssistantEvent -Message "Resetting connection for workload $workloadInternalName" -Source $source
                $Script:MSCloudLoginConnectionProfile.$workloadInternalName.Connected = $false
            }
        }

        # Optional parameters that describe the session rather than the identity. Omitting one that
        # was provided on an earlier call is drift as well, so the profile falls back to the unset
        # value instead of silently keeping the previous one.
        $optionalSessionParameters = @{
            Azure          = @{ SubscriptionId = '' }
            ExchangeOnline = @{ CmdletsToLoad = @() }
        }

        # Apply the parameters to the connection profile, but only when a (re)connect is about
        # to happen. This keeps the invariant that the profile describes the live session.
        if (-not $Script:MSCloudLoginConnectionProfile.$workloadInternalName.Connected)
        {
            foreach ($key in $authenticationParameters.Keys)
            {
                $Script:MSCloudLoginConnectionProfile.$workloadInternalName.($key) = $authenticationParameters[$key]
            }

            if ($optionalSessionParameters.ContainsKey($workloadInternalName))
            {
                foreach ($key in $optionalSessionParameters[$workloadInternalName].Keys)
                {
                    if (-not $authenticationParameters.ContainsKey($key))
                    {
                        $Script:MSCloudLoginConnectionProfile.$workloadInternalName.($key) = $optionalSessionParameters[$workloadInternalName][$key]
                    }
                }
            }
        }

        Invoke-MSCloudLoginAssistantConnectionLock -ConnectScript {
            switch ($Workload)
            {
                'AdminAPI'
                {
                    $Script:MSCloudLoginConnectionProfile.AdminAPI.Connect()
                    break
                }
                'Azure'
                {
                    $Script:MSCloudLoginConnectionProfile.Azure.Connect()
                    break
                }
                'AzureDevOPS'
                {
                    $Script:MSCloudLoginConnectionProfile.AzureDevOPS.Connect()
                    break
                }
                'DefenderForEndpoint'
                {
                    $Script:MSCloudLoginConnectionProfile.DefenderForEndpoint.Connect()
                    break
                }
                'EngageHub'
                {
                    $Script:MSCloudLoginConnectionProfile.EngageHub.Connect()
                    break
                }
                'ExchangeOnline'
                {
                    $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connect()
                    break
                }
                'Fabric'
                {
                    $Script:MSCloudLoginConnectionProfile.Fabric.Connect()
                    break
                }
                'Licensing'
                {
                    $Script:MSCloudLoginConnectionProfile.Licensing.Connect()
                    break
                }
                'O365Portal'
                {
                    $Script:MSCloudLoginConnectionProfile.O365Portal.Connect()
                    break
                }
                'MicrosoftGraph'
                {
                    $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connect()
                    break
                }
                'MicrosoftTeams'
                {
                    $Script:MSCloudLoginConnectionProfile.Teams.Connect()
                    break
                }
                'PnP'
                {
                    # Mark as disconnected if we are trying to connect to a different url then we previously connected to.
                    if ($Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl -ne $Url -or `
                            -not $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl -and `
                            $Url -or (-not $Url -and -not $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl))
                    {
                        Add-MSCloudLoginAssistantEvent -Message "Connecting to a different connection URL. Old URL: $($Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl), New URL: $Url" -Source $source
                        $ForceRefresh = $false
                        if ($Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl -ne $Url -and `
                            -not [System.String]::IsNullOrEmpty($url))
                        {
                            $ForceRefresh = $true
                        }
                        $Script:MSCloudLoginConnectionProfile.PnP.Connected = $false
                        $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl = $Url
                        $Script:MSCloudLoginConnectionProfile.PnP.Connect($ForceRefresh)
                    }
                    else
                    {
                        try
                        {
                            $contextUrl = (Get-PnPContext).Url
                            if ([System.String]::IsNullOrEmpty($url))
                            {
                                $Url = $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl
                                if (-not $Url.EndsWith('/') -and $contextUrl.EndsWith('/'))
                                {
                                    $Url += '/'
                                }
                            }
                            if ($contextUrl -ne $Url)
                            {
                                $ForceRefresh = $true
                                Add-MSCloudLoginAssistantEvent -Message "Connecting to a different context URL. Old URL: $contextUrl, New URL: $Url" -Source $source
                                $Script:MSCloudLoginConnectionProfile.PnP.Connected = $false
                                if ($url)
                                {
                                    $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl = $Url
                                }
                                else
                                {
                                    $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl = $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl
                                }
                                $Script:MSCloudLoginConnectionProfile.PnP.Connect($ForceRefresh)
                            }
                        }
                        catch
                        {
                            Add-MSCloudLoginAssistantEvent -Message "Couldn't acquire PnP Context to evaluate a URL change: $($_.Exception.Message)" -Source $source -EntryType 'Warning'
                        }
                    }

                    # If the AdminUrl is empty and a URL was provided, assume that the url
                    # provided is the admin center;
                    if (-not $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl -and $Url)
                    {
                        $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl = $Url
                    }
                    break
                }
                'PowerPlatforms'
                {
                    $Script:MSCloudLoginConnectionProfile.PowerPlatform.Connect()
                    break
                }
                'PowerPlatformREST'
                {
                    $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.Connect()
                    break
                }
                'SecurityComplianceCenter'
                {
                    $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.EnableSearchOnlySession = $EnableSearchOnlySession
                    $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Connect()
                    break
                }
                'SharePointOnlineREST'
                {
                    $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.ConnectionUrl = $Url
                    $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.Connect()

                    # If the AdminUrl is empty and a URL was provided, assume that the url
                    # provided is the admin center;
                    if (-not $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl -and $Url)
                    {
                        $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl = $Url
                    }
                    break
                }
                'Tasks'
                {
                    $Script:MSCloudLoginConnectionProfile.Tasks.Connect()
                    break
                }
            }
        }
    }
}

function Get-AuthenticationTypeFromParameters
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.Object]
        $AuthenticationObject
    )

    $authenticationType = ''
    if ($AuthenticationObject.ApplicationId -and $AuthenticationObject.TenantId -and $AuthenticationObject.CertificateThumbprint)
    {
        $authenticationType = 'ServicePrincipalWithThumbprint'
    }
    elseif ($AuthenticationObject.ApplicationId -and $AuthenticationObject.TenantId -and $AuthenticationObject.ApplicationSecret)
    {
        $authenticationType = 'ServicePrincipalWithSecret'
    }
    elseif ($AuthenticationObject.ApplicationId -and $AuthenticationObject.TenantId -and $AuthenticationObject.CertificatePath -and $AuthenticationObject.CertificatePassword)
    {
        $authenticationType = 'ServicePrincipalWithPath'
    }
    elseif ($AuthenticationObject.Credentials -and $AuthenticationObject.ApplicationId)
    {
        $authenticationType = 'CredentialsWithApplicationId'
    }
    elseif ($AuthenticationObject.Credentials -and $AuthenticationObject.TenantId)
    {
        $authenticationType = 'CredentialsWithTenantId'
    }
    elseif ($AuthenticationObject.Credentials)
    {
        $authenticationType = 'Credentials'
    }
    elseif ($AuthenticationObject.Identity)
    {
        $authenticationType = 'Identity'
    }
    elseif ($AuthenticationObject.AccessTokens -and -not [System.String]::IsNullOrEmpty($AuthenticationObject.TenantId))
    {
        $authenticationType = 'AccessTokens'
    }
    else
    {
        $authenticationType = 'Interactive'
    }

    return $authenticationType
}

<#
.SYNOPSIS
    This function returns the connection profile for a specific workload.
.DESCRIPTION
    This function returns the connection profile for a specific workload. A caller can use this function to get connection information for a specific workload.
.OUTPUTS
    Object (or $null). Get-MSCloudLoginConnectionProfile returns the connection profile for a specific workload or $null, if no connection profile exists.
.EXAMPLE
    Get-MSCloudLoginConnectionProfile -Workload 'MicrosoftGraph'
#>
function Get-MSCloudLoginConnectionProfile
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateSet('AdminAPI', 'Azure', 'AzureDevOPS', 'EngageHub', 'ExchangeOnline', 'Fabric', 'Licensing', `
                'O365Portal', 'SecurityComplianceCenter', 'PnP', 'PowerPlatforms', 'PowerPlatformREST', `
                'MicrosoftTeams', 'Teams', 'MicrosoftGraph', 'SharePointOnlineREST', 'Tasks', 'DefenderForEndpoint')]
        [System.String]
        $Workload
    )

    if ($Workload -eq 'MicrosoftTeams')
    {
        $Workload = 'Teams'
    }

    if ($null -ne $Script:MSCloudLoginConnectionProfile.$Workload)
    {
        return $Script:MSCloudLoginConnectionProfile.$Workload.Clone()
    }
}

<#
.SYNOPSIS
    This function resets the entire connection profile.
.DESCRIPTION
    This function resets the entire connection profile. It is used to disconnect all workloads and reset the connection profile.
.EXAMPLE
    Reset-MSCloudLoginConnectionProfileContext
#>
function Reset-MSCloudLoginConnectionProfileContext
{
    [CmdletBinding()]
    param (
        [Parameter()]
        [ValidateSet('AdminAPI', 'Azure', 'AzureDevOPS', 'EngageHub', 'ExchangeOnline', 'Fabric', 'Licensing', `
                'O365Portal', 'SecurityComplianceCenter', 'PnP', 'PowerPlatform', 'PowerPlatformREST', `
                'MicrosoftTeams', 'MicrosoftGraph', 'SharePointOnlineREST', 'Tasks', 'DefenderForEndpoint')]
        [System.String[]]
        $Workload
    )

    $fullReset = $false
    if ($Workload.Count -eq 0)
    {
        $workloads = $Script:MSCloudLoginConnectionProfile.PSObject.Properties.Name | Where-Object { $_ -notin @('CreatedTime', 'OrganizationName', 'Teams') }
        $workloads += 'MicrosoftTeams'
        $Workload = $workloads
        $fullReset = $true
    }

    $source = 'Reset-MSCloudLoginConnectionProfileContext'
    Add-MSCloudLoginAssistantEvent -Message 'Resetting connection profile' -Source $source
    foreach ($workloadToReset in $Workload)
    {
        if ($workloadToReset -eq 'MicrosoftTeams')
        {
            $workloadToReset = 'Teams'
        }
        $disconnectExists = $null -ne ($Script:MSCloudLoginConnectionProfile.$workloadToReset | Get-Member -Name 'Disconnect' -MemberType Method -ErrorAction SilentlyContinue)
        if ($disconnectExists)
        {
            try
            {
                $Script:MSCloudLoginConnectionProfile.$workloadToReset.Disconnect()
            }
            catch
            {
                Add-MSCloudLoginAssistantEvent -Message "Failed to disconnect workload {$workloadToReset}: $($_.Exception.Message)" -Source $source -EntryType 'Error'
            }
        }
        else
        {
            Add-MSCloudLoginAssistantEvent -Message "No disconnect method found for workload {$workloadToReset}. Operation ignored." -Source $source
        }
    }

    if ($fullReset)
    {
        $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
    }
}

<#
.Description
    This function creates a new entry in the MSCloudLoginAssistant event log, based on the provided information
.Functionality
    Internal
#>
function Add-MSCloudLoginAssistantEvent
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $Message,

        [Parameter(Mandatory = $true)]
        [System.String]
        $Source,

        [Parameter()]
        [ValidateSet('Error', 'Information', 'FailureAudit', 'SuccessAudit', 'Warning')]
        [System.String]
        $EntryType = 'Information',

        [Parameter()]
        [System.UInt32]
        $EventID = 1
    )

    if ($EntryType -eq 'Error')
    {
        Write-Verbose -Message "ERROR: [$Source] $Message"
    }
    else
    {
        Write-Verbose -Message "[$Source] $Message"
    }

    if (-not $Script:WriteToEventLog)
    {
        return
    }

    $logName = 'MSCloudLoginAssistant'

    try
    {
        try
        {
            $sourceExists = [System.Diagnostics.EventLog]::SourceExists($Source)
        }
        catch [System.Security.SecurityException]
        {
            Write-Warning -Message "MSCloudLoginAssistant - Access to an event log is denied. The message {$Message} from {$Source} will not be written to the event log."
            return
        }

        if ($sourceExists)
        {
            $sourceLogName = [System.Diagnostics.EventLog]::LogNameFromSourceName($Source, '.')
            if ($logName -ne $sourceLogName)
            {
                Write-Warning -Message "[ERROR] Specified source {$Source} already exists on log {$sourceLogName}"
                return
            }
        }
        else
        {
            try
            {
                [System.Diagnostics.EventLog]::CreateEventSource($Source, $logName)
            }
            catch [System.Security.SecurityException]
            {
                Write-Verbose -Message "[WARNING] Not all event logs could be searched. Source might exist in another event log."
            }
        }

        # Limit the size of the message. Maximum is about 32,766
        $outputMessage = $Message
        if ($outputMessage.Length -gt 32766)
        {
            $outputMessage = $outputMessage.Substring(0, 32766)
        }

        try
        {
            [System.Diagnostics.EventLog]::WriteEntry($Source, $outputMessage, $EntryType, $EventID)
        }
        catch
        {
            Write-Warning -Message "MSCloudLoginAssistant - Failed to save event: $_"
        }
    }
    catch
    {
        $messageText = "MSCloudLoginAssistant - Could not write to event log Source {$Source} EntryType {$EntryType} Message {$Message}. Error message $_"
        Write-Warning -Message $messageText
    }
}

<#
.SYNOPSIS
    This functions compares the authentication parameters for a change compared to the currently used parameters.
.DESCRIPTION
    This functions compares the authentication parameters for a change compared to the currently used parameters.
    It is used to determine if a new connection needs to be made.
.OUTPUTS
    Boolean. Compare-InputParametersForChange returns $true if something changed, $false otherwise.
.EXAMPLE
    Compare-InputParametersForChange -CurrentParamSet $PSBoundParameters
#>
function Compare-InputParametersForChange
{
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param (
        [Parameter()]
        [System.Collections.Hashtable]
        $CurrentParamSet,

        [Parameter()]
        [System.String[]]
        $LimitToKeys
    )

    $source = 'Compare-InputParametersForChange'

    if ($null -eq $Script:MSCloudLoginConnectionProfile)
    {
        return $true
    }

    if ($null -eq $CurrentParamSet)
    {
        $CurrentParamSet = @{}
    }

    $workload = $CurrentParamSet['Workload']
    $workloadInternalName = switch ($workload)
    {
        'MicrosoftTeams' { 'Teams' }
        'PowerPlatforms' { 'PowerPlatform' }
        default { $workload }
    }
    $workloadProfile = $Script:MSCloudLoginConnectionProfile.$workloadInternalName
    if ($null -eq $workloadProfile)
    {
        return $true
    }

    if ($null -eq $LimitToKeys -and $workloadProfile.RequestedAuthenticationType -ne $workloadProfile.AuthenticationType)
    {
        # Authentication type changed, so we need to reconnect
        Add-MSCloudLoginAssistantEvent -Message "Authentication type changed from {$($workloadProfile.AuthenticationType)} to {$($workloadProfile.RequestedAuthenticationType)}" -Source $source
        return $true
    }

    $desired = @{}
    foreach ($entry in $CurrentParamSet.GetEnumerator())
    {
        switch ($entry.Key)
        {
            'Credential'
            {
                $desired['Credentials'] = $entry.Value
            }
            'ExchangeOnlineCmdlets'
            {
                $desired['CmdletsToLoad'] = $entry.Value
            }
            'Url'
            {
                $desired['ConnectionUrl'] = $entry.Value
            }
            { $_ -in @('ApplicationId', 'TenantId', 'TenantGUID', 'ApplicationSecret',
                    'CertificateThumbprint', 'CertificatePath', 'CertificatePassword',
                    'Identity', 'AccessTokens', 'Endpoints', 'SubscriptionId',
                    'EnableSearchOnlySession') }
            {
                $desired[$entry.Key] = $entry.Value
            }
            # Everything else (Workload, UseModernAuth, CustomEnvironmentFileName, SkipModuleReload,
            # common parameters, ...) is intentionally not compared.
        }
    }

    # Active state: the same canonical keys, read from the workload profile.
    $active = @{}
    foreach ($key in @('Credentials', 'ApplicationId', 'TenantId', 'TenantGUID', 'ApplicationSecret',
            'CertificateThumbprint', 'CertificatePath', 'CertificatePassword',
            'Identity', 'AccessTokens', 'Endpoints'))
    {
        $active[$key] = $workloadProfile.$key
    }
    switch ($workloadInternalName)
    {
        'Azure'
        {
            $active['SubscriptionId'] = $workloadProfile.SubscriptionId
        }
        'ExchangeOnline'
        {
            $active['CmdletsToLoad'] = $workloadProfile.CmdletsToLoad
        }
        'SecurityComplianceCenter'
        {
            $active['EnableSearchOnlySession'] = $workloadProfile.EnableSearchOnlySession
        }
        { $_ -in @('PnP', 'SharePointOnlineREST') }
        {
            $active['ConnectionUrl'] = $workloadProfile.ConnectionUrl
        }
    }

    foreach ($key in @('SubscriptionId', 'CmdletsToLoad', 'ConnectionUrl', 'EnableSearchOnlySession'))
    {
        if ($desired.ContainsKey($key) -and -not $active.ContainsKey($key))
        {
            $desired.Remove($key)
        }
    }

    # Workload specific normalization to prevent false positives for the common
    # credentials-only Microsoft Graph connection pattern.
    if ($workloadInternalName -eq 'MicrosoftGraph')
    {
        if ($active['ApplicationId'] -eq '14d82eec-204b-4c2f-b7e8-296a70dab67e' -and `
                [System.String]::IsNullOrEmpty($desired['ApplicationId']))
        {
            # The default Microsoft Graph PowerShell app id is injected by the module itself.
            $active.Remove('ApplicationId')
        }
        if ($null -ne $active['Credentials'] -and `
                [System.String]::IsNullOrEmpty($desired['TenantId']) -and `
                $active['TenantId'] -eq ($active['Credentials'].UserName -split '@')[1])
        {
            # The TenantId was inferred from the credential UPN suffix.
            $active.Remove('TenantId')
        }
    }

    # When the caller only supplied session parameters (no identity parameters at all), the
    # identity keys of the active profile must not be compared. Otherwise, they would all show up as
    # 'present on one side only' and report a change on every single call.
    if ($null -ne $LimitToKeys)
    {
        foreach ($table in @($desired, $active))
        {
            foreach ($key in @($table.Keys))
            {
                if ($key -notin $LimitToKeys)
                {
                    $table.Remove($key)
                }
            }
        }
    }

    # Drop empty values so that absent / $null / '' / @() / $false are all equivalent on both sides.
    foreach ($table in @($desired, $active))
    {
        foreach ($key in @($table.Keys))
        {
            if (Test-MSCloudLoginParameterValueEmpty -Value $table[$key])
            {
                $table.Remove($key)
            }
        }
    }

    # Key-wise comparison over the union of keys: a key that is present on only one side counts as a change.
    $changedKeys = [System.Collections.Generic.List[string]]::new()
    $allKeys = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($key in $desired.Keys)
    {
        $null = $allKeys.Add($key)
    }
    foreach ($key in $active.Keys)
    {
        $null = $allKeys.Add($key)
    }

    foreach ($key in $allKeys)
    {
        if ($desired.ContainsKey($key) -ne $active.ContainsKey($key))
        {
            $changedKeys.Add($key)
        }
        elseif (-not (Test-MSCloudLoginParameterValueEqual -KeyName $key -Left $desired[$key] -Right $active[$key]))
        {
            $changedKeys.Add($key)
        }
    }

    if ($changedKeys.Count -eq 0)
    {
        # no differences were found
        return $false
    }

    # SECURITY: only the NAMES of the changed parameters are logged, never their values.
    Add-MSCloudLoginAssistantEvent -Message "Authentication parameters changed for workload {$workloadInternalName}: $($changedKeys -join ', ')" -Source $source
    return $true
}

function Get-SPOAdminUrl
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter()]
        [System.Management.Automation.PSCredential]
        $Credential
    )

    $source = 'Get-SPOAdminUrl'
    Add-MSCloudLoginAssistantEvent -Message 'Connection to Microsoft Graph is required to automatically determine SharePoint Online admin URL...' -Source $source

    # Only pass the credential through when one was actually provided, otherwise the
    # re-entrant Connect-M365Tenant call would be treated as an 'Interactive' request.
    $graphConnectParams = @{
        Workload = 'MicrosoftGraph'
    }
    if ($null -ne $Credential)
    {
        $graphConnectParams['Credential'] = $Credential
    }

    try
    {
        $result = Invoke-MgGraphRequest -Uri '/v1.0/sites/root' -ErrorAction SilentlyContinue
        $weburl = $result.webUrl
        if (-not $weburl)
        {
            Connect-M365Tenant @graphConnectParams
            $weburl = (Invoke-MgGraphRequest -Uri '/v1.0/sites/root').webUrl
        }
    }
    catch
    {
        Connect-M365Tenant @graphConnectParams
        try
        {
            $weburl = (Invoke-MgGraphRequest -Uri /v1.0/sites/root).webUrl
        }
        catch
        {
            if ((Assert-IsNonInteractiveShell) -eq $false)
            {
                # Only run interactive command when Exporting
                Add-MSCloudLoginAssistantEvent -Message 'Requesting access to read information about the domain' -Source $source
                Connect-MgGraph -Scopes Sites.Read.All -ErrorAction 'Stop'
                $weburl = (Invoke-MgGraphRequest -Uri /v1.0/sites/root).webUrl
            }
            else
            {
                if ($_.Exception.Message -eq 'Insufficient privileges to complete the operation.' -or `
                    $_.Exception.Message -like "*Forbidden*")
                {
                    throw "The Graph application does not have the correct permissions to access Domains. Make sure you run 'Connect-MgGraph -Scopes Sites.Read.All' first!"
                }
            }
        }
    }

    if ($null -eq $weburl)
    {
        throw 'Unable to retrieve SPO Admin URL. Please check connectivity and if you have the Sites.Read.All permission.'
    }

    $spoAdminUrl = $webUrl -replace '^https:\/\/(\w*)\.', 'https://$1-admin.'
    Add-MSCloudLoginAssistantEvent -Message "SharePoint Online admin URL is $spoAdminUrl" -Source $source
    return $spoAdminUrl
}

function Get-MSCloudLoginAccessToken
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param(
        [Parameter(Mandatory = $True)]
        [System.String]
        $ConnectionUri,

        [Parameter(Mandatory = $True)]
        [System.String]
        $AuthorizationUrl,

        [Parameter(Mandatory = $True)]
        [System.String]
        $AzureADAuthorizationEndpointUri,

        [Parameter(Mandatory = $True)]
        [System.String]
        $ApplicationId,

        [Parameter(Mandatory = $True)]
        [System.String]
        $TenantId,

        [Parameter(Mandatory = $True)]
        [System.String]
        $CertificateThumbprint
    )

    $source = 'Get-MSCloudLoginAccessToken'

    try
    {
        Add-MSCloudLoginAssistantEvent -Message 'Connecting by endpoints URI' -Source $source
        $response = Get-AuthToken -AuthorizationUrl $AuthorizationUrl `
            -TokenEndpoint $AzureADAuthorizationEndpointUri `
            -Scope $ConnectionUri `
            -ClientId $ApplicationId `
            -TenantId $TenantId `
            -CertificateThumbprint $CertificateThumbprint
        return $response.access_token
    }
    catch
    {
        Add-MSCloudLoginAssistantEvent -Message $_ -Source $source -EntryType Error
        throw
    }
}

function Get-CloudEnvironmentInfo
{
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter()]
        [System.Management.Automation.PSCredential]
        $Credentials,

        [Parameter()]
        [System.String]
        $ApplicationId,

        [Parameter()]
        [System.String]
        $TenantId,

        [Parameter()]
        [System.String]
        $ApplicationSecret,

        [Parameter()]
        [System.String]
        $CertificateThumbprint,

        [Parameter()]
        [switch]
        $Identity
    )

    $source = 'Get-CloudEnvironmentInfo'
    Add-MSCloudLoginAssistantEvent -Message 'Retrieving Environment Details' -Source $source

    if ($null -ne $Credentials)
    {
        $tenantName = Get-MSCloudLoginTenantDomainFromCredentials -Credentials $Credentials
    }
    elseif (-not [string]::IsNullOrEmpty($TenantId))
    {
        $tenantName = $TenantId
    }
    elseif ($Identity.IsPresent)
    {
        return
    }
    else
    {
        throw 'TenantId or Credentials must be provided'
    }
    ## endpoint will work with TenantId or tenantName
    switch ($tenantName)
    {
        { $_ -like '*.onsovcloud.de*' }
        {
            $loginEndpoint = 'login.sovcloud-identity.de'
            break
        }
        { $_ -like '*.onsovcloud.fr*' }
        {
            $loginEndpoint = 'login.sovcloud-identity.fr'
            break
        }
        default
        {
            $loginEndpoint = "login.microsoftonline.com"
        }
    }
    $response = Invoke-WebRequest -Uri "https://$loginEndpoint/$tenantName/v2.0/.well-known/openid-configuration" -Method Get -UseBasicParsing

    $content = $response.Content
    $result = ConvertFrom-Json $content
    return $result
}

function Get-MSCloudLoginOrganizationName
{
    param(
        [Parameter()]
        [System.String]
        $ApplicationId,

        [Parameter()]
        [System.String]
        $TenantId,

        [Parameter()]
        [System.String]
        $CertificateThumbprint,

        [Parameter()]
        [System.String]
        $ApplicationSecret,

        [Parameter()]
        [switch]
        $Identity,

        [Parameter()]
        [System.String[]]
        $AccessTokens
    )

    $source = 'Get-MSCloudLoginOrganizationName'
    try
    {
        if (-not [string]::IsNullOrEmpty($ApplicationId) -and -not [System.String]::IsNullOrEmpty($CertificateThumbprint))
        {
            Connect-M365Tenant -Workload MicrosoftGraph -ApplicationId $ApplicationId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint
        }
        elseif (-not [string]::IsNullOrEmpty($ApplicationId) -and -not [System.String]::IsNullOrEmpty($ApplicationSecret))
        {
            Connect-M365Tenant -Workload MicrosoftGraph -ApplicationId $ApplicationId -TenantId $TenantId -ApplicationSecret $ApplicationSecret
        }
        elseif ($Identity.IsPresent)
        {
            Connect-M365Tenant -Workload MicrosoftGraph -Identity -TenantId $TenantId
        }
        elseif ($null -ne $AccessTokens)
        {
            Connect-M365Tenant -Workload MicrosoftGraph -AccessTokens $AccessTokens
        }
        $domain = (Invoke-MgGraphRequest -Method GET -Uri "/v1.0/domains" -ErrorAction Stop).value | Where-Object { $_.IsInitial -eq $True }

        if ($null -ne $domain)
        {
            return $domain.Id
        }
    }
    catch
    {
        if ([System.String]::IsNullOrEmpty($TenantId))
        {
            Add-MSCloudLoginAssistantEvent -Message "Couldn't get domain and no TenantId was provided as fallback: $($_.Exception.Message)" -Source $source -EntryType 'Error'
            throw
        }
        Add-MSCloudLoginAssistantEvent -Message "Couldn't get domain ($($_.Exception.Message)). Using TenantId instead" -Source $source
        return $TenantId
    }
}

function Assert-IsNonInteractiveShell
{
    # Test each Arg for match of abbreviated '-NonInteractive' command.
    $NonInteractive = [Environment]::GetCommandLineArgs() | Where-Object { $_ -like '-NonI*' }

    if ([Environment]::UserInteractive -and -not $NonInteractive)
    {
        # We are in an interactive shell.
        return $false
    }

    return $true
}

function ConvertTo-Base64Url {
    [CmdletBinding()]
    param(
        [byte[]] $bytes
    )

    [System.Convert]::ToBase64String($bytes).TrimEnd('=') | ForEach-Object { $_.Replace('+', '-').Replace('/', '_') }
}

function Get-AuthToken {
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    param (
        [Parameter(Mandatory = $true,  ParameterSetName = "ClientSecret")]
        [Parameter(Mandatory = $true,  ParameterSetName = "CertificateThumbprint")]
        [Parameter(Mandatory = $true,  ParameterSetName = "CertificatePath")]
        [Parameter(Mandatory = $true,  ParameterSetName = "Default")]
        [Parameter(Mandatory = $true,  ParameterSetName = "Device")]
        [Parameter(Mandatory = $false, ParameterSetName = "Identity")]
        [System.String]
        $AuthorizationUrl,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $Credentials,

        [Parameter()]
        [System.String]
        $TenantId,

        [Parameter()]
        [System.String]
        $ClientId,

        [Parameter(ParameterSetName = "ClientSecret")]
        [System.String]
        $ClientSecret,

        [Parameter(ParameterSetName = "CertificateThumbprint")]
        [System.String]
        $CertificateThumbprint,

        [Parameter(ParameterSetName = "CertificatePath")]
        [SecureString]
        $CertificatePassword,

        [Parameter(ParameterSetName = "CertificatePath")]
        [System.String]
        $CertificatePath,

        [Parameter(ParameterSetName = "Device")]
        [switch]
        $DeviceCode,

        [Parameter(ParameterSetName = "Identity")]
        [switch]
        $Identity,

        [Parameter()]
        [System.String]
        $RefreshToken,

        [Parameter(Mandatory = $false, ParameterSetName = "ClientSecret")]
        [Parameter(Mandatory = $false, ParameterSetName = "CertificateThumbprint")]
        [Parameter(Mandatory = $false, ParameterSetName = "CertificatePath")]
        [Parameter(Mandatory = $false, ParameterSetName = "Default")]
        [Parameter(Mandatory = $false, ParameterSetName = "Device")]
        [Parameter(Mandatory = $true,  ParameterSetName = "Identity")]
        [ValidateNotNullOrEmpty()]
        [System.String]
        $Resource,

        [Parameter()]
        [System.String]
        $Scope,

        [Parameter()]
        [System.String]
        $TokenEndpoint
    )

    if ($Identity.IsPresent) {
        $accessToken = ""
        if ($env:AZUREPS_HOST_ENVIRONMENT -like 'AzureAutomation*')
        {
            $url = $env:IDENTITY_ENDPOINT
            $headers = @{
                'Metadata' = $true
                'X-IDENTITY-HEADER' = $env:IDENTITY_HEADER
            }
            $body = @{
                resource = $Resource
            }
            $oauth2 = Invoke-RestMethod $url -Method 'POST' -Headers $headers -ContentType 'application/x-www-form-urlencoded' -Body $body
            $accessToken = $oauth2.access_token
        }
        elseif ('http://localhost:40342' -eq $env:IMDS_ENDPOINT)
        {
            # Get endpoint for Azure Arc Connected Device
            $apiVersion = '2020-06-01'
            $endpoint = '{0}?resource={1}&api-version={2}' -f $env:IDENTITY_ENDPOINT, $Resource, $apiVersion
            $secretFile = ''
            try
            {
                # This request is expected to fail with a 401 whose WWW-Authenticate header
                # points at the secret file used for the authenticated retry below.
                $null = Invoke-WebRequest -Method GET -Uri $endpoint -Headers @{
                    Metadata = $true
                } -UseBasicParsing
            }
            catch
            {
                $wwwAuthHeader = $_.Exception.Response.Headers['WWW-Authenticate']
                if ($wwwAuthHeader -match 'Basic realm=.+')
                {
                    $secretFile = ($wwwAuthHeader -split 'Basic realm=')[1]
                }
            }

            if ([System.String]::IsNullOrEmpty($secretFile))
            {
                throw "Unable to determine the Azure Arc managed identity secret file: the challenge request to '$endpoint' did not return the expected WWW-Authenticate header."
            }
            $secret = Get-Content -Raw $secretFile
            $response = Invoke-WebRequest -Method GET -Uri $endpoint -Headers @{
                Metadata = $true
                Authorization = "Basic $secret"
            } -UseBasicParsing

            if ($response)
            {
                $accessToken = (ConvertFrom-Json -InputObject $response.Content).access_token
            }
        }
        else
        {
            # Get correct endpoint for AzureVM
            $oauth2 = Invoke-RestMethod -Uri "http://169.254.169.254/metadata/identity/oauth2/token?api-version=2018-02-01&resource=$Resource" -Headers @{
                Metadata = $true
            }
            $accessToken = $oauth2.access_token
        }
        return $accessToken
    }

    $useResource = $PSBoundParameters.ContainsKey('Resource') -and $Resource
    if (-not [System.String]::IsNullOrEmpty($TokenEndpoint)) {
        # Custom environments provide the full token endpoint directly.
        $tokenEndpoint = $TokenEndpoint
    } elseif ($useResource) {
        $tokenEndpoint = "$AuthorizationUrl/$TenantId/oauth2/token"
    } else {
        $tokenEndpoint = "$AuthorizationUrl/$TenantId/oauth2/v2.0/token"
    }

    if ($ClientSecret -or $CertificatePath -or $CertificateThumbprint) {
        if ($CertificatePath) {
            $certificate = Get-MSCloudLoginCertificate -CertificatePath $CertificatePath -CertificatePassword $CertificatePassword
        }

        if ($CertificateThumbprint) {
            $certificate = Get-MSCloudLoginCertificate -CertificateThumbprint $CertificateThumbprint
        }

        if ($useResource) {
            $body = @{
                client_id = $ClientId
                resource = $Resource
                grant_type = 'client_credentials'
            }
        } else {
            $body = @{
                client_id = $ClientId
                scope = $Scope
                grant_type = 'client_credentials'
            }
        }

        if ($ClientSecret) {
            $body.client_secret = $ClientSecret
        } elseif ($certificate) {
            $now = (Get-Date).ToUniversalTime()
            $header = @{
                alg = 'RS256'
                typ = 'JWT'
            }

            if ($CertificateThumbprint -or $CertificatePath) {
                # RFC 7515 4.1.7: x5t is the base64url encoding of the certificate's SHA-1 hash.
                $header.Add('x5t', (ConvertTo-Base64Url -Bytes $certificate.GetCertHash()))
            }
            $payload = @{
                aud = $tokenEndpoint
                iss = $ClientId
                sub = $ClientId
                jti = [guid]::NewGuid().Guid
                nbf = [int][Math]::Floor(($now.AddMinutes(-5) - (Get-Date '1970-01-01Z').ToUniversalTime()).TotalSeconds)
                exp = [int][Math]::Floor(($now.AddMinutes(10) - (Get-Date '1970-01-01Z').ToUniversalTime()).TotalSeconds)
            }
            $headerEnc = ConvertTo-Base64Url -Bytes ([System.Text.Encoding]::UTF8.GetBytes((ConvertTo-Json $header -Compress)))
            $payloadEnc = ConvertTo-Base64Url -Bytes ([System.Text.Encoding]::UTF8.GetBytes((ConvertTo-Json $payload -Compress)))
            $unsigned = "$headerEnc.$payloadEnc"
            $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($certificate)
            $signature = $rsa.SignData([System.Text.Encoding]::UTF8.GetBytes($unsigned), [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
            $signed = "$unsigned.$(ConvertTo-Base64Url -Bytes $signature)"
            $body.client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
            $body.client_assertion = $signed

            if ($CertificateThumbprint) {
                $headers = @{
                    Authorization = "Bearer $($body.client_assertion)"
                }
            }
        }

        if ($headers) {
            return Invoke-RestMethod -Method Post -Uri $tokenEndpoint -Body $body -ContentType 'application/x-www-form-urlencoded' -Headers $headers
        } else {
            return Invoke-RestMethod -Method Post -Uri $tokenEndpoint -Body $body -ContentType 'application/x-www-form-urlencoded'
        }
    }

    if ($RefreshToken) {
        if ($useResource) {
            $body = @{
                client_id     = $ClientId
                resource      = $Resource
                grant_type    = 'refresh_token'
                refresh_token = $RefreshToken
            }
        } else {
            $body = @{
                client_id     = $ClientId
                scope         = $Scope
                grant_type    = 'refresh_token'
                refresh_token = $RefreshToken
            }
        }

        return Invoke-RestMethod -Method Post -Uri $tokenEndpoint -Body $body -ContentType 'application/x-www-form-urlencoded'
    }

    if ($Credentials -and -not $DeviceCode) {
        if ($useResource) {
            $body = @{
                client_id  = $ClientId
                resource   = $Resource
                grant_type = 'password'
                username   = $Credentials.UserName
                password   = $Credentials.GetNetworkCredential().Password
            }
        } else {
            $body = @{
                client_id  = $ClientId
                scope      = $Scope
                grant_type = 'password'
                username   = $Credentials.UserName
                password   = $Credentials.GetNetworkCredential().Password
            }
        }

        return Invoke-RestMethod -Method Post -Uri $tokenEndpoint -Body $body -ContentType 'application/x-www-form-urlencoded'
    }

    if ($DeviceCode) {
        $deviceEndpoint = "$AuthorizationUrl/$TenantId/oauth2/v2.0/devicecode"
        $deviceBody = @{
            client_id = $ClientId
            scope = $(if ($useResource) { $Resource } else { $Scope } )
        }
        $deviceCodeResponse = Invoke-RestMethod -Method Post -Uri $deviceEndpoint -Body $deviceBody -ContentType 'application/x-www-form-urlencoded'

        Write-Verbose -Message "`n$($deviceCodeResponse.message)" -Verbose
        $pollBody = @{
            grant_type = 'urn:ietf:params:oauth:grant-type:device_code'
            client_id = $ClientId
            device_code = $deviceCodeResponse.device_code
        }

        $timeoutTimer = [System.Diagnostics.Stopwatch]::StartNew()
        do {
            if ($timeoutTimer.Elapsed.TotalSeconds -gt 300)
            {
                throw 'Login timed out, please try again.'
            }
            Start-Sleep -Seconds $deviceCodeResponse.interval
            try {
                $result = Invoke-RestMethod -Method Post -Uri $tokenEndpoint -Body $pollBody -ContentType 'application/x-www-form-urlencoded'
            } catch {
                # Keep polling only while authorization is still pending. Terminal OAuth errors
                # (access_denied, expired_token, invalid_client, ...) must surface immediately
                # instead of timing out after 5 minutes with a misleading message.
                $oauthError = $null
                try {
                    $oauthError = ($_.ErrorDetails.Message | ConvertFrom-Json).error
                } catch {
                    # Not an OAuth error payload (e.g. a network failure) - treat as terminal.
                    $oauthError = $null
                }
                if ($oauthError -notin @('authorization_pending', 'slow_down')) {
                    throw
                }
                $result = $null
            }
        } while ($null -eq $result)
        return $result
    }

    $verifierBytes = New-Object byte[] 32
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    $rng.GetBytes($verifierBytes)
    $codeVerifier = [System.Convert]::ToBase64String($verifierBytes).TrimEnd('=')
    $codeVerifier = $codeVerifier.Replace('+', '-').Replace('/', '_')
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $challengeBytes = $sha.ComputeHash([System.Text.Encoding]::ASCII.GetBytes($codeVerifier))
    $codeChallenge = [System.Convert]::ToBase64String($challengeBytes).TrimEnd('=')
    $codeChallenge = $codeChallenge.Replace('+', '-').Replace('/', '_')
    $redirectUri = "http://localhost:8400/"
    $authorizeUrl = "$AuthorizationUrl/$TenantId/oauth2/v2.0/authorize?client_id=$ClientId&response_type=code&redirect_uri=$([System.Uri]::EscapeDataString($redirectUri))&response_mode=query&scope=$([System.Uri]::EscapeDataString($Scope))&code_challenge=$codeChallenge&code_challenge_method=S256"

    $listener = [System.Net.HttpListener]::new()
    $listener.Prefixes.Add($redirectUri)
    $listener.Start()
    try {
        try {
            if ($IsWindows) {
                Start-Process $authorizeUrl
            } else {
                Write-Host "Open $authorizeUrl in your browser to authenticate"
            }
        } catch {
            Write-Verbose "Unable to automatically open browser: $($_.Exception.Message)"
            Write-Host "Open $authorizeUrl in your browser to authenticate"
        }
        $context = $listener.GetContext()
        $query = [System.Web.HttpUtility]::ParseQueryString($context.Request.Url.Query)
        $code = $query['code']
        $responseBytes = [System.Text.Encoding]::UTF8.GetBytes('<html><body>You may close this window.</body></html>')
        $context.Response.ContentLength64 = $responseBytes.Length
        $context.Response.OutputStream.Write($responseBytes, 0, $responseBytes.Length)
        $context.Response.OutputStream.Close()
    } finally {
        $listener.Stop()
        $listener.Close()
    }

    $body = @{
        client_id     = $ClientId
        scope         = $Scope
        grant_type    = 'authorization_code'
        code          = $code
        redirect_uri  = $redirectUri
        code_verifier = $codeVerifier
    }
    $response = Invoke-RestMethod -Method Post -Uri $tokenEndpoint -Body $body -ContentType 'application/x-www-form-urlencoded'

    return $response
}

<#
.SYNOPSIS
    Generic connection function for REST API-based workloads that use OAuth tokens.

.DESCRIPTION
    This function handles the common authentication patterns for workloads that connect
    via REST APIs and OAuth tokens. It supports Credentials, Credentials with MFA (DeviceCode),
    ServicePrincipalWithSecret, ServicePrincipalWithThumbprint, ServicePrincipalWithPath,
    Identity, and AccessTokens authentication methods.

.PARAMETER WorkloadName
    The name of the workload being connected to (used for logging and accessing connection profile).

.PARAMETER AuthorizationUrl
    The OAuth authorization URL endpoint.

.PARAMETER Scope
    The OAuth scope to request (for v2.0 endpoints).

.PARAMETER ClientId
    The client/application ID for delegated auth flows.

.PARAMETER SupportedAuthMethods
    An array of supported authentication methods for this workload.

.PARAMETER TokenExpireCheckMinutes
    Number of minutes before token expiration to trigger renewal. Default is 50.

.EXAMPLE
    Connect-MSCloudLoginRESTWorkload -WorkloadName 'AdminAPI' -AuthorizationUrl $authUrl -Scope $scope -ClientId $clientId
#>
function Connect-MSCloudLoginRESTWorkload
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.String]
        $WorkloadName,

        [Parameter()]
        [System.String]
        $AuthorizationUrl,

        [Parameter()]
        [System.String]
        $Scope,

        [Parameter()]
        [System.String]
        $ClientId,

        [Parameter()]
        [System.String[]]
        $SupportedAuthMethods = @('Credentials', 'CredentialsWithApplicationId', 'CredentialsWithTenantId', 'ServicePrincipalWithThumbprint', 'ServicePrincipalWithSecret', 'ServicePrincipalWithPath', 'Identity', 'AccessTokens'),

        [Parameter()]
        [System.Int32]
        $TokenExpireCheckMinutes = 50
    )

    $InformationPreference = 'SilentlyContinue'
    $ProgressPreference = 'SilentlyContinue'
    $source = "Connect-MSCloudLoginRESTWorkload ($WorkloadName)"

    $workloadProfile = $Script:MSCloudLoginConnectionProfile.$WorkloadName
    $authType = $workloadProfile.AuthenticationType

    # Token expiration check for applicable auth types
    if (Test-MSCloudLoginConnectionReusable -WorkloadProfile $workloadProfile `
            -TokenExpirationMinutes $TokenExpireCheckMinutes `
            -Source $source)
    {
        return
    }

    # Validate authentication method is supported
    if ($authType -notin $SupportedAuthMethods)
    {
        throw "Authentication method '$authType' is not supported for workload '$WorkloadName'. Supported methods: $($SupportedAuthMethods -join ', ')"
    }

    try
    {
        # Determine TenantId
        $tenantId = $workloadProfile.TenantId
        if ([System.String]::IsNullOrEmpty($tenantId) -and $null -ne $workloadProfile.Credentials)
        {
            $tenantId = Get-MSCloudLoginTenantDomainFromCredentials -Credentials $workloadProfile.Credentials
        }

        $accessToken = $null
        $useMFA = $false

        switch ($authType)
        {
            { $_ -in @('Credentials', 'CredentialsWithApplicationId', 'CredentialsWithTenantId') }
            {
                Add-MSCloudLoginAssistantEvent -Message 'Attempting to connect with user credentials' -Source $source
                $authParams = @{
                    AuthorizationUrl = $AuthorizationUrl
                    Credentials      = $workloadProfile.Credentials
                    TenantId         = $tenantId
                    ClientId         = if ($workloadProfile.ApplicationId) { $workloadProfile.ApplicationId } else { $ClientId }
                    Scope            = $Scope
                }

                try
                {
                    $tokenResponse = Get-AuthToken @authParams
                    $accessToken = "$($tokenResponse.token_type) $($tokenResponse.access_token)"
                }
                catch
                {
                    if ((Test-MSCloudLoginMFARequiredError -ErrorRecord $_))
                    {
                        Add-MSCloudLoginAssistantEvent -Message 'Account requires MFA, using device code flow' -Source $source
                        $authParams.DeviceCode = $true
                        $tokenResponse = Get-AuthToken @authParams
                        $accessToken = "$($tokenResponse.token_type) $($tokenResponse.access_token)"
                        $useMFA = $true
                    }
                    else
                    {
                        throw
                    }
                }
            }

            'ServicePrincipalWithThumbprint'
            {
                Add-MSCloudLoginAssistantEvent -Message "Attempting to connect using certificate thumbprint" -Source $source
                $authParams = @{
                    AuthorizationUrl      = $AuthorizationUrl
                    CertificateThumbprint = $workloadProfile.CertificateThumbprint
                    TenantId              = $tenantId
                    ClientId              = $workloadProfile.ApplicationId
                    Scope                 = $Scope
                }

                $tokenResponse = Get-AuthToken @authParams
                $accessToken = "$($tokenResponse.token_type) $($tokenResponse.access_token)"
            }

            'ServicePrincipalWithSecret'
            {
                Add-MSCloudLoginAssistantEvent -Message 'Attempting to connect using application secret' -Source $source
                $authParams = @{
                    AuthorizationUrl = $AuthorizationUrl
                    ClientSecret     = $workloadProfile.ApplicationSecret
                    TenantId         = $tenantId
                    ClientId         = $workloadProfile.ApplicationId
                    Scope            = $Scope
                }

                $tokenResponse = Get-AuthToken @authParams
                $accessToken = "$($tokenResponse.token_type) $($tokenResponse.access_token)"
            }

            'ServicePrincipalWithPath'
            {
                Add-MSCloudLoginAssistantEvent -Message 'Attempting to connect using certificate path' -Source $source
                $authParams = @{
                    AuthorizationUrl    = $AuthorizationUrl
                    CertificatePath     = $workloadProfile.CertificatePath
                    CertificatePassword = $workloadProfile.CertificatePassword
                    TenantId            = $tenantId
                    ClientId            = $workloadProfile.ApplicationId
                    Scope               = $Scope
                }

                $tokenResponse = Get-AuthToken @authParams
                $accessToken = "$($tokenResponse.token_type) $($tokenResponse.access_token)"
            }

            'Identity'
            {
                Add-MSCloudLoginAssistantEvent -Message 'Attempting to connect using Managed Identity' -Source $source
                $resourceValue = $Scope -replace '/\.default$', ''
                $tokenValue = Get-AuthToken -Resource $resourceValue -Identity
                $accessToken = "Bearer $tokenValue"
            }

            'AccessTokens'
            {
                Add-MSCloudLoginAssistantEvent -Message 'Using provided access token' -Source $source
                $providedToken = Get-MSCloudLoginAccessTokenValue -Token $workloadProfile.AccessTokens[0]
                $accessToken = if ($providedToken -like 'Bearer *') { $providedToken } else { "Bearer $providedToken" }
            }
        }

        # Set the access token and connection state
        $workloadProfile.AccessToken = $accessToken
        $workloadProfile.CompleteConnection($useMFA)

        Add-MSCloudLoginAssistantEvent -Message "Successfully connected to $WorkloadName" -Source $source
    }
    catch
    {
        $workloadProfile.Connected = $false
        Add-MSCloudLoginAssistantEvent -Message "Failed to connect to ${WorkloadName}: $($_.Exception.Message)" -Source $source -EntryType 'Error'
        throw
    }
}

<#
.SYNOPSIS
    Generic disconnect function for REST API-based workloads.

.DESCRIPTION
    This function handles the disconnect logic for workloads that use REST APIs.

.PARAMETER WorkloadName
    The name of the workload to disconnect from.
#>
function Disconnect-MSCloudLoginRESTWorkload
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.String]
        $WorkloadName
    )

    $source = "Disconnect-MSCloudLoginRESTWorkload ($WorkloadName)"
    $workloadProfile = $Script:MSCloudLoginConnectionProfile.$WorkloadName

    if ($workloadProfile.Connected)
    {
        Add-MSCloudLoginAssistantEvent -Message "Attempting to disconnect from $WorkloadName" -Source $source
        $workloadProfile.Connected = $false
        $workloadProfile.AccessToken = $null
        Add-MSCloudLoginAssistantEvent -Message "Successfully disconnected from $WorkloadName" -Source $source
    }
    else
    {
        Add-MSCloudLoginAssistantEvent -Message "No connections to $WorkloadName were found" -Source $source
    }
}
