$Script:WriteToEventLog = $env:MSCLOUDLOGINASSISTANT_WRITETOEVENTLOG -eq 'true'

. "$PSScriptRoot\ConnectionProfile.ps1"
$privateModules = Get-ChildItem -Path "$PSScriptRoot\Workloads" -Filter '*.ps1' -Recurse
foreach ($module in $privateModules)
{
    Write-Verbose "Importing workload $($module.FullName)"
    . $module.FullName
}

$requiredModules = @(
    'Microsoft.Graph.Beta.Identity.DirectoryManagement'
)
foreach ($module in $requiredModules)
{
    if (-not (Get-Module -Name $module -ListAvailable))
    {
        throw "The module $module is required to be installed. Please install the module and try again."
    }
}

function Connect-M365Tenant
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateSet('AdminAPI', 'Azure', 'AzureDevOPS', 'ExchangeOnline', 'Fabric', `
                'SecurityComplianceCenter', 'PnP', 'PowerPlatforms', `
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
        [System.Boolean]
        $SkipModuleReload = $false,

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
        [ValidateScript(
            { $Workload -eq 'ExchangeOnline' }
        )]
        [System.String[]]
        $ExchangeOnlineCmdlets = @()
    )

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

    if ($null -eq $Script:MSCloudLoginConnectionProfile)
    {
        $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
    }
    # Only validate the parameters if we are not already connected
    elseif ( $Script:MSCloudLoginConnectionProfile.$workloadInternalName.Connected `
            -and (Compare-InputParametersForChange -CurrentParamSet $PSBoundParameters))
    {
        Add-MSCloudLoginAssistantEvent -Message 'Resetting connection profile' -Source $source
        $Script:MSCloudLoginConnectionProfile.$workloadInternalName.Connected = $false
    }

    Add-MSCloudLoginAssistantEvent -Message "Trying to connect to platform {$Workload}" -Source $source
    switch ($Workload)
    {
        'AdminAPI'
        {
            $Script:MSCloudLoginConnectionProfile.AdminAPI.Credentials = $Credential
            $Script:MSCloudLoginConnectionProfile.AdminAPI.ApplicationId = $ApplicationId
            $Script:MSCloudLoginConnectionProfile.AdminAPI.ApplicationSecret = $ApplicationSecret
            $Script:MSCloudLoginConnectionProfile.AdminAPI.TenantId = $TenantId
            $Script:MSCloudLoginConnectionProfile.AdminAPI.CertificateThumbprint = $CertificateThumbprint
            $Script:MSCloudLoginConnectionProfile.AdminAPI.AccessTokens = $AccessTokens
            $Script:MSCloudLoginConnectionProfile.AdminAPI.Endpoints = $Endpoints
            $Script:MSCloudLoginConnectionProfile.AdminAPI.Connected = $false
            $Script:MSCloudLoginConnectionProfile.AdminAPI.Connect()
        }
        'Azure'
        {
            $Script:MSCloudLoginConnectionProfile.Azure.Credentials = $Credential
            $Script:MSCloudLoginConnectionProfile.Azure.ApplicationId = $ApplicationId
            $Script:MSCloudLoginConnectionProfile.Azure.ApplicationSecret = $ApplicationSecret
            $Script:MSCloudLoginConnectionProfile.Azure.TenantId = $TenantId
            $Script:MSCloudLoginConnectionProfile.Azure.CertificateThumbprint = $CertificateThumbprint
            $Script:MSCloudLoginConnectionProfile.Azure.AccessTokens = $AccessTokens
            $Script:MSCloudLoginConnectionProfile.Azure.Endpoints = $Endpoints
            $Script:MSCloudLoginConnectionProfile.Azure.Connected = $false
            $Script:MSCloudLoginConnectionProfile.Azure.Connect()
        }
        'AzureDevOPS'
        {
            $Script:MSCloudLoginConnectionProfile.AzureDevOPS.Credentials = $Credential
            $Script:MSCloudLoginConnectionProfile.AzureDevOPS.ApplicationId = $ApplicationId
            $Script:MSCloudLoginConnectionProfile.AzureDevOPS.ApplicationSecret = $ApplicationSecret
            $Script:MSCloudLoginConnectionProfile.AzureDevOPS.TenantId = $TenantId
            $Script:MSCloudLoginConnectionProfile.AzureDevOPS.CertificateThumbprint = $CertificateThumbprint
            $Script:MSCloudLoginConnectionProfile.AzureDevOPS.AccessTokens = $AccessTokens
            $Script:MSCloudLoginConnectionProfile.AzureDevOPS.Identity = $Identity
            $Script:MSCloudLoginConnectionProfile.AzureDevOPS.Endpoints = $Endpoints
            $Script:MSCloudLoginConnectionProfile.AzureDevOPS.Connect()
        }
        'DefenderForEndpoint'
        {
            $Script:MSCloudLoginConnectionProfile.DefenderForEndpoint.Credentials = $Credential
            $Script:MSCloudLoginConnectionProfile.DefenderForEndpoint.ApplicationId = $ApplicationId
            $Script:MSCloudLoginConnectionProfile.DefenderForEndpoint.ApplicationSecret = $ApplicationSecret
            $Script:MSCloudLoginConnectionProfile.DefenderForEndpoint.TenantId = $TenantId
            $Script:MSCloudLoginConnectionProfile.DefenderForEndpoint.CertificateThumbprint = $CertificateThumbprint
            $Script:MSCloudLoginConnectionProfile.DefenderForEndpoint.AccessTokens = $AccessTokens
            $Script:MSCloudLoginConnectionProfile.DefenderForEndpoint.Identity = $Identity
            $Script:MSCloudLoginConnectionProfile.DefenderForEndpoint.Endpoints = $Endpoints
            $Script:MSCloudLoginConnectionProfile.DefenderForEndpoint.Connect()
        }
        'ExchangeOnline'
        {
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Credentials = $Credential
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ApplicationId = $ApplicationId
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ApplicationSecret = $ApplicationSecret
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.TenantId = $TenantId
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CertificateThumbprint = $CertificateThumbprint
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.SkipModuleReload = $SkipModuleReload
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.AccessTokens = $AccessTokens
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Identity = $Identity
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Endpoints = $Endpoints
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CmdletsToLoad = $ExchangeOnlineCmdlets
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connect()
        }
        'Fabric'
        {
            $Script:MSCloudLoginConnectionProfile.Fabric.Credentials = $Credential
            $Script:MSCloudLoginConnectionProfile.Fabric.ApplicationId = $ApplicationId
            $Script:MSCloudLoginConnectionProfile.Fabric.ApplicationSecret = $ApplicationSecret
            $Script:MSCloudLoginConnectionProfile.Fabric.TenantId = $TenantId
            $Script:MSCloudLoginConnectionProfile.Fabric.CertificateThumbprint = $CertificateThumbprint
            $Script:MSCloudLoginConnectionProfile.Fabric.AccessTokens = $AccessTokens
            $Script:MSCloudLoginConnectionProfile.Fabric.Identity = $Identity
            $Script:MSCloudLoginConnectionProfile.Fabric.Endpoints = $Endpoints
            $Script:MSCloudLoginConnectionProfile.Fabric.Connect()
        }
        'MicrosoftGraph'
        {
            $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Credentials = $Credential
            $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ApplicationId = $ApplicationId
            $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ApplicationSecret = $ApplicationSecret
            $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.TenantId = $TenantId
            $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.CertificateThumbprint = $CertificateThumbprint
            $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AccessTokens = $AccessTokens
            $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Identity = $Identity
            $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Endpoints = $Endpoints
            $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connect()
        }
        'MicrosoftTeams'
        {
            $Script:MSCloudLoginConnectionProfile.Teams.Credentials = $Credential
            $Script:MSCloudLoginConnectionProfile.Teams.ApplicationId = $ApplicationId
            $Script:MSCloudLoginConnectionProfile.Teams.ApplicationSecret = $ApplicationSecret
            $Script:MSCloudLoginConnectionProfile.Teams.TenantId = $TenantId
            $Script:MSCloudLoginConnectionProfile.Teams.CertificateThumbprint = $CertificateThumbprint
            $Script:MSCloudLoginConnectionProfile.Teams.CertificatePath = $CertificatePath
            $Script:MSCloudLoginConnectionProfile.Teams.CertificatePassword = $CertificatePassword
            $Script:MSCloudLoginConnectionProfile.Teams.AccessTokens = $AccessTokens
            $Script:MSCloudLoginConnectionProfile.Teams.Identity = $Identity
            $Script:MSCloudLoginConnectionProfile.Teams.Endpoints = $Endpoints
            $Script:MSCloudLoginConnectionProfile.Teams.Connect()
        }
        'PnP'
        {
            $Script:MSCloudLoginConnectionProfile.PnP.Credentials = $Credential
            $Script:MSCloudLoginConnectionProfile.PnP.ApplicationId = $ApplicationId
            $Script:MSCloudLoginConnectionProfile.PnP.ApplicationSecret = $ApplicationSecret
            $Script:MSCloudLoginConnectionProfile.PnP.TenantId = $TenantId
            $Script:MSCloudLoginConnectionProfile.PnP.CertificateThumbprint = $CertificateThumbprint
            $Script:MSCloudLoginConnectionProfile.PnP.CertificatePath = $CertificatePath
            $Script:MSCloudLoginConnectionProfile.PnP.AccessTokens = $AccessTokens
            $Script:MSCloudLoginConnectionProfile.PnP.Identity = $Identity
            $Script:MSCloudLoginConnectionProfile.PnP.Endpoints = $Endpoints
            $Script:MSCloudLoginConnectionProfile.PnP.CertificatePassword = $CertificatePassword

            # Mark as disconnected if we are trying to connect to a different url then we previously connected to.
            if ($Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl -ne $Url -or `
                    -not $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl -and `
                    $Url -or (-not $Url -and -not $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl))
            {
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
                    Write-Information -MessageData "Couldn't acquire PnP Context"
                }
            }

            # If the AdminUrl is empty and a URL was provided, assume that the url
            # provided is the admin center;
            if (-not $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl -and $Url)
            {
                $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl = $Url
            }
        }
        'PowerPlatforms'
        {
            $Script:MSCloudLoginConnectionProfile.PowerPlatform.Credentials = $Credential
            $Script:MSCloudLoginConnectionProfile.PowerPlatform.ApplicationId = $ApplicationId
            $Script:MSCloudLoginConnectionProfile.PowerPlatform.TenantId = $TenantId
            $Script:MSCloudLoginConnectionProfile.PowerPlatform.CertificateThumbprint = $CertificateThumbprint
            $Script:MSCloudLoginConnectionProfile.PowerPlatform.ApplicationSecret = $ApplicationSecret
            $Script:MSCloudLoginConnectionProfile.PowerPlatform.AccessTokens = $AccessTokens
            $Script:MSCloudLoginConnectionProfile.PowerPlatform.Endpoints = $Endpoints
            $Script:MSCloudLoginConnectionProfile.PowerPlatform.Connect()
        }
        'SecurityComplianceCenter'
        {
            $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Credentials = $Credential
            $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ApplicationId = $ApplicationId
            $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ApplicationSecret = $ApplicationSecret
            $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.TenantId = $TenantId
            $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.CertificateThumbprint = $CertificateThumbprint
            $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.CertificatePath = $CertificatePath
            $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.CertificatePassword = $CertificatePassword
            $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.AccessTokens = $AccessTokens
            $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.SkipModuleReload = $SkipModuleReload
            $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Endpoints = $Endpoints
            $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Connect()
        }
        'SharePointOnlineREST'
        {
            $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.Credentials = $Credential
            $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.ApplicationId = $ApplicationId
            $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.ApplicationSecret = $ApplicationSecret
            $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.TenantId = $TenantId
            $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.CertificateThumbprint = $CertificateThumbprint
            $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.AccessTokens = $AccessTokens
            $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.Identity = $Identity
            $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.Endpoints = $Endpoints

            $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.Connected = $false
            $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.ConnectionUrl = $Url
            $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.Connect()

            # If the AdminUrl is empty and a URL was provided, assume that the url
            # provided is the admin center;
            if (-not $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl -and $Url)
            {
                $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl = $Url
            }
        }
        'Tasks'
        {
            $Script:MSCloudLoginConnectionProfile.Tasks.Credentials = $Credential
            $Script:MSCloudLoginConnectionProfile.Tasks.ApplicationId = $ApplicationId
            $Script:MSCloudLoginConnectionProfile.Tasks.ApplicationSecret = $ApplicationSecret
            $Script:MSCloudLoginConnectionProfile.Tasks.TenantId = $TenantId
            $Script:MSCloudLoginConnectionProfile.Tasks.CertificateThumbprint = $CertificateThumbprint
            $Script:MSCloudLoginConnectionProfile.Tasks.CertificatePath = $CertificatePath
            $Script:MSCloudLoginConnectionProfile.Tasks.CertificatePassword = $CertificatePassword
            $Script:MSCloudLoginConnectionProfile.Tasks.AccessTokens = $AccessTokens
            $Script:MSCloudLoginConnectionProfile.Tasks.Endpoints = $Endpoints
            $Script:MSCloudLoginConnectionProfile.Tasks.Connect()
        }
    }
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
        [ValidateSet('AdminAPI', 'Azure', 'AzureDevOPS', 'ExchangeOnline', 'Fabric', `
                'SecurityComplianceCenter', 'PnP', 'PowerPlatforms', `
                'MicrosoftTeams', 'MicrosoftGraph', 'SharePointOnlineREST', 'Tasks', 'DefenderForEndpoint')]
        [System.String]
        $Workload
    )

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
        [ValidateSet('AdminAPI', 'Azure', 'AzureDevOPS', 'ExchangeOnline', 'Fabric', `
                'SecurityComplianceCenter', 'PnP', 'PowerPlatforms', `
                'MicrosoftTeams', 'MicrosoftGraph', 'SharePointOnlineREST', 'Tasks', 'DefenderForEndpoint')]
        [System.String[]]
        $Workload
    )

    $fullReset = $false
    if ($Workload.Count -eq 0)
    {
        $Workload = $Script:MSCloudLoginConnectionProfile.PSObject.Properties.Name | Where-Object { $_ -notin @('CreatedTime', 'OrganizationName') }
        $fullReset = $true
    }

    $source = 'Reset-MSCloudLoginConnectionProfileContext'
    Add-MSCloudLoginAssistantEvent -Message 'Resetting connection profile' -Source $source
    foreach ($workloadToReset in $Workload)
    {
        $disconnectExists = $null -ne ($Script:MSCloudLoginConnectionProfile.$workloadToReset | Get-Member -Name 'Disconnect' -MemberType Method)
        if ($disconnectExists)
        {
            $Script:MSCloudLoginConnectionProfile.$workloadToReset.Disconnect()
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
            if ([System.Diagnostics.EventLog]::Exists($logName) -eq $false)
            {
                # Create event log
                $null = New-EventLog -LogName $logName -Source $Source
            }
            else
            {
                [System.Diagnostics.EventLog]::CreateEventSource($Source, $logName)
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
            Write-EventLog -LogName $logName -Source $Source `
                -EventId $EventID -Message $outputMessage -EntryType $EntryType -ErrorAction Stop
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
    param (
        [Parameter()]
        [System.Collections.Hashtable]
        $CurrentParamSet
    )

    $currentParameters = $currentParamSet

    if ($null -ne $currentParameters['Credential'].UserName)
    {
        $currentParameters.Add('UserName', $currentParameters['Credential'].UserName)
    }
    $currentParameters.Remove('Credential') | Out-Null
    $currentParameters.Remove('SkipModuleReload') | Out-Null
    $currentParameters.Remove('CmdletsToLoad') | Out-Null
    $currentParameters.Remove('UseModernAuth') | Out-Null
    $currentParameters.Remove('ProfileName') | Out-Null
    $currentParameters.Remove('Verbose') | Out-Null

    $globalParameters = @{}

    $workloadProfile = $Script:MSCloudLoginConnectionProfile

    if ($null -eq $workloadProfile)
    {
        # No Workload profile yet, so we need to connect
        # This should not happen, but just in case
        # We are not able to detect a change, so we return $false
        return $false
    }
    else
    {
        $workload = $currentParameters['Workload']

        if ($Workload -eq 'MicrosoftTeams')
        {
            $workloadInternalName = 'Teams'
        }
        elseif ($Workload -eq 'PowerPlatforms')
        {
            $workloadInternalName = 'PowerPlatform'
        }
        else
        {
            $workloadInternalName = $workload
        }
        $workloadProfile = $Script:MSCloudLoginConnectionProfile.$workloadInternalName
    }

    # Clean the global Params
    if (-not [System.String]::IsNullOrEmpty($workloadProfile.TenantId))
    {
        $globalParameters.Add('TenantId', $workloadProfile.TenantId)
    }
    if (-not [System.String]::IsNullOrEmpty($workloadProfile.Credentials.UserName))
    {
        $globalParameters.Add('UserName', $workloadProfile.Credentials.UserName)

        # If the tenant id is part of the username, we need to remove it from the global parameters
        if ($workloadInternalName -eq 'MicrosoftGraph' `
                -and $globalParameters.ContainsKey('TenantId') `
                -and $globalParameters.TenantId -eq $workloadProfile.Credentials.UserName.Split('@')[1])
        {
            $globalParameters.Remove('TenantId') | Out-Null
        }
    }
    if ($workloadInternalName -eq 'PNP' -and $currentParameters.ContainsKey('Url') -and `
        -not [System.String]::IsNullOrEmpty($currentParameters.Url))
    {
        $globalParameters.Add('Url', $workloadProfile.ConnectionUrl)
    }
    if ($null -ne $workloadProfile.ExchangeOnlineCmdlets)
    {
        $globalParameters.Add('ExchangeOnlineCmdlets', $ExchangeOnlineCmdlets)
    }

    # This is the global graph application id. If it is something different, it means that we should compare the parameters
    if (-not [System.String]::IsNullOrEmpty($workloadProfile.ApplicationId) `
            -and -not($workloadInternalName -eq 'MicrosoftGraph' -and $workloadProfile.ApplicationId -eq '14d82eec-204b-4c2f-b7e8-296a70dab67e'))
    {
        $globalParameters.Add('ApplicationId', $workloadProfile.ApplicationId)
    }

    if (-not [System.String]::IsNullOrEmpty($workloadProfile.ApplicationSecret))
    {
        $globalParameters.Add('ApplicationSecret', $workloadProfile.ApplicationSecret)
    }
    if (-not [System.String]::IsNullOrEmpty($workloadProfile.CertificateThumbprint))
    {
        $globalParameters.Add('CertificateThumbprint', $workloadProfile.CertificateThumbprint)
    }
    if (-not [System.String]::IsNullOrEmpty($workloadProfile.CertificatePassword))
    {
        $globalParameters.Add('CertificatePassword', $workloadProfile.CertificatePassword)
    }
    if (-not [System.String]::IsNullOrEmpty($workloadProfile.CertificatePath))
    {
        $globalParameters.Add('CertificatePath', $workloadProfile.CertificatePath)
    }
    if ($workloadProfile.Identity)
    {
        $globalParameters.Add('Identity', $workloadProfile.Identity)
    }
    if ($workloadProfile.AccessTokens)
    {
        $globalParameters.Add('AccessTokens', $workloadProfile.AccessTokens)
    }

    # Clean the current parameters

    # Remove the workload, as we don't need to compare that
    $currentParameters.Remove('Workload') | Out-Null

    if ([System.String]::IsNullOrEmpty($currentParameters.ApplicationId))
    {
        $currentParameters.Remove('ApplicationId') | Out-Null
    }
    if ([System.String]::IsNullOrEmpty($currentParameters.TenantId))
    {
        $currentParameters.Remove('TenantId') | Out-Null
    }
    if ([System.String]::IsNullOrEmpty($currentParameters.ApplicationSecret))
    {
        $currentParameters.Remove('ApplicationSecret') | Out-Null
    }
    if ([System.String]::IsNullOrEmpty($currentParameters.CertificateThumbprint))
    {
        $currentParameters.Remove('CertificateThumbprint') | Out-Null
    }
    if ([System.String]::IsNullOrEmpty($currentParameters.CertificatePassword))
    {
        $currentParameters.Remove('CertificatePassword') | Out-Null
    }
    if ([System.String]::IsNullOrEmpty($currentParameters.CertificatePath))
    {
        $currentParameters.Remove('CertificatePath') | Out-Null
    }
    if ($currentParameters.ContainsKey('Identity') -and -not ($currentParameters.Identity))
    {
        $currentParameters.Remove('Identity') | Out-Null
    }

    if ($null -ne $globalParameters)
    {
        $diffKeys = Compare-Object -ReferenceObject @($currentParameters.Keys) -DifferenceObject @($globalParameters.Keys) -PassThru
        $compareValues = @($currentParameters.Values) | Where-Object { $_ -ne $null }
        $diffValues = Compare-Object -ReferenceObject $compareValues -DifferenceObject @($globalParameters.Values) -PassThru
    }

    if ($null -eq $diffKeys -and $null -eq $diffValues)
    {
        # no differences were found
        return $false
    }

    # We found differences, so we need to connect
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

    try
    {
        $result = Invoke-MgGraphRequest -Uri /v1.0/sites/root -ErrorAction SilentlyContinue
        $weburl = $result.webUrl
        if (-not $weburl)
        {
            Connect-M365Tenant -Workload 'MicrosoftGraph' -Credential $Credential
            $weburl = (Invoke-MgGraphRequest -Uri /v1.0/sites/root).webUrl
        }
    }
    catch
    {
        Connect-M365Tenant -Workload 'MicrosoftGraph' -Credential $Credential
        try
        {
            $weburl = (Invoke-MgGraphRequest -Uri /v1.0/sites/root).webUrl
        }
        catch
        {
            if (Assert-IsNonInteractiveShell -eq $false)
            {
                # Only run interactive command when Exporting
                Add-MSCloudLoginAssistantEvent -Message 'Requesting access to read information about the domain' -Source $source
                Connect-MgGraph -Scopes Sites.Read.All -ErrorAction 'Stop'
                $weburl = (Invoke-MgGraphRequest -Uri /v1.0/sites/root).webUrl
            }
            else
            {
                if ($_.Exception.Message -eq 'Insufficient privileges to complete the operation.')
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

function Get-AzureADDLL
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param(
    )
    [array]$AzureADModules = Get-Module -ListAvailable | Where-Object { $_.name -eq 'AzureADPreview' }

    if ($AzureADModules.count -eq 0)
    {
        Throw "Can't find Azure AD DLL. Install the module manually 'Install-Module AzureADPreview'"
    }
    else
    {
        $AzureDLL = Join-Path (($AzureADModules | Sort-Object version -Descending | Select-Object -First 1).Path | Split-Path) Microsoft.IdentityModel.Clients.ActiveDirectory.dll
        return $AzureDLL
    }

}

function Get-TenantLoginEndPoint
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param(
        [Parameter(Mandatory = $True)]
        [System.String]
        $TenantName,
        [Parameter(Mandatory = $false)]
        [System.String]
        [ValidateSet('MicrosoftOnline', 'EvoSTS')]
        $LoginSource = 'EvoSTS'
    )

    $TenantInfo = @{ }
    if ($LoginSource -eq 'EvoSTS')
    {
        $webrequest = Invoke-WebRequest -Uri https://login.windows.net/$($TenantName)/.well-known/openid-configuration -UseBasicParsing
    }
    else
    {
        $webrequest = Invoke-WebRequest -Uri https://login.microsoftonline.com/$($TenantName)/.well-known/openid-configuration -UseBasicParsing
    }
    if ($webrequest.StatusCode -eq 200)
    {
        $TenantInfo = $webrequest.Content | ConvertFrom-Json
    }
    return $TenantInfo
}

function New-ADALServiceInfo
{
    [CmdletBinding()]
    [OutputType([System.Collections.HashTable])]
    param(
        [Parameter(Mandatory = $True)]
        [System.String]
        $TenantName,

        [Parameter(Mandatory = $True)]
        [System.String]
        $UserPrincipalName,

        [Parameter(Mandatory = $false)]
        [System.String]
        [ValidateSet('MicrosoftOnline', 'EvoSTS')]
        $LoginSource = 'EvoSTS'
    )

    $source = 'New-ADALServiceInfo'
    $AzureADDLL = Get-AzureADDLL
    if ([string]::IsNullOrEmpty($AzureADDLL))
    {
        Throw "Can't find Azure AD DLL"
        Exit
    }
    else
    {
        Add-MSCloudLoginAssistantEvent -Message "AzureADDLL: $AzureADDLL" -Source $source
        $tMod = [System.Reflection.Assembly]::LoadFrom($AzureADDLL)
    }

    $TenantInfo = Get-TenantLoginEndPoint -TenantName $TenantName
    if ([string]::IsNullOrEmpty($TenantInfo))
    {
        Throw "Can't find Tenant Login Endpoint"
        Exit
    }
    else
    {
        [string] $authority = $TenantInfo.authorization_endpoint
    }
    $PromptBehavior = [Microsoft.IdentityModel.Clients.ActiveDirectory.PromptBehavior]::Auto
    $Service = @{ }
    $Service['authContext'] = [Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext]::new($authority, $false)
    $Service['platformParam'] = New-Object 'Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters' -ArgumentList $PromptBehavior
    $Service['userId'] = New-Object 'Microsoft.IdentityModel.Clients.ActiveDirectory.UserIdentifier' -ArgumentList $UserPrincipalName, 'OptionalDisplayableId'

    Add-MSCloudLoginAssistantEvent -Message "Current Assembly for AD AuthenticationContext: $([Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext].Assembly | Out-String)" -Source $source

    return $Service
}

function Get-AuthHeader
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $True)]
        [System.String]
        $UserPrincipalName,
        [Parameter(Mandatory = $True)]
        [Alias('RessourceURI')] # For backward compat with anything using the misspelled parameter
        $ResourceURI,
        [Parameter(Mandatory = $True)]
        $clientId,
        [Parameter(Mandatory = $True)]
        [System.String]
        $RedirectURI
    )

    if ($null -eq $Script:ADALServicePoint)
    {
        $TenantName = $UserPrincipalName.split('@')[1]
        $Script:ADALServicePoint = New-ADALServiceInfo -TenantName $TenantName -UserPrincipalName $UserPrincipalName
    }

    try
    {
        Write-Debug 'Looking for a refresh token'
        $authResult = $Script:ADALServicePoint.authContext.AcquireTokenSilentAsync($ResourceURI, $clientId)
        if ($null -eq $authResult.result)
        {
            $RedirectURI = [System.Uri]::new($RedirectURI)
            $authResult = $Script:ADALServicePoint.authContext.AcquireTokenAsync($ResourceURI, $clientId, $RedirectURI, $Script:ADALServicePoint.platformParam, $Script:ADALServicePoint.userId, '', '')
        }
        $AuthHeader = $authResult.result.CreateAuthorizationHeader()
    }
    catch
    {
        Throw "Can't create Authorization header: $_"
    }
    Return $AuthHeader
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
        $Certificate = Get-Item "Cert:\CurrentUser\My\$($CertificateThumbprint)" -ErrorAction SilentlyContinue
        if ($null -eq $Certificate)
        {
            Add-MSCloudLoginAssistantEvent 'Certificate not found in CurrentUser\My' -Source $source
            $Certificate = Get-ChildItem "Cert:\LocalMachine\My\$($CertificateThumbprint)" -ErrorAction SilentlyContinue

            if ($null -eq $Certificate)
            {
                throw 'Certificate not found in LocalMachine\My'
            }
        }
        # Create base64 hash of certificate
        $CertificateBase64Hash = [System.Convert]::ToBase64String($Certificate.GetCertHash())

        # Create JWT timestamp for expiration
        $StartDate = (Get-Date '1970-01-01T00:00:00Z' ).ToUniversalTime()
        $JWTExpirationTimeSpan = (New-TimeSpan -Start $StartDate -End (Get-Date).ToUniversalTime().AddMinutes(2)).TotalSeconds
        $JWTExpiration = [math]::Round($JWTExpirationTimeSpan, 0)

        # Create JWT validity start timestamp
        $NotBeforeExpirationTimeSpan = (New-TimeSpan -Start $StartDate -End ((Get-Date).ToUniversalTime())).TotalSeconds
        $NotBefore = [math]::Round($NotBeforeExpirationTimeSpan, 0)

        # Create JWT header
        $JWTHeader = @{
            alg = 'RS256'
            typ = 'JWT'
            # Use the CertificateBase64Hash and replace/strip to match web encoding of base64
            x5t = $CertificateBase64Hash -replace '\+', '-' -replace '/', '_' -replace '='
        }

        # Create JWT payload
        $JWTPayLoad = @{
            # What endpoint is allowed to use this JWT
            aud = "$($AzureADAuthorizationEndpointUri)/$($TenantId)/oauth2/token"

            # Expiration timestamp
            exp = $JWTExpiration

            # Issuer = your application
            iss = $ApplicationId

            # JWT ID: random guid
            jti = [guid]::NewGuid()

            # Not to be used before
            nbf = $NotBefore

            # JWT Subject
            sub = $ApplicationId
        }

        # Convert header and payload to base64
        $JWTHeaderToByte = [System.Text.Encoding]::UTF8.GetBytes(($JWTHeader | ConvertTo-Json))
        $EncodedHeader = [System.Convert]::ToBase64String($JWTHeaderToByte)

        $JWTPayLoadToByte = [System.Text.Encoding]::UTF8.GetBytes(($JWTPayload | ConvertTo-Json))
        $EncodedPayload = [System.Convert]::ToBase64String($JWTPayLoadToByte)

        # Join header and Payload with "." to create a valid (unsigned) JWT
        $JWT = $EncodedHeader + '.' + $EncodedPayload

        # Get the private key object of your certificate
        $PrivateKey = ([System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Certificate))

        # Define RSA signature and hashing algorithm
        $RSAPadding = [Security.Cryptography.RSASignaturePadding]::Pkcs1
        $HashAlgorithm = [Security.Cryptography.HashAlgorithmName]::SHA256

        # Create a signature of the JWT
        $Signature = [Convert]::ToBase64String(
            $PrivateKey.SignData([System.Text.Encoding]::UTF8.GetBytes($JWT), $HashAlgorithm, $RSAPadding)
        ) -replace '\+', '-' -replace '/', '_' -replace '='

        # Join the signature to the JWT with "."
        $JWT = $JWT + '.' + $Signature

        # Create a hash with body parameters
        $Body = @{
            client_id             = $appId
            client_assertion      = $JWT
            client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
            scope                 = $ConnectionUri
            grant_type            = 'client_credentials'
        }

        $Url = "$($AzureADAuthorizationEndpointUri)/$($TenantId)/oauth2/v2.0/token"

        # Use the self-generated JWT as Authorization
        $Header = @{
            Authorization = "Bearer $JWT"
        }

        # Splat the parameters for Invoke-Restmethod for cleaner code
        $PostSplat = @{
            ContentType = 'application/x-www-form-urlencoded'
            Method      = 'POST'
            Body        = $Body
            Uri         = $Url
            Headers     = $Header
        }

        $Request = Invoke-RestMethod @PostSplat
        return $Request.access_token
    }
    catch
    {
        Add-MSCloudLoginAssistantEvent -Message $_ -Source $source -EntryType Error
        throw $_
    }
}

function Get-PowerPlatformTokenInfo
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.String]
        $Audience,

        [Parameter()]
        [System.Management.Automation.PSCredential]
        $Credentials
    )

    $jobName = 'AcquireTokenAsync' + (New-Guid).ToString()
    Start-Job -Name $jobName -ScriptBlock {
        param(
            [Parameter(Mandatory = $true)]
            [System.Management.Automation.PSCredential]
            $O365Credentials,

            [Parameter(Mandatory = $true)]
            [System.String]
            $Audience
        )

        try
        {
            Import-Module -Name 'Microsoft.PowerApps.Administration.PowerShell' -Force
            $authContext = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext('https://login.windows.net/common')
            $credential = [Microsoft.IdentityModel.Clients.ActiveDirectory.UserCredential]::new($O365Credentials.Username, $O365Credentials.Password)
            $authResult = $authContext.AcquireToken($Audience, '1950a258-227b-4e31-a9cf-717495945fc2', $credential)

            $JwtToken = $authResult.IdToken
            $tokenSplit = $JwtToken.Split('.')
            $claimsSegment = $tokenSplit[1].Replace(' ', '+')

            $mod = $claimsSegment.Length % 4
            if ($mod -gt 0)
            {
                $paddingCount = 4 - $mod
                for ($i = 0; $i -lt $paddingCount; $i++)
                {
                    $claimsSegment += '='
                }
            }
            $decodedClaimsSegment = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($claimsSegment))
            $claims = ConvertFrom-Json $decodedClaimsSegment
        }
        catch
        {
            $_ | Out-File "$env:temp\MSCloudLoginAssistant_Error.txt"
        }
        return @{
            JwtToken     = $JwtToken
            Claims       = $claims
            RefreshToken = $authResult.RefreshToken
            AccessToken  = $authResult.AccessToken
            ExpiresOn    = $authResult.ExpiresOn
        }
    } -ArgumentList @($Credentials, $Audience) | Out-Null

    $job = Get-Job | Where-Object -FilterScript { $_.Name -eq $jobName }
    do
    {
        Start-Sleep -Seconds 1
    } while ($job.JobStateInfo.State -ne 'Completed')
    $TokenInfo = Receive-Job -Name $jobName
    return $TokenInfo
}

function Test-MSCloudLoginCommand
{
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param(
        [Parameter(Mandatory = $true)]
        [System.String]
        $Command
    )

    try
    {
        $testResult = Invoke-Command $Command
        return $true
    }
    catch
    {
        return $false
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

    try
    {
        if ($null -ne $Credentials)
        {
            $tenantName = $Credentials.UserName.Split('@')[1]
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
        $response = Invoke-WebRequest -Uri "https://login.microsoftonline.com/$tenantName/v2.0/.well-known/openid-configuration" -Method Get -UseBasicParsing

        $content = $response.Content
        $result = ConvertFrom-Json $content
        return $result
    }
    catch
    {
        throw $_
    }
}

function Get-MSCloudLoginTenantDomain
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
        [switch]
        $Identity,

        [Parameter()]
        [System.String[]]
        $AccessTokens
    )

    if (-not [string]::IsNullOrEmpty($ApplicationId))
    {
        Connect-M365Tenant -Workload MicrosoftGraph `
            -ApplicationId $ApplicationId `
            -TenantId $TenantId `
            -CertificateThumbprint $CertificateThumbprint
    }
    elseif ($Identity.IsPresent)
    {
        Connect-M365Tenant -Workload MicrosoftGraph `
            -Identity `
            -TenantId $TenantId
    }
    elseif ($null -ne $AccessTokens)
    {
        Connect-M365Tenant -Workload MicrosoftGraph `
            -AccessTokens $AccessTokens
    }

    try
    {
        $domain = Get-MgDomain | Where-Object { $_.IsInitial -eq $True }
    }
    catch
    {
        $domain = Get-MgBetaDomain | Where-Object { $_.IsInitial -eq $True }
    }

    if ($null -ne $domain)
    {
        return $domain.Id.split('.')[0]
    }
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
        $domain = Get-MgDomain -ErrorAction Stop | Where-Object { $_.IsInitial -eq $True }

        if ($null -ne $domain)
        {
            return $domain.Id
        }
    }
    catch
    {
        Add-MSCloudLoginAssistantEvent -Message "Couldn't get domain. Using TenantId instead" -Source $source
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
