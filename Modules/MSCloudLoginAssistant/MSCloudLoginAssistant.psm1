$Script:WriteToEventLog = ([Environment]::GetEnvironmentVariable('MSCLOUDLOGINASSISTANT_WRITETOEVENTLOG', 'Machine') -eq 'true') -or `
                          ($env:MSCLOUDLOGINASSISTANT_WRITETOEVENTLOG -eq 'true')

. "$PSScriptRoot\ConnectionProfile.ps1"
. "$PSScriptRoot\CustomEnvironment.ps1"
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
        Add-MSCloudLoginAssistantEvent -Message "Resetting connection for workload $workloadInternalName" -Source $source
        $Script:MSCloudLoginConnectionProfile.$workloadInternalName.Connected = $false
    }

    Add-MSCloudLoginAssistantEvent -Message "Checking connection to platform {$Workload}" -Source $source
    foreach ($parameter in $PSBoundParameters.GetEnumerator())
    {
        if ($parameter.Key -eq 'Credential')
        {
            $Script:MSCloudLoginConnectionProfile.$workloadInternalName.Credentials = $parameter.Value
        }
        else
        {
            if ($parameter.Key -in @('AccessTokens', 'ApplicationId', 'ApplicationSecret', 'CertificateThumbprint', 'CertificatePath', 'CertificatePassword', 'Identity', 'Endpoints', 'TenantId', 'TenantGUID'))
            {
                $Script:MSCloudLoginConnectionProfile.$workloadInternalName.($parameter.Key) = $parameter.Value
            }
        }
    }

    switch ($Workload)
    {
        'AdminAPI'
        {
            $Script:MSCloudLoginConnectionProfile.AdminAPI.Connect()
        }
        'Azure'
        {
            $Script:MSCloudLoginConnectionProfile.Azure.Connect()
        }
        'AzureDevOPS'
        {
            $Script:MSCloudLoginConnectionProfile.AzureDevOPS.Connect()
        }
        'DefenderForEndpoint'
        {
            $Script:MSCloudLoginConnectionProfile.DefenderForEndpoint.Connect()
        }
        'EngageHub'
        {
            $Script:MSCloudLoginConnectionProfile.EngageHub.Connect()
        }
        'ExchangeOnline'
        {
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CmdletsToLoad = $ExchangeOnlineCmdlets
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connect()
        }
        'Fabric'
        {
            $Script:MSCloudLoginConnectionProfile.Fabric.Connect()
        }
        'Licensing'
        {
            $Script:MSCloudLoginConnectionProfile.Licensing.Connect()
        }
        'O365Portal'
        {
            $Script:MSCloudLoginConnectionProfile.O365Portal.Connect()
        }
        'MicrosoftGraph'
        {
            $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connect()
        }
        'MicrosoftTeams'
        {
            $Script:MSCloudLoginConnectionProfile.Teams.Connect()
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
            $Script:MSCloudLoginConnectionProfile.PowerPlatform.Connect()
        }
        'PowerPlatformREST'
        {
            $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.Connect()
        }
        'SecurityComplianceCenter'
        {
            $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Connect()
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
        }
        'Tasks'
        {
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
        [ValidateSet('AdminAPI', 'Azure', 'AzureDevOPS', 'EngageHub', 'ExchangeOnline', 'Fabric', 'Licensing', `
                'O365Portal', 'SecurityComplianceCenter', 'PnP', 'PowerPlatforms', 'PowerPlatformREST', `
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
    $source = 'Compare-InputParametersForChange'

    if ($null -ne $currentParameters['Credential'].UserName)
    {
        $currentParameters.Add('UserName', $currentParameters['Credential'].UserName)
    }
    $currentParameters.Remove('Credential') | Out-Null
    $currentParameters.Remove('CmdletsToLoad') | Out-Null
    $currentParameters.Remove('UseModernAuth') | Out-Null
    $currentParameters.Remove('ProfileName') | Out-Null
    $currentParameters.Remove('Verbose') | Out-Null
    $currentParameters.Remove('ErrorAction') | Out-Null

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
        # and the current parameters, otherwise it would report as a drift
        if ($workloadInternalName -eq 'MicrosoftGraph' `
                -and $globalParameters.ContainsKey('TenantId') `
                -and $globalParameters.TenantId -eq $workloadProfile.Credentials.UserName.Split('@')[1])
        {
            $currentParameters.Remove('TenantId') | Out-Null
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
    if ([System.String]::IsNullOrEmpty($currentParameters.Url))
    {
        $currentParameters.Remove('Url') | Out-Null
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
    Add-MSCloudLoginAssistantEvent -Message "Found differences in parameters: $diffKeys, with values: $($diffValues | ConvertTo-Json)" -Source $source
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
        $result = Invoke-MgGraphRequest -Uri '/v1.0/sites/root' -ErrorAction SilentlyContinue
        $weburl = $result.webUrl
        if (-not $weburl)
        {
            Connect-M365Tenant -Workload 'MicrosoftGraph' -Credential $Credential
            $weburl = (Invoke-MgGraphRequest -Uri '/v1.0/sites/root').webUrl
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
        $certificate = Get-Item "Cert:\CurrentUser\My\$($CertificateThumbprint)" -ErrorAction SilentlyContinue
        if ($null -eq $certificate)
        {
            Add-MSCloudLoginAssistantEvent 'Certificate not found in CurrentUser\My' -Source $source
            $certificate = Get-ChildItem "Cert:\LocalMachine\My\$($CertificateThumbprint)" -ErrorAction SilentlyContinue
            if ($null -eq $certificate)
            {
                throw 'Certificate not found in LocalMachine\My'
            }
        }
        # Create base64 hash of certificate
        $CertificateBase64Hash = [System.Convert]::ToBase64String($certificate.GetCertHash())

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
            aud = $AzureADAuthorizationEndpointUri

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
            client_id             = $ApplicationId
            client_assertion      = $JWT
            client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
            scope                 = $ConnectionUri
            grant_type            = 'client_credentials'
        }

        $Url = $AzureADAuthorizationEndpointUri

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
        $Scope
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
                Invoke-WebRequest -Method GET -Uri $endpoint -Headers @{
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
    if ($useResource) {
        $tokenEndpoint = "$AuthorizationUrl/$TenantId/oauth2/token"
    } else {
        $tokenEndpoint = "$AuthorizationUrl/$TenantId/oauth2/v2.0/token"
    }

    if ($ClientSecret -or $CertificatePath -or $CertificateThumbprint) {
        if ($CertificatePath) {
            if (Test-Path $CertificatePath) {
                if ($CertificatePassword) {
                    $certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new((Resolve-Path $CertificatePath), $CertificatePassword, [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::UserKeySet)
                } else {
                    $certificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new((Resolve-Path $CertificatePath))
                }
            } else {
                throw "Certificate path '$CertificatePath' not found"
            }
        }

        if ($CertificateThumbprint) {
            $certificate = Get-Item "Cert:\CurrentUser\My\$CertificateThumbprint" -ErrorAction SilentlyContinue
            if ($null -eq $certificate) {
                $certificate = Get-Item "Cert:\LocalMachine\My\$CertificateThumbprint" -ErrorAction SilentlyContinue
                if ($null -eq $certificate) {
                    throw "Certificate with thumbprint '$CertificateThumbprint' not found in LocalMachine\My nor CurrentUser\My"
                }
            }
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
                $base64Hash = [System.Convert]::ToBase64String($certificate.GetCertHash())
                $header.Add('x5t', $base64Hash)
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
            scope = $Scope
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
    $authorizeUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/authorize?client_id=$ClientId&response_type=code&redirect_uri=$([System.Uri]::EscapeDataString($redirectUri))&response_mode=query&scope=$([System.Uri]::EscapeDataString($Scope))&code_challenge=$codeChallenge&code_challenge_method=S256"

    $listener = [System.Net.HttpListener]::new()
    $listener.Prefixes.Add($redirectUri)
    $listener.Start()
    try {
        if ($IsWindows) {
            Start-Process $authorizeUrl
        } else {
            return
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
    $listener.Stop()

    $body = @{
        client_id     = $ClientId
        scope         = $Scope
        grant_type    = 'authorization_code'
        code          = $code
        redirect_uri  = $redirectUri
        code_verifier = $codeVerifier
    }
    Invoke-RestMethod -Method Post -Uri $tokenEndpoint -Body $body -ContentType 'application/x-www-form-urlencoded'

    return $response.access_token
}
