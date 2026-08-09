<#
.SYNOPSIS
    Shared internal helper functions for the MSCloudLoginAssistant module.

.DESCRIPTION
    This file contains helper functions that are shared between the main module
    and the individual workload connection scripts. None of these functions are
    exported from the module.
#>

<#
.SYNOPSIS
    Converts a SecureString to its plain text representation.

.DESCRIPTION
    Converts a SecureString to its plain text representation using BSTR marshalling,
    which works on both Windows PowerShell 5.1 and PowerShell 7+. The unmanaged
    memory is zeroed and freed after the conversion.

.PARAMETER SecureString
    The SecureString to convert.

.OUTPUTS
    System.String. The plain text value.
#>
function ConvertFrom-SecureStringToPlainText
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.Security.SecureString]
        $SecureString
    )

    $bstr = [System.IntPtr]::Zero
    try
    {
        $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
        return [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    }
    finally
    {
        if ($bstr -ne [System.IntPtr]::Zero)
        {
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        }
    }
}

<#
.SYNOPSIS
    Extracts the plain text value from an access token in any of its supported representations.

.DESCRIPTION
    Access tokens can reach the module as plain strings, SecureStrings or PSCredentials.
    This function normalizes all three representations to the plain text token value.

.PARAMETER Token
    The token as a string, SecureString or PSCredential.

.OUTPUTS
    System.String. The plain text token value.
#>
function Get-MSCloudLoginAccessTokenValue
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.Object]
        $Token
    )

    if ($Token -is [System.Security.SecureString])
    {
        return (ConvertFrom-SecureStringToPlainText -SecureString $Token)
    }

    if ($Token -is [System.Management.Automation.PSCredential])
    {
        return (ConvertFrom-SecureStringToPlainText -SecureString $Token.Password)
    }

    return [System.String]$Token
}

<#
.SYNOPSIS
    Extracts the tenant domain from the UserName of a credential.

.DESCRIPTION
    Extracts the tenant domain (the part after the '@') from the UserName of the
    provided credential and throws a clear error if the UserName is not a UPN.

.PARAMETER Credentials
    The credential whose UserName is evaluated.

.OUTPUTS
    System.String. The tenant domain.
#>
function Get-MSCloudLoginTenantDomainFromCredentials
{
    [CmdletBinding()]
    [OutputType([System.String])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCredential]
        $Credentials
    )

    if ($Credentials.UserName -notmatch '@')
    {
        throw "Unable to determine the tenant domain: the credential UserName '$($Credentials.UserName)' is not a user principal name (user@domain)."
    }

    return $Credentials.UserName.Split('@')[1]
}

<#
.SYNOPSIS
    Removes a loaded implicit remoting proxy module that exports the specified command.

.DESCRIPTION
    Exchange Online and Security & Compliance connections generate temporary proxy
    modules. Before establishing a new connection, an existing proxy module that
    exports the given probe command needs to be removed so that a new session
    can be created.

.PARAMETER ProbeCommand
    A command name that identifies the proxy module (e.g. 'Get-AcceptedDomain').

.PARAMETER Source
    The event source to use for logging.
#>
function Remove-MSCloudLoginProxyModule
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $ProbeCommand,

        [Parameter(Mandatory = $true)]
        [System.String]
        $Source
    )

    $loadedModules = Get-Module
    $modulesToRemove = $loadedModules | Where-Object -FilterScript {
        $_.ExportedCommands.Keys.Contains($ProbeCommand)
    }

    foreach ($moduleToRemove in $modulesToRemove)
    {
        Add-MSCloudLoginAssistantEvent -Message "Removing proxy module {$($moduleToRemove.Name)}" -Source $Source
        # Temporarily set ErrorActionPreference because a failure to remove the in-use
        # proxy module must not abort the connection attempt.
        $currentErrorActionPreference = $ErrorActionPreference
        $ErrorActionPreference = 'SilentlyContinue'
        Remove-Module -Name $moduleToRemove -Force -Verbose:$false | Out-Null
        $ErrorActionPreference = $currentErrorActionPreference
    }
}

<#
.SYNOPSIS
    Resolves a certificate either by thumbprint from the certificate stores or from a PFX file.

.DESCRIPTION
    When a thumbprint is provided, the CurrentUser\My store is searched first, then
    LocalMachine\My. When a path is provided, the certificate is loaded from the PFX
    file with the optional password. A clear error is thrown when the certificate
    cannot be found.

.PARAMETER CertificateThumbprint
    The thumbprint of the certificate to look up in the certificate stores.

.PARAMETER CertificatePath
    The path to a PFX file to load the certificate from.

.PARAMETER CertificatePassword
    The password of the PFX file.

.OUTPUTS
    System.Security.Cryptography.X509Certificates.X509Certificate2. The resolved certificate.
#>
function Get-MSCloudLoginCertificate
{
    [CmdletBinding(DefaultParameterSetName = 'Thumbprint')]
    [OutputType([System.Security.Cryptography.X509Certificates.X509Certificate2])]
    param
    (
        [Parameter(Mandatory = $true, ParameterSetName = 'Thumbprint')]
        [System.String]
        $CertificateThumbprint,

        [Parameter(Mandatory = $true, ParameterSetName = 'Path')]
        [System.String]
        $CertificatePath,

        [Parameter(ParameterSetName = 'Path')]
        [System.Security.SecureString]
        $CertificatePassword
    )

    if ($PSCmdlet.ParameterSetName -eq 'Thumbprint')
    {
        $certificate = Get-Item -Path "Cert:\CurrentUser\My\$CertificateThumbprint" -ErrorAction SilentlyContinue
        if ($null -eq $certificate)
        {
            $certificate = Get-Item -Path "Cert:\LocalMachine\My\$CertificateThumbprint" -ErrorAction SilentlyContinue
        }
        if ($null -eq $certificate)
        {
            throw "Certificate with thumbprint '$CertificateThumbprint' was not found in the CurrentUser\My nor the LocalMachine\My certificate store."
        }
        return $certificate
    }

    if (-not (Test-Path -Path $CertificatePath))
    {
        throw "Certificate path '$CertificatePath' was not found."
    }

    $resolvedPath = (Resolve-Path -Path $CertificatePath).ProviderPath
    if ($null -ne $CertificatePassword)
    {
        return [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
            $resolvedPath,
            $CertificatePassword,
            [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::UserKeySet)
    }

    return [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($resolvedPath)
}

<#
.SYNOPSIS
    Determines whether an error indicates that multi-factor authentication is required.

.DESCRIPTION
    Normalizes the detection of MFA-required errors across all workloads by checking
    the ErrorDetails message, the exception message and the string representation of
    the error record against the known MFA-related error patterns.

.PARAMETER ErrorRecord
    The error record to inspect.

.PARAMETER AdditionalPatterns
    Additional workload-specific wildcard patterns that also indicate an MFA requirement.

.OUTPUTS
    System.Boolean. $true when the error indicates that MFA is required.
#>
function Test-MSCloudLoginMFARequiredError
{
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.ErrorRecord]
        $ErrorRecord,

        [Parameter()]
        [System.String[]]
        $AdditionalPatterns = @()
    )

    $texts = @(
        $ErrorRecord.ErrorDetails.Message
        $ErrorRecord.Exception.Message
        $ErrorRecord.ToString()
    ) | Where-Object -FilterScript { -not [System.String]::IsNullOrEmpty($_) }

    # AADSTS50076: MFA required. AADSTS50079: MFA enrollment required.
    # AADSTS50158: conditional access / external security challenge.
    $patterns = @(
        '*AADSTS50076*'
        '*AADSTS50079*'
        '*AADSTS50158*'
        '*multi-factor authentication*'
    ) + $AdditionalPatterns

    foreach ($text in $texts)
    {
        foreach ($pattern in $patterns)
        {
            if ($text -like $pattern)
            {
                return $true
            }
        }
    }
    return $false
}

<#
.SYNOPSIS
    Resolves the endpoint information of a workload for a given cloud environment.

.DESCRIPTION
    Reads the per-workload, per-environment endpoint table from WorkloadEndpoints.psd1
    (cached after the first load) and returns a hashtable with the resolved endpoint
    values. Environments without an explicit entry fall back to the workload's
    'default' entry. For the 'Custom' environment, the table values are the names of
    keys in the custom environment configuration and are resolved against
    $Script:CustomEnvConfig. '{Placeholder}' tokens are replaced with the values
    provided via the Replacements parameter.

.PARAMETER Workload
    The workload name as used in the endpoint table (e.g. 'AdminAPI').

.PARAMETER EnvironmentName
    The cloud environment name (e.g. 'AzureCloud', 'AzureDOD', 'Custom').

.PARAMETER Replacements
    Optional hashtable of placeholder names to values, e.g. @{ Resource = '...' }
    replaces '{Resource}' in all endpoint values.

.OUTPUTS
    System.Collections.Hashtable. The resolved endpoint values.
#>
function Get-MSCloudLoginEndpointInfo
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $Workload,

        [Parameter(Mandatory = $true)]
        [System.String]
        $EnvironmentName,

        [Parameter()]
        [System.Collections.Hashtable]
        $Replacements = @{}
    )

    if ($null -eq $Script:WorkloadEndpointData)
    {
        $Script:WorkloadEndpointData = Import-PowerShellDataFile -Path "$PSScriptRoot\WorkloadEndpoints.psd1" -ErrorAction Stop
    }

    $workloadTable = $Script:WorkloadEndpointData[$Workload]
    if ($null -eq $workloadTable)
    {
        throw "No endpoint information is defined for workload '$Workload'."
    }

    $entry = $workloadTable[$EnvironmentName]
    if ($null -eq $entry)
    {
        $entry = $workloadTable['default']
    }
    if ($null -eq $entry)
    {
        throw "No endpoint information is defined for workload '$Workload' in environment '$EnvironmentName' and the workload has no default entry."
    }

    $result = @{}
    foreach ($property in $entry.Keys)
    {
        $value = $entry[$property]
        if ($EnvironmentName -eq 'Custom')
        {
            # Custom entries hold the key names of the custom environment configuration.
            $value = $Script:CustomEnvConfig[$value]
        }
        elseif ($value -is [System.String])
        {
            foreach ($placeholder in $Replacements.Keys)
            {
                $value = $value.Replace("{$placeholder}", [System.String]$Replacements[$placeholder])
            }
        }
        $result[$property] = $value
    }
    return $result
}

<#
.SYNOPSIS
    Derives the SharePoint Online admin and connection URLs from an onmicrosoft tenant name.

.DESCRIPTION
    Central implementation of the tenant-name-to-SharePoint-URL mapping used by the
    PnP and SharePointOnlineREST workloads. Supports commercial, GCC High, DoD,
    China and .onms. tenants.

.PARAMETER TenantId
    The tenant name, e.g. contoso.onmicrosoft.com, contoso.partner.onmschina.cn or contoso.onms.tld.

.PARAMETER EnvironmentName
    The cloud environment name (e.g. AzureCloud, AzureUSGovernment, AzureDOD).

.OUTPUTS
    System.Collections.Hashtable with the keys AdminUrl and ConnectionUrl.
#>
function Get-MSCloudLoginSPOUrlFromTenantId
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $TenantId,

        [Parameter(Mandatory = $true)]
        [System.String]
        $EnvironmentName
    )

    if ($TenantId.Contains('onmicrosoft'))
    {
        if ($EnvironmentName -eq 'AzureDOD')
        {
            $domain = $TenantId.Replace('.onmicrosoft.', '-admin.sharepoint-mil.')
        }
        else
        {
            $domain = $TenantId.Replace('.onmicrosoft.', '-admin.sharepoint.')
        }
        if ($EnvironmentName -in @('AzureUSGovernment', 'AzureDOD'))
        {
            # If the tenant id is in the format of contoso.onmicrosoft.com, replace the .com with .us for sovereign clouds
            $domain = $domain.Replace('.com', '.us')
        }
    }
    elseif ($TenantId.Contains('.onmschina.'))
    {
        $domain = $TenantId.Replace('.partner.onmschina.', '-admin.sharepoint.')
    }
    elseif ($TenantId.Contains('.onms.'))
    {
        $domain = $TenantId.Replace('.onms.', '-admin.spo.')
    }
    else
    {
        throw 'TenantId must be in format contoso.onmicrosoft.com'
    }

    return @{
        AdminUrl      = "https://$domain"
        ConnectionUrl = ("https://$domain").Replace('-admin', '')
    }
}

<#
.SYNOPSIS
    Determines whether an existing workload connection can be reused.

.DESCRIPTION
    Central connection-freshness check for all workloads. The connection is NOT
    reusable when the profile is not connected, when the connection timestamp is
    missing, when a token-based authentication type has exceeded its expiration
    window or when the optional probe script indicates that the underlying SDK
    context is gone. In all of those cases the profile is marked as disconnected
    so that a reconnect is performed.

.PARAMETER WorkloadProfile
    The workload connection profile to check.

.PARAMETER TokenExpirationMinutes
    The number of minutes after which a token-based connection is considered expired.

.PARAMETER TokenBasedAuthTypes
    The authentication types whose tokens expire and require renewal.

.PARAMETER ProbeScript
    Optional script block that returns the SDK context (e.g. { Get-MgContext }).
    A $null result marks the connection as not reusable.

.PARAMETER Source
    The event source to use for logging.

.OUTPUTS
    System.Boolean. $true when the existing connection can be reused.
#>
function Test-MSCloudLoginConnectionReusable
{
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.Object]
        $WorkloadProfile,

        [Parameter()]
        [System.Int32]
        $TokenExpirationMinutes = 50,

        [Parameter()]
        [System.String[]]
        $TokenBasedAuthTypes = @('ServicePrincipalWithSecret', 'Identity'),

        [Parameter()]
        [System.Management.Automation.ScriptBlock]
        $ProbeScript,

        [Parameter(Mandatory = $true)]
        [System.String]
        $Source
    )

    if (-not $WorkloadProfile.Connected)
    {
        return $false
    }

    if ([System.String]::IsNullOrEmpty($WorkloadProfile.ConnectedDateTime))
    {
        Add-MSCloudLoginAssistantEvent -Message 'Connection has no timestamp, reconnecting' -Source $Source
        $WorkloadProfile.Connected = $false
        return $false
    }

    if ($WorkloadProfile.AuthenticationType -in $TokenBasedAuthTypes -and `
        (Get-Date -Date $WorkloadProfile.ConnectedDateTime) -lt [System.DateTime]::Now.AddMinutes(-$TokenExpirationMinutes))
    {
        Add-MSCloudLoginAssistantEvent -Message 'Token is about to expire, renewing' -Source $Source
        $WorkloadProfile.Connected = $false
        return $false
    }

    if ($null -ne $ProbeScript)
    {
        $probeResult = $null
        try
        {
            $probeResult = & $ProbeScript
        }
        catch
        {
            # Liveness probe: a failure only means that the SDK context is gone.
            Add-MSCloudLoginAssistantEvent -Message "Connection probe failed: $($_.Exception.Message)" -Source $Source
        }
        if ($null -eq $probeResult)
        {
            $WorkloadProfile.Connected = $false
            return $false
        }
    }

    return $true
}

<#
.SYNOPSIS
    Determines whether a parameter value is considered empty.

.DESCRIPTION
    Used by the connection parameter comparison to treat absent, $null, empty string,
    unset switch, $false, empty SecureString, empty collection and empty dictionary
    values as equivalent.

.PARAMETER Value
    The value to test.

.OUTPUTS
    System.Boolean. $true when the value is considered empty.
#>
function Test-MSCloudLoginParameterValueEmpty
{
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param
    (
        [Parameter()]
        [System.Object]
        $Value
    )

    if ($null -eq $Value)
    {
        return $true
    }
    if ($Value -is [System.String])
    {
        return [System.String]::IsNullOrEmpty($Value)
    }
    if ($Value -is [System.Management.Automation.SwitchParameter])
    {
        return -not $Value.IsPresent
    }
    if ($Value -is [System.Boolean])
    {
        return -not $Value
    }
    if ($Value -is [System.Security.SecureString])
    {
        return $Value.Length -eq 0
    }
    if ($Value -is [System.Collections.IDictionary])
    {
        return $Value.Count -eq 0
    }
    if ($Value -is [System.Collections.ICollection])
    {
        return $Value.Count -eq 0
    }
    return $false
}

<#
.SYNOPSIS
    Compares two parameter values for equality with type-aware semantics.

.DESCRIPTION
    Compares two values of a connection parameter. SecureStrings are decrypted in
    memory only and compared ordinally, PSCredentials compare user name (case-insensitive)
    and password, dictionaries are compared recursively per key, collections are compared
    element-wise (order-sensitive except for CmdletsToLoad) and strings are compared
    case-insensitively for identifiers and case-sensitively for secrets.

.PARAMETER KeyName
    The canonical name of the parameter being compared. Determines the comparison semantics.

.PARAMETER Left
    The first value.

.PARAMETER Right
    The second value.

.OUTPUTS
    System.Boolean. $true when both values are considered equal.
#>
function Test-MSCloudLoginParameterValueEqual
{
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $KeyName,

        [Parameter()]
        [System.Object]
        $Left,

        [Parameter()]
        [System.Object]
        $Right
    )

    # SecureString: decrypt in memory only, compare ordinally. Values are never logged.
    if ($Left -is [System.Security.SecureString] -or $Right -is [System.Security.SecureString])
    {
        if (-not ($Left -is [System.Security.SecureString] -and $Right -is [System.Security.SecureString]))
        {
            return $false
        }
        $leftPlain = ConvertFrom-SecureStringToPlainText -SecureString $Left
        $rightPlain = ConvertFrom-SecureStringToPlainText -SecureString $Right
        $result = [System.String]::Equals($leftPlain, $rightPlain, [System.StringComparison]::Ordinal)
        $leftPlain = $null
        $rightPlain = $null
        return $result
    }

    # PSCredential: user name case-insensitive AND password (delegated to the SecureString branch).
    if ($Left -is [System.Management.Automation.PSCredential] -or $Right -is [System.Management.Automation.PSCredential])
    {
        if (-not ($Left -is [System.Management.Automation.PSCredential] -and $Right -is [System.Management.Automation.PSCredential]))
        {
            return $false
        }
        if (-not [System.String]::Equals($Left.UserName, $Right.UserName, [System.StringComparison]::OrdinalIgnoreCase))
        {
            return $false
        }
        return (Test-MSCloudLoginParameterValueEqual -KeyName "$KeyName.Password" -Left $Left.Password -Right $Right.Password)
    }

    # Dictionaries (e.g. Endpoints): key-wise recursive comparison.
    if ($Left -is [System.Collections.IDictionary] -or $Right -is [System.Collections.IDictionary])
    {
        if (-not ($Left -is [System.Collections.IDictionary] -and $Right -is [System.Collections.IDictionary]))
        {
            return $false
        }
        if ($Left.Count -ne $Right.Count)
        {
            return $false
        }
        foreach ($dictKey in $Left.Keys)
        {
            if (-not $Right.Contains($dictKey))
            {
                return $false
            }
            if (-not (Test-MSCloudLoginParameterValueEqual -KeyName "$KeyName.$dictKey" -Left $Left[$dictKey] -Right $Right[$dictKey]))
            {
                return $false
            }
        }
        return $true
    }

    # Collections: AccessTokens are positional and therefore order-sensitive,
    # CmdletsToLoad are cmdlet names and therefore order- and case-insensitive.
    $leftIsCollection = ($Left -is [System.Collections.IEnumerable] -and $Left -isnot [System.String])
    $rightIsCollection = ($Right -is [System.Collections.IEnumerable] -and $Right -isnot [System.String])
    if ($leftIsCollection -or $rightIsCollection)
    {
        $leftArray = @($Left)
        $rightArray = @($Right)
        if ($leftArray.Count -ne $rightArray.Count)
        {
            return $false
        }
        if ($KeyName -eq 'CmdletsToLoad')
        {
            $leftArray = @($leftArray | Sort-Object)
            $rightArray = @($rightArray | Sort-Object)
        }
        for ($i = 0; $i -lt $leftArray.Count; $i++)
        {
            if (-not (Test-MSCloudLoginParameterValueEqual -KeyName $KeyName -Left $leftArray[$i] -Right $rightArray[$i]))
            {
                return $false
            }
        }
        return $true
    }

    # Booleans / switches.
    if ($Left -is [System.Boolean] -or $Left -is [System.Management.Automation.SwitchParameter] -or `
            $Right -is [System.Boolean] -or $Right -is [System.Management.Automation.SwitchParameter])
    {
        return ([System.Boolean]$Left) -eq ([System.Boolean]$Right)
    }

    # Strings: secrets are case-sensitive, identifiers are not.
    $comparison = [System.StringComparison]::OrdinalIgnoreCase
    if ($KeyName -in @('ApplicationSecret', 'AccessTokens') -or $KeyName -like '*.Password')
    {
        $comparison = [System.StringComparison]::Ordinal
    }
    return [System.String]::Equals([System.String]$Left, [System.String]$Right, $comparison)
}
