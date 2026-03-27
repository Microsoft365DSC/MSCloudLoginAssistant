# Stubs for MSCloudLoginAssistant unit tests.
# These stub functions prevent calls to real external modules during testing.

function Connect-AzAccount
{
    [CmdletBinding()]
    param (
        [Parameter()] [PSCredential] $Credential,
        [Parameter()] [String]       $TenantId,
        [Parameter()] [String]       $Subscription,
        [Parameter()] [String]       $Environment,
        [Parameter()] [Switch]       $Identity,
        [Parameter()] [Switch]       $ServicePrincipal,
        [Parameter()] [String]       $CertificateThumbprint,
        [Parameter()] [String]       $ApplicationId,
        [Parameter()] [SecureString] $CertificatePassword,
        [Parameter()] [String]       $CertificatePath
    )
}

function Disconnect-AzAccount
{
    [CmdletBinding()]
    param (
        [Parameter()] [String] $Username,
        [Parameter()] [String] $TenantId
    )
}

function Connect-ExchangeOnline
{
    [CmdletBinding()]
    param (
        [Parameter()] [PSCredential] $Credential,
        [Parameter()] [String]       $ConnectionUri,
        [Parameter()] [String]       $AzureADAuthorizationEndpointUri,
        [Parameter()] [String]       $ExchangeEnvironmentName,
        [Parameter()] [String]       $Organization,
        [Parameter()] [String]       $AppId,
        [Parameter()] [String]       $CertificateThumbprint,
        [Parameter()] [String]       $CertificateFilePath,
        [Parameter()] [SecureString] $CertificatePassword,
        [Parameter()] [Switch]       $ManagedIdentity,
        [Parameter()] [String[]]     $CommandName,
        [Parameter()] [Switch]       $ShowBanner
    )
}

function Disconnect-ExchangeOnline
{
    [CmdletBinding()]
    param (
        [Parameter()] [Switch] $Confirm
    )
}

function Connect-MgGraph
{
    [CmdletBinding()]
    param (
        [Parameter()] [String]       $TenantId,
        [Parameter()] [String[]]     $Scopes,
        [Parameter()] [String]       $ClientId,
        [Parameter()] [String]       $CertificateThumbprint,
        [Parameter()] [String]       $Environment,
        [Parameter()] [SecureString] $AccessToken
    )
}

function Disconnect-MgGraph
{
    [CmdletBinding()]
    param ()
}

function Connect-PnPOnline
{
    [CmdletBinding()]
    param (
        [Parameter()] [String]       $Url,
        [Parameter()] [PSCredential] $Credentials,
        [Parameter()] [String]       $ClientId,
        [Parameter()] [String]       $Tenant,
        [Parameter()] [String]       $Thumbprint,
        [Parameter()] [String]       $CertificatePath,
        [Parameter()] [SecureString] $CertificatePassword,
        [Parameter()] [String]       $AzureEnvironment,
        [Parameter()] [Switch]       $ManagedIdentity,
        [Parameter()] [SecureString] $AccessToken,
        [Parameter()] [String]       $Region,
        [Parameter()] [Switch]       $ForceAuthentication
    )
}

function Disconnect-PnPOnline
{
    [CmdletBinding()]
    param ()
}

function Connect-MicrosoftTeams
{
    [CmdletBinding()]
    param (
        [Parameter()] [PSCredential] $Credential,
        [Parameter()] [String]       $TenantId,
        [Parameter()] [String]       $ApplicationId,
        [Parameter()] [String]       $CertificateThumbprint,
        [Parameter()] [Switch]       $Identity,
        [Parameter()] [SecureString] $AccessTokens
    )
}

function Disconnect-MicrosoftTeams
{
    [CmdletBinding()]
    param ()
}

function Connect-IPPSSession
{
    [CmdletBinding()]
    param (
        [Parameter()] [PSCredential] $Credential,
        [Parameter()] [String]       $ConnectionUri,
        [Parameter()] [String]       $AzureADAuthorizationEndpointUri,
        [Parameter()] [String]       $AppId,
        [Parameter()] [String]       $Organization,
        [Parameter()] [String]       $CertificateThumbprint,
        [Parameter()] [String]       $CertificateFilePath,
        [Parameter()] [SecureString] $CertificatePassword,
        [Parameter()] [String[]]     $CommandName
    )
}

function Add-PowerAppsAccount
{
    [CmdletBinding()]
    param (
        [Parameter()] [String] $Endpoint,
        [Parameter()] [String] $TenantID,
        [Parameter()] [String] $ApplicationId,
        [Parameter()] [String] $ClientSecret,
        [Parameter()] [String] $CertificateThumbprint
    )
}

function Remove-PowerAppsAccount
{
    [CmdletBinding()]
    param ()
}

function Get-ConnectionInformation
{
    [CmdletBinding()]
    param ()
}

function Get-PnPContext
{
    [CmdletBinding()]
    param ()
}

function Get-PnPConnection
{
    [CmdletBinding()]
    param ()
}

function Get-MgBetaOrganization
{
    [CmdletBinding()]
    param ()
}

function Get-AzAccessToken
{
    [CmdletBinding()]
    param (
        [Parameter()] [String] $ResourceUrl,
        [Parameter()] [Switch] $AsSecureString
    )
}
