# Stubs for MSCloudLoginAssistant unit tests.
# These stub functions prevent calls to real external modules during testing.

function Start-ThreadJob
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)] [ScriptBlock] $ScriptBlock,
        [Parameter()] [Object[]] $ArgumentList
    )
}

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
        [Parameter()] [String]       $CertificatePath,
        [Parameter()] [String]       $AccessToken,
        [Parameter()] [String]       $AccountId
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
        [Parameter()] [Switch]       $ShowBanner,
        [Parameter()] [Switch]       $ShowProgress,
        [Parameter()] [Switch]       $SkipLoadingCmdletHelp,
        [Parameter()] [String]       $AccessToken,
        [Parameter()] [String]       $DelegatedOrganization,
        [Parameter()] [String]       $UserPrincipalName
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
        [Parameter()] [SecureString] $AccessToken,
        [Parameter()] [PSCredential] $ClientSecretCredential,
        [Parameter()] [System.Security.Cryptography.X509Certificates.X509Certificate2] $Certificate,
        [Parameter()] [Switch]       $NoWelcome
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
        [Parameter()] [String]       $AccessToken,
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
        [Parameter()] [String[]]     $AccessTokens,
        [Parameter()] [System.Security.Cryptography.X509Certificates.X509Certificate2] $Certificate,
        [Parameter()] [String]       $TeamsEnvironmentName
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
        [Parameter()] [String[]]     $CommandName,
        [Parameter()] [Switch]       $EnableSearchOnlySession,
        [Parameter()] [Switch]       $ShowBanner
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
        [Parameter()] [String] $CertificateThumbprint,
        [Parameter()] [String] $Username,
        [Parameter()] [String] $Password
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

function Get-AuthToken
{
    [CmdletBinding()]
    param (
        [Parameter()] [String] $AuthorizationUrl,
        [Parameter()] [PSCredential] $Credentials,
        [Parameter()] [String] $TenantId,
        [Parameter()] [String] $ClientId,
        [Parameter()] [String] $ClientSecret,
        [Parameter()] [String] $CertificateThumbprint,
        [Parameter()] [SecureString] $CertificatePassword,
        [Parameter()] [String] $CertificatePath,
        [Parameter()] [Switch] $DeviceCode,
        [Parameter()] [Switch] $Identity,
        [Parameter()] [String] $RefreshToken,
        [Parameter()] [String] $Resource,
        [Parameter()] [String] $Scope,
        [Parameter()] [String] $TokenEndpoint
    )
    return @{ access_token = 'test-token' }
}

function Connect-PnPOnline
{
    [CmdletBinding()]
    param(
        [Parameter()] [String] $Url,
        [Parameter()] [String] $AccessToken,
        [Parameter()] [String] $AzureEnvironment
    )
}

function Disconnect-PnPOnline
{
    [CmdletBinding()]
    param()
}

function Connect-MicrosoftTeams
{
    [CmdletBinding()]
    param(
        [Parameter()] [String]       $ApplicationId,
        [Parameter()] [PSCredential] $Credential,
        [Parameter()] [Switch]       $Identity,
        [Parameter()] [String[]]     $AccessTokens,
        [Parameter()] [System.Security.Cryptography.X509Certificates.X509Certificate2] $Certificate,
        [Parameter()] [String]       $TenantId
    )
}

function Get-CsTeamsCallingPolicy
{
    [CmdletBinding()]
    param()
}
