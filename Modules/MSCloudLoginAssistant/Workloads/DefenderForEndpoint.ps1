function Connect-MSCloudLoginDefenderForEndpoint
{
    [CmdletBinding()]
    param()

    Connect-MSCloudLoginRESTWorkload -WorkloadName 'DefenderForEndpoint' `
        -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.DefenderForEndpoint.AuthorizationUrl `
        -ClientId $Script:MSCloudLoginConnectionProfile.DefenderForEndpoint.ApplicationId `
        -Scope $Script:MSCloudLoginConnectionProfile.DefenderForEndpoint.Scope `
        -SupportedAuthMethods @('AccessTokens', 'Credentials', 'CredentialsWithApplicationId', 'CredentialsWithTenantId', 'Identity', 'ServicePrincipalWithSecret', 'ServicePrincipalWithPath', 'ServicePrincipalWithThumbprint')
}

function Disconnect-MSCloudLoginDefenderForEndpoint
{
    [CmdletBinding()]
    param()

    Disconnect-MSCloudLoginRESTWorkload -WorkloadName 'DefenderForEndpoint'
}
