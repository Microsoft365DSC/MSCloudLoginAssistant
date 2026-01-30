function Connect-MSCloudLoginFabric
{
    [CmdletBinding()]
    param()

    Connect-MSCloudLoginRESTWorkload -WorkloadName 'Fabric' `
        -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.Fabric.AuthorizationUrl `
        -Scope $Script:MSCloudLoginConnectionProfile.Fabric.Scope `
        -SupportedAuthMethods @('AccessTokens', 'Credentials', 'CredentialsWithApplicationId', 'CredentialsWithTenantId', 'Identity', 'ServicePrincipalWithPath', 'ServicePrincipalWithSecret', 'ServicePrincipalWithThumbprint')
}
