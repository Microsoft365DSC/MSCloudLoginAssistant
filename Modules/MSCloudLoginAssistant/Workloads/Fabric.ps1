function Connect-MSCloudLoginFabric
{
    [CmdletBinding()]
    param()

    Connect-MSCloudLoginRESTWorkload -WorkloadName 'Fabric' `
        -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.Fabric.AuthorizationUrl `
        -ClientId $Script:MSCloudLoginConnectionProfile.Fabric.ApplicationId `
        -Scope $Script:MSCloudLoginConnectionProfile.Fabric.Scope `
        -SupportedAuthMethods @('AccessTokens', 'Credentials', 'CredentialsWithApplicationId', 'CredentialsWithTenantId', 'Identity', 'ServicePrincipalWithPath', 'ServicePrincipalWithSecret', 'ServicePrincipalWithThumbprint')
}

function Disconnect-MSCloudLoginFabric
{
    [CmdletBinding()]
    param()

    Disconnect-MSCloudLoginRESTWorkload -WorkloadName 'Fabric'
}
