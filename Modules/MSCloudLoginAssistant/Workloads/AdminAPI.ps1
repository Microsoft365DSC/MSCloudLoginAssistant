function Connect-MSCloudLoginAdminAPI
{
    [CmdletBinding()]
    param()

    Connect-MSCloudLoginRESTWorkload -WorkloadName 'AdminAPI' `
        -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.AdminAPI.AuthorizationUrl `
        -Scope $Script:MSCloudLoginConnectionProfile.AdminAPI.Scope `
        -ClientId $Script:MSCloudLoginConnectionProfile.AdminAPI.ApplicationId `
        -SupportedAuthMethods @('AccessTokens', 'Credentials', 'CredentialsWithApplicationId', 'CredentialsWithTenantId', 'Identity', 'ServicePrincipalWithPath', 'ServicePrincipalWithSecret', 'ServicePrincipalWithThumbprint')
}

function Disconnect-MSCloudLoginAdminAPI
{
    [CmdletBinding()]
    param()

    Disconnect-MSCloudLoginRESTWorkload -WorkloadName 'AdminAPI'
}
