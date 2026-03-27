function Connect-MSCloudLoginEngageHub
{
    [CmdletBinding()]
    param()

    Connect-MSCloudLoginRESTWorkload -WorkloadName 'EngageHub' `
        -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.EngageHub.AuthorizationUrl `
        -Scope $Script:MSCloudLoginConnectionProfile.EngageHub.Scope `
        -ClientId $Script:MSCloudLoginConnectionProfile.EngageHub.ApplicationId `
        -SupportedAuthMethods @('AccessTokens', 'Credentials', 'CredentialsWithApplicationId', 'CredentialsWithTenantId', 'Identity', 'ServicePrincipalWithPath', 'ServicePrincipalWithSecret', 'ServicePrincipalWithThumbprint')
}

function Disconnect-MSCloudLoginEngageHub
{
    [CmdletBinding()]
    param()

    Disconnect-MSCloudLoginRESTWorkload -WorkloadName 'EngageHub'
}
