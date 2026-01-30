function Connect-MSCloudLoginSharePointOnlineREST
{
    [CmdletBinding()]
    param()

    Connect-MSCloudLoginRESTWorkload -WorkloadName 'SharePointOnlineREST' `
        -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.AuthorizationUrl `
        -Scope $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.Scope `
        -ClientId $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.ApplicationId `
        -SupportedAuthMethods @('AccessTokens', 'Credentials', 'CredentialsWithApplicationId', 'CredentialsWithTenantId', 'Identity', 'ServicePrincipalWithPath', 'ServicePrincipalWithSecret', 'ServicePrincipalWithThumbprint')
}
