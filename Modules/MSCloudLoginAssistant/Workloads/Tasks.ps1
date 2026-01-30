function Connect-MSCloudLoginTasks
{
    [CmdletBinding()]
    param()

    Connect-MSCloudLoginRESTWorkload -WorkloadName 'Tasks' `
        -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.Tasks.AuthorizationUrl `
        -Scope $Script:MSCloudLoginConnectionProfile.Tasks.Scope `
        -ClientId $Script:MSCloudLoginConnectionProfile.Tasks.ApplicationId `
        -SupportedAuthMethods @('AccessTokens', 'Credentials', 'CredentialsWithApplicationId', 'CredentialsWithTenantId', 'Identity', 'ServicePrincipalWithSecret', 'ServicePrincipalWithThumbprint', 'ServicePrincipalWithPath')
}
