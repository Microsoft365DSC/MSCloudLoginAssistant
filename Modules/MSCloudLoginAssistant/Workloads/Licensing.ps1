function Connect-MSCloudLoginLicensing
{
    [CmdletBinding()]
    param()

    Connect-MSCloudLoginRESTWorkload -WorkloadName 'Licensing' `
        -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.Licensing.AuthorizationUrl `
        -ClientId $Script:MSCloudLoginConnectionProfile.Licensing.ApplicationId `
        -Scope $Script:MSCloudLoginConnectionProfile.Licensing.Scope `
        -SupportedAuthMethods @('AccessTokens', 'Credentials', 'CredentialsWithApplicationId', 'CredentialsWithTenantId', 'Identity', 'ServicePrincipalWithPath', 'ServicePrincipalWithSecret', 'ServicePrincipalWithThumbprint')
}

function Disconnect-MSCloudLoginLicensing
{
    [CmdletBinding()]
    param()

    Disconnect-MSCloudLoginRESTWorkload -WorkloadName 'Licensing'
}
