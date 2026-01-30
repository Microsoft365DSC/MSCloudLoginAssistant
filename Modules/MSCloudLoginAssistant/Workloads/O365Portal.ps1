function Connect-MSCloudLoginO365Portal
{
    [CmdletBinding()]
    param()

    Connect-MSCloudLoginRESTWorkload -WorkloadName 'O365Portal' `
        -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.O365Portal.AuthorizationUrl `
        -Scope $Script:MSCloudLoginConnectionProfile.O365Portal.Scope `
        -ClientId $Script:MSCloudLoginConnectionProfile.O365Portal.ApplicationId `
        -SupportedAuthMethods @('Credentials', 'CredentialsWithTenantId', 'AccessTokens')
}
