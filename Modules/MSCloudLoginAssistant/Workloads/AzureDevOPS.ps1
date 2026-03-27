function Connect-MSCloudLoginAzureDevOPS
{
    [CmdletBinding()]
    param()

    Connect-MSCloudLoginRESTWorkload -WorkloadName 'AzureDevOPS' `
        -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.AzureDevOPS.AuthorizationUrl `
        -Scope $Script:MSCloudLoginConnectionProfile.AzureDevOPS.Scope `
        -ClientId $Script:MSCloudLoginConnectionProfile.AzureDevOPS.ApplicationId `
        -SupportedAuthMethods @('AccessTokens', 'Credentials', 'CredentialsWithApplicationId', 'CredentialsWithTenantId', 'Identity', 'ServicePrincipalWithPath', 'ServicePrincipalWithSecret', 'ServicePrincipalWithThumbprint')
}

function Disconnect-MSCloudLoginAzureDevOPS
{
    [CmdletBinding()]
    param()

    Disconnect-MSCloudLoginRESTWorkload -WorkloadName 'AzureDevOPS'
}
