function Connect-MSCloudLoginPowerPlatformREST
{
    [CmdletBinding()]
    param()

    $InformationPreference = 'SilentlyContinue'
    $ProgressPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginPowerPlatformREST'

    # Test authentication to make sure the token hasn't expired
    try
    {
        $uri = "https://" + $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.BapEndpoint + `
               "/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments"
        $headers = @{
            Authorization = $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AccessToken
        }
        $null = Invoke-WebRequest -Method 'GET' `
            -Uri $Uri `
            -Headers $headers `
            -ContentType 'application/json; charset=utf-8' `
            -UseBasicParsing `
            -ErrorAction Stop
    }
    catch
    {
        $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AccessToken = $null
    }

    Connect-MSCloudLoginRESTWorkload -WorkloadName 'PowerPlatformREST' `
        -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AuthorizationUrl `
        -Scope $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.Scope `
        -ClientId $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.ClientId `
        -SupportedAuthMethods @('AccessTokens', 'Credentials', 'CredentialsWithApplicationId', 'CredentialsWithTenantId', 'Identity', 'ServicePrincipalWithPath', 'ServicePrincipalWithSecret', 'ServicePrincipalWithThumbprint')
}

function Disconnect-MSCloudLoginPowerPlatformREST
{
    [CmdletBinding()]
    param()

    Disconnect-MSCloudLoginRESTWorkload -WorkloadName 'PowerPlatformREST'
}
