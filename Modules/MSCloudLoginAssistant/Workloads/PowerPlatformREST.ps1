function Connect-MSCloudLoginPowerPlatformREST
{
    [CmdletBinding()]
    param()

    $InformationPreference = 'SilentlyContinue'
    $ProgressPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginPowerPlatformREST'

    # Test authentication to make sure the token hasn't expired. Only probe when we
    # believe we are connected and already have a token - probing without one is a
    # guaranteed 401 on the first connection attempt.
    if ($Script:MSCloudLoginConnectionProfile.PowerPlatformREST.Connected -and `
        -not [System.String]::IsNullOrEmpty($Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AccessToken))
    {
        try
        {
            $uri = "https://" + $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.BapEndpoint + `
                "/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments?api-version=2024-05-01"
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
            # Liveness probe: the token is no longer valid, force a full reconnect.
            Add-MSCloudLoginAssistantEvent -Message "PowerPlatformREST token probe failed, reconnecting: $($_.Exception.Message)" -Source $source
            $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AccessToken = $null
            $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.Connected = $false
        }
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
