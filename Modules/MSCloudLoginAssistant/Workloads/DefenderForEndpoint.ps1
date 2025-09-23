function Connect-MSCloudLoginDefenderForEndpoint
{
    [CmdletBinding()]
    param()

    $ProgressPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginDefenderForEndpoint'

    if ($Script:MSCloudLoginConnectionProfile.DefenderForEndpoint.AuthenticationType -eq 'ServicePrincipalWithSecret')
    {
        Add-MSCloudLoginAssistantEvent -Message 'Will try connecting with Application Secret' -Source $source
        Connect-MSCloudLoginDefenderForEndpointWithAppSecret
    }
    elseif ($Script:MSCloudLoginConnectionProfile.DefenderForEndpoint.AuthenticationType -eq 'ServicePrincipalWithThumbprint')
    {
        Add-MSCloudLoginAssistantEvent -Message 'Will try connecting with Application Secret' -Source $source
        Connect-MSCloudLoginDefenderForEndpointWithCertificateThumbprint
    }
    elseif ($Script:MSCloudLoginConnectionProfile.DefenderForEndpoint.AuthenticationType -eq 'AccessToken')
    {
        Add-MSCloudLoginAssistantEvent -Message 'Will try connecting with Access Token' -Source $source
        $Ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToCoTaskMemUnicode($Script:MSCloudLoginConnectionProfile.DefenderForEndpoint.AccessTokens[0])
        $AccessTokenValue = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($Ptr)
        [System.Runtime.InteropServices.Marshal]::ZeroFreeCoTaskMemUnicode($Ptr)
        $Script:MSCloudLoginConnectionProfile.DefenderForEndpoint.AccessToken = $AccessTokenValue
    }
}

function Connect-MSCloudLoginDefenderForEndpointWithAppSecret
{
    [CmdletBinding()]
    param()

    $ProgressPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginDefenderForEndpointWithAppSecret'

    $managementToken = Get-AuthToken -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.DefenderForEndpoint.AuthorizationUrl `
        -ClientId $Script:MSCloudLoginConnectionProfile.DefenderForEndpoint.ApplicationId `
        -ClientSecret $Script:MSCloudLoginConnectionProfile.DefenderForEndpoint.ApplicationSecret `
        -TenantId $Script:MSCloudLoginConnectionProfile.DefenderForEndpoint.TenantId `
        -Scope $Script:MSCloudLoginConnectionProfile.DefenderForEndpoint.Scope

    Add-MSCloudLoginAssistantEvent -Message 'Successfully connected to the DefenderForEndpoint API using Application Secret' -Source $source
    $Script:MSCloudLoginConnectionProfile.DefenderForEndpoint.AccessToken = $managementToken.token_type.ToString() + ' ' + $managementToken.access_token.ToString()
}

function Connect-MSCloudLoginDefenderForEndpointWithCertificateThumbprint
{
    [CmdletBinding()]
    param()

    $ProgressPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginDefenderForEndpointWithCertificateThumbprint'

    try
    {
        $request = Get-AuthToken -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.DefenderForEndpoint.AuthorizationUrl `
            -CertificateThumbprint $Script:MSCloudLoginConnectionProfile.DefenderForEndpoint.CertificateThumbprint `
            -TenantId $Script:MSCloudLoginConnectionProfile.DefenderForEndpoint.TenantId `
            -ClientId $Script:MSCloudLoginConnectionProfile.DefenderForEndpoint.ApplicationId `
            -Scope $Script:MSCloudLoginConnectionProfile.DefenderForEndpoint.Scope

        $Script:MSCloudLoginConnectionProfile.DefenderForEndpoint.AccessToken = 'Bearer ' + $request.access_token
        Add-MSCloudLoginAssistantEvent -Message 'Successfully connected to the DefenderForEndpoint API using Certificate Thumbprint' -Source $source
    }
    catch
    {
        throw $_
    }
}
