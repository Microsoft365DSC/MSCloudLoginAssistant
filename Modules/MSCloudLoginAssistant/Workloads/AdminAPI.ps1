function Connect-MSCloudLoginAdminAPI
{
    [CmdletBinding()]
    param()

    $InformationPreference = 'SilentlyContinue'
    $ProgressPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginAdminAPI'

    if ($Script:MSCloudLoginConnectionProfile.AdminAPI.Connected)
    {
        if (($Script:MSCloudLoginConnectionProfile.AdminAPI.AuthenticationType -eq 'ServicePrincipalWithSecret' `
                    -or $Script:MSCloudLoginConnectionProfile.AdminAPI.AuthenticationType -eq 'Identity') `
                -and (Get-Date -Date $Script:MSCloudLoginConnectionProfile.AdminAPI.ConnectedDateTime) -lt [System.DateTime]::Now.AddMinutes(-50))
        {
            Add-MSCloudLoginAssistantEvent -Message 'Token is about to expire, renewing' -Source $source
            $Script:MSCloudLoginConnectionProfile.AdminAPI.Connected = $false
        }
    }

    try
    {
        if ($Script:MSCloudLoginConnectionProfile.AdminAPI.AuthenticationType -eq 'CredentialsWithApplicationId' -or
            $Script:MSCloudLoginConnectionProfile.AdminAPI.AuthenticationType -eq 'Credentials' -or
            $Script:MSCloudLoginConnectionProfile.AdminAPI.AuthenticationType -eq 'CredentialsWithTenantId')
        {
            Add-MSCloudLoginAssistantEvent -Message 'Will try connecting with user credentials' -Source $source
            Connect-MSCloudLoginAdminAPIWithUser
        }
        elseif ($Script:MSCloudLoginConnectionProfile.AdminAPI.AuthenticationType -eq 'ServicePrincipalWithThumbprint')
        {
            Add-MSCloudLoginAssistantEvent -Message "Attempting to connect to Admin API using AAD App {$ApplicationID}" -Source $source
            Connect-MSCloudLoginAdminAPIWithCertificateThumbprint
        }
        elseif ($Script:MSCloudLoginConnectionProfile.AdminAPI.AuthenticationType -eq 'AccessTokens')
        {
            Add-MSCloudLoginAssistantEvent -Message 'Using provided access token to connect to Admin API' -Source $source
            $accessToken = if ($Script:MSCloudLoginConnectionProfile.AdminAPI.AccessTokens[0] -like 'Bearer *')
            {
                $Script:MSCloudLoginConnectionProfile.AdminAPI.AccessTokens[0]
            }
            else
            {
                'Bearer ' + $Script:MSCloudLoginConnectionProfile.AdminAPI.AccessTokens[0]
            }
            $Script:MSCloudLoginConnectionProfile.AdminAPI.AccessToken = $accessToken
        }
        elseif ($Script:MSCloudLoginConnectionProfile.AdminAPI.AuthenticationType -eq 'Identity')
        {
            Add-MSCloudLoginAssistantEvent -Message 'Attempting to connect to Admin API using Managed Identity' -Source $source
            $accessToken = Get-AuthToken -Resource $Script:MSCloudLoginConnectionProfile.AdminAPI.Resource -Identity
            $Script:MSCloudLoginConnectionProfile.AdminAPI.AccessToken = 'Bearer ' + $accessToken
        }
        else
        {
            throw 'Specified authentication method is not supported.'
        }

        $Script:MSCloudLoginConnectionProfile.AdminAPI.ConnectedDateTime = [System.DateTime]::Now.ToString()
        $Script:MSCloudLoginConnectionProfile.AdminAPI.Connected = $true
        $Script:MSCloudLoginConnectionProfile.AdminAPI.MultiFactorAuthentication = $false
        Add-MSCloudLoginAssistantEvent -Message "Successfully connected to Admin API using AAD App {$ApplicationID}" -Source $source
    }
    catch
    {
        throw $_
    }
}

function Connect-MSCloudLoginAdminAPIWithUser
{
    [CmdletBinding()]
    param()

    $source = 'Connect-MSCloudLoginAdminAPIWithUser'

    if ([System.String]::IsNullOrEmpty($Script:MSCloudLoginConnectionProfile.AdminAPI.TenantId))
    {
        $tenantId = $Script:MSCloudLoginConnectionProfile.AdminAPI.Credentials.UserName.Split('@')[1]
    }
    else
    {
        $tenantId = $Script:MSCloudLoginConnectionProfile.AdminAPI.TenantId
    }

    try
    {
        $managementToken = Get-AuthToken -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.AdminAPI.AuthorizationUrl `
            -Credentials $Script:MSCloudLoginConnectionProfile.AdminAPI.Credentials `
            -TenantId $tenantId `
            -ClientId $Script:MSCloudLoginConnectionProfile.AdminAPI.ApplicationId `
            -Resource $Script:MSCloudLoginConnectionProfile.AdminAPI.Resource

        $Script:MSCloudLoginConnectionProfile.AdminAPI.AccessToken = $managementToken.token_type.ToString() + ' ' + $managementToken.access_token.ToString()
        $Script:MSCloudLoginConnectionProfile.AdminAPI.Connected = $true
        $Script:MSCloudLoginConnectionProfile.AdminAPI.ConnectedDateTime = [System.DateTime]::Now.ToString()
    }
    catch
    {
        if ($_.ErrorDetails.Message -like '*AADSTS50076*')
        {
            Add-MSCloudLoginAssistantEvent -Message 'Account used required MFA' -Source $source
            Connect-MSCloudLoginAdminAPIWithUserMFA
        }
        else
        {
            $Script:MSCloudLoginConnectionProfile.AdminAPI.Connected = $false
            throw $_
        }
    }
}
function Connect-MSCloudLoginAdminAPIWithUserMFA
{
    [CmdletBinding()]
    param()

    if ([System.String]::IsNullOrEmpty($Script:MSCloudLoginConnectionProfile.AdminAPI.TenantId))
    {
        $tenantid = $Script:MSCloudLoginConnectionProfile.AdminAPI.Credentials.UserName.Split('@')[1]
    }
    else
    {
        $tenantId = $Script:MSCloudLoginConnectionProfile.AdminAPI.TenantId
    }

    $managementToken = Get-AuthToken -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.AdminAPI.AuthorizationUrl `
        -Credentials $Script:MSCloudLoginConnectionProfile.AdminAPI.Credentials `
        -TenantId $tenantId `
        -ClientId $Script:MSCloudLoginConnectionProfile.AdminAPI.ApplicationId `
        -Resource $Script:MSCloudLoginConnectionProfile.AdminAPI.Resource `
        -DeviceCode

    $Script:MSCloudLoginConnectionProfile.AdminAPI.AccessToken = $managementToken.token_type.ToString() + ' ' + $managementToken.access_token.ToString()
    $Script:MSCloudLoginConnectionProfile.AdminAPI.Connected = $true
    $Script:MSCloudLoginConnectionProfile.AdminAPI.MultiFactorAuthentication = $true
    $Script:MSCloudLoginConnectionProfile.AdminAPI.ConnectedDateTime = [System.DateTime]::Now.ToString()
}

function Connect-MSCloudLoginAdminAPIWithCertificateThumbprint
{
    [CmdletBinding()]
    param()

    $ProgressPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginAdminAPIWithCertificateThumbprint'

    Add-MSCloudLoginAssistantEvent -Message 'Attempting to connect to AdminAPI using CertificateThumbprint' -Source $source
    $tenantId = $Script:MSCloudLoginConnectionProfile.AdminAPI.TenantId

    try
    {
        $request = Get-AuthToken -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.AdminAPI.AuthorizationUrl `
            -CertificateThumbprint $Script:MSCloudLoginConnectionProfile.AdminAPI.CertificateThumbprint `
            -ClientId $Script:MSCloudLoginConnectionProfile.AdminAPI.ApplicationId `
            -Resource $Script:MSCloudLoginConnectionProfile.AdminAPI.Resource `
            -TenantId $tenantId

        $Script:MSCloudLoginConnectionProfile.AdminAPI.AccessToken = 'Bearer ' + $Request.access_token
        Add-MSCloudLoginAssistantEvent -Message 'Successfully connected to the Admin API using Certificate Thumbprint' -Source $source

        $Script:MSCloudLoginConnectionProfile.AdminAPI.Connected = $true
        $Script:MSCloudLoginConnectionProfile.AdminAPI.ConnectedDateTime = [System.DateTime]::Now.ToString()
    }
    catch
    {
        throw $_
    }
}
