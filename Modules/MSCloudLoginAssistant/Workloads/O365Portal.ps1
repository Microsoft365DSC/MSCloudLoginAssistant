function Connect-MSCloudLoginO365Portal
{
    [CmdletBinding()]
    param()

    $InformationPreference = 'SilentlyContinue'
    $ProgressPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginO365Portal'

    if ($Script:MSCloudLoginConnectionProfile.O365Portal.Connected)
    {
        if (($Script:MSCloudLoginConnectionProfile.O365Portal.AuthenticationType -eq 'ServicePrincipalWithSecret' `
                    -or $Script:MSCloudLoginConnectionProfile.O365Portal.AuthenticationType -eq 'Identity') `
                -and (Get-Date -Date $Script:MSCloudLoginConnectionProfile.O365Portal.ConnectedDateTime) -lt [System.DateTime]::Now.AddMinutes(-50))
        {
            Add-MSCloudLoginAssistantEvent -Message 'Token is about to expire, renewing' -Source $source
            $Script:MSCloudLoginConnectionProfile.O365Portal.Connected = $false
        }
    }

    try
    {
        if ($Script:MSCloudLoginConnectionProfile.O365Portal.AuthenticationType -eq 'CredentialsWithApplicationId' -or
            $Script:MSCloudLoginConnectionProfile.O365Portal.AuthenticationType -eq 'Credentials' -or
            $Script:MSCloudLoginConnectionProfile.O365Portal.AuthenticationType -eq 'CredentialsWithTenantId')
        {
            Add-MSCloudLoginAssistantEvent -Message 'Will try connecting with user credentials' -Source $source
            Connect-MSCloudLoginO365PortalWithUser
        }
        elseif ($Script:MSCloudLoginConnectionProfile.O365Portal.AuthenticationType -eq 'AccessTokens')
        {
            Add-MSCloudLoginAssistantEvent -Message 'Using provided access token to connect to O365 Portal' -Source $source
            $accessToken = if ($Script:MSCloudLoginConnectionProfile.O365Portal.AccessTokens[0] -like 'Bearer *')
            {
                $Script:MSCloudLoginConnectionProfile.O365Portal.AccessTokens[0]
            }
            else
            {
                'Bearer ' + $Script:MSCloudLoginConnectionProfile.O365Portal.AccessTokens[0]
            }
            $Script:MSCloudLoginConnectionProfile.O365Portal.AccessToken = $accessToken
        }
        else
        {
            throw 'Specified authentication method is not supported.'
        }

        $Script:MSCloudLoginConnectionProfile.O365Portal.ConnectedDateTime = [System.DateTime]::Now.ToString()
        $Script:MSCloudLoginConnectionProfile.O365Portal.Connected = $true
        $Script:MSCloudLoginConnectionProfile.O365Portal.MultiFactorAuthentication = $false
        Add-MSCloudLoginAssistantEvent -Message "Successfully connected to O365 Portal using AAD App {$ApplicationID}" -Source $source
    }
    catch
    {
        throw $_
    }
}

function Connect-MSCloudLoginO365PortalWithUser
{
    [CmdletBinding()]
    param()

    $source = 'Connect-MSCloudLoginO365PortalWithUser'

    if ([System.String]::IsNullOrEmpty($Script:MSCloudLoginConnectionProfile.O365Portal.TenantId))
    {
        $tenantId = $Script:MSCloudLoginConnectionProfile.O365Portal.Credentials.UserName.Split('@')[1]
    }
    else
    {
        $tenantId = $Script:MSCloudLoginConnectionProfile.O365Portal.TenantId
    }

    try
    {
        $managementToken = Get-AuthToken -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.O365Portal.AuthorizationUrl `
            -Credentials $Script:MSCloudLoginConnectionProfile.O365Portal.Credentials `
            -TenantId $tenantId `
            -ClientId $Script:MSCloudLoginConnectionProfile.O365Portal.ApplicationId `
            -Scope $Script:MSCloudLoginConnectionProfile.O365Portal.Scope

        $Script:MSCloudLoginConnectionProfile.O365Portal.AccessToken = $managementToken.token_type.ToString() + ' ' + $managementToken.access_token.ToString()
        $Script:MSCloudLoginConnectionProfile.O365Portal.Connected = $true
        $Script:MSCloudLoginConnectionProfile.O365Portal.ConnectedDateTime = [System.DateTime]::Now.ToString()
    }
    catch
    {
        if ($_.ErrorDetails.Message -like '*AADSTS50076*')
        {
            Add-MSCloudLoginAssistantEvent -Message 'Account used required MFA' -Source $source
            Connect-MSCloudLoginO365PortalWithUserMFA
        }
    }
}
function Connect-MSCloudLoginO365PortalWithUserMFA
{
    [CmdletBinding()]
    param()

    if ([System.String]::IsNullOrEmpty($Script:MSCloudLoginConnectionProfile.O365Portal.TenantId))
    {
        $tenantid = $Script:MSCloudLoginConnectionProfile.O365Portal.Credentials.UserName.Split('@')[1]
    }
    else
    {
        $tenantId = $Script:MSCloudLoginConnectionProfile.O365Portal.TenantId
    }

    $managementToken = Get-AuthToken -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.O365Portal.AuthorizationUrl `
        -Credentials $Script:MSCloudLoginConnectionProfile.O365Portal.Credentials `
        -TenantId $tenantId `
        -ClientId $Script:MSCloudLoginConnectionProfile.O365Portal.ApplicationId `
        -Scope $Script:MSCloudLoginConnectionProfile.O365Portal.Scope `
        -DeviceCode

    $Script:MSCloudLoginConnectionProfile.O365Portal.AccessToken = $managementToken.token_type.ToString() + ' ' + $managementToken.access_token.ToString()
    $Script:MSCloudLoginConnectionProfile.O365Portal.Connected = $true
    $Script:MSCloudLoginConnectionProfile.O365Portal.MultiFactorAuthentication = $true
    $Script:MSCloudLoginConnectionProfile.O365Portal.ConnectedDateTime = [System.DateTime]::Now.ToString()
}
