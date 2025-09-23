function Connect-MSCloudLoginEngageHub
{
    [CmdletBinding()]
    param()

    $InformationPreference = 'SilentlyContinue'
    $ProgressPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginEngageHub'

    if (-not $Script:MSCloudLoginConnectionProfile.EngageHub.AccessToken)
    {
        try
        {
            if ($Script:MSCloudLoginConnectionProfile.EngageHub.AuthenticationType -eq 'CredentialsWithApplicationId' -or
                $Script:MSCloudLoginConnectionProfile.EngageHub.AuthenticationType -eq 'Credentials' -or
                $Script:MSCloudLoginConnectionProfile.EngageHub.AuthenticationType -eq 'CredentialsWithTenantId')
            {
                Add-MSCloudLoginAssistantEvent -Message 'Will try connecting with user credentials' -Source $source
                Connect-MSCloudLoginEngageHubWithUser
            }
            elseif ($Script:MSCloudLoginConnectionProfile.EngageHub.AuthenticationType -eq 'ServicePrincipalWithThumbprint')
            {
                Add-MSCloudLoginAssistantEvent -Message "Attempting to connect to Admin API using AAD App {$ApplicationID}" -Source $source
                Connect-MSCloudLoginEngageHubWithCertificateThumbprint
            }
            else
            {
                throw 'Specified authentication method is not supported.'
            }

            $Script:MSCloudLoginConnectionProfile.EngageHub.ConnectedDateTime = [System.DateTime]::Now.ToString()
            $Script:MSCloudLoginConnectionProfile.EngageHub.Connected = $true
            $Script:MSCloudLoginConnectionProfile.EngageHub.MultiFactorAuthentication = $false
            Add-MSCloudLoginAssistantEvent -Message "Successfully connected to Admin API using AAD App {$ApplicationID}" -Source $source
        }
        catch
        {
            throw $_
        }
    }
}

function Connect-MSCloudLoginEngageHubWithUser
{
    [CmdletBinding()]
    param()

    $source = 'Connect-MSCloudLoginEngageHubWithUser'

    if ([System.String]::IsNullOrEmpty($Script:MSCloudLoginConnectionProfile.EngageHub.TenantId))
    {
        $tenantId = $Script:MSCloudLoginConnectionProfile.EngageHub.Credentials.UserName.Split('@')[1]
    }
    else
    {
        $tenantId = $Script:MSCloudLoginConnectionProfile.EngageHub.TenantId
    }

    # Request token through ROPC
    try
    {
        $managementToken = Get-AuthToken -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.EngageHub.AuthorizationUrl `
            -Credentials $Script:MSCloudLoginConnectionProfile.EngageHub.Credentials `
            -TenantId $tenantId `
            -ClientId $Script:MSCloudLoginConnectionProfile.EngageHub.ApplicationId `
            -Scope $Script:MSCloudLoginConnectionProfile.EngageHub.Scope

        $Script:MSCloudLoginConnectionProfile.EngageHub.AccessToken = $managementToken.token_type.ToString() + ' ' + $managementToken.access_token.ToString()
        $Script:MSCloudLoginConnectionProfile.EngageHub.Connected = $true
        $Script:MSCloudLoginConnectionProfile.EngageHub.ConnectedDateTime = [System.DateTime]::Now.ToString()
    }
    catch
    {
        if ($_.ErrorDetails.Message -like '*AADSTS50076*')
        {
            Add-MSCloudLoginAssistantEvent -Message 'Account used required MFA' -Source $source
            Connect-MSCloudLoginEngageHubWithUserMFA
        }
    }
}
function Connect-MSCloudLoginEngageHubWithUserMFA
{
    [CmdletBinding()]
    param()

    if ([System.String]::IsNullOrEmpty($Script:MSCloudLoginConnectionProfile.EngageHub.TenantId))
    {
        $tenantId = $Script:MSCloudLoginConnectionProfile.EngageHub.Credentials.UserName.Split('@')[1]
    }
    else
    {
        $tenantId = $Script:MSCloudLoginConnectionProfile.EngageHub.TenantId
    }

    $managementToken = Get-AuthToken -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.EngageHub.AuthorizationUrl `
        -Credentials $Script:MSCloudLoginConnectionProfile.EngageHub.Credentials `
        -TenantId $tenantId `
        -ClientId $Script:MSCloudLoginConnectionProfile.EngageHub.ApplicationId `
        -Scope $Script:MSCloudLoginConnectionProfile.EngageHub.Scope

    $Script:MSCloudLoginConnectionProfile.EngageHub.AccessToken = $managementToken.token_type.ToString() + ' ' + $managementToken.access_token.ToString()
    $Script:MSCloudLoginConnectionProfile.EngageHub.Connected = $true
    $Script:MSCloudLoginConnectionProfile.EngageHub.MultiFactorAuthentication = $true
    $Script:MSCloudLoginConnectionProfile.EngageHub.ConnectedDateTime = [System.DateTime]::Now.ToString()
}

function Connect-MSCloudLoginEngageHubWithCertificateThumbprint
{
    [CmdletBinding()]
    param()

    $ProgressPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginEngageHubWithCertificateThumbprint'

    Add-MSCloudLoginAssistantEvent -Message 'Attempting to connect to EngageHub using CertificateThumbprint' -Source $source
    $tenantId = $Script:MSCloudLoginConnectionProfile.EngageHub.TenantId

    try
    {
        $request = Get-AuthToken -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.EngageHub.AuthorizationUrl `
            -CertificateThumbprint $Script:MSCloudLoginConnectionProfile.EngageHub.CertificateThumbprint `
            -TenantId $tenantId `
            -ClientId $Script:MSCloudLoginConnectionProfile.EngageHub.ApplicationId
            -Scope $Script:MSCloudLoginConnectionProfile.EngageHub.Scope

        $Script:MSCloudLoginConnectionProfile.EngageHub.AccessToken = 'Bearer ' + $Request.access_token
        Add-MSCloudLoginAssistantEvent -Message 'Successfully connected to the Admin API API using Certificate Thumbprint' -Source $source

        $Script:MSCloudLoginConnectionProfile.EngageHub.Connected = $true
        $Script:MSCloudLoginConnectionProfile.EngageHub.ConnectedDateTime = [System.DateTime]::Now.ToString()
    }
    catch
    {
        throw $_
    }
}

function Disconnect-MSCloudLoginEngageHub
{
    [CmdletBinding()]
    param()

    $source = 'Disconnect-MSCloudLoginEngageHub'

    if ($Script:MSCloudLoginConnectionProfile.EngageHub.Connected)
    {
        Add-MSCloudLoginAssistantEvent -Message 'Attempting to disconnect from EngageHub API' -Source $source
        $Script:MSCloudLoginConnectionProfile.EngageHub.Connected = $false
        Add-MSCloudLoginAssistantEvent -Message 'Successfully disconnected from EngageHub API' -Source $source
    }
    else
    {
        Add-MSCloudLoginAssistantEvent -Message 'No connections to EngageHub API were found' -Source $source
    }
}
