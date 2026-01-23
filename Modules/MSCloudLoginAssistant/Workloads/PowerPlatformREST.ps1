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

    if (-not $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AccessToken)
    {
        try
        {
            if ($Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AuthenticationType -eq 'CredentialsWithApplicationId' -or
                $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AuthenticationType -eq 'Credentials' -or
                $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AuthenticationType -eq 'CredentialsWithTenantId')
            {
                Add-MSCloudLoginAssistantEvent -Message 'Will try connecting with user credentials' -Source $source
                Connect-MSCloudLoginPowerPlatformRESTWithUser
            }
            elseif ($Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AuthenticationType -eq 'ServicePrincipalWithThumbprint')
            {
                Add-MSCloudLoginAssistantEvent -Message "Attempting to connect to Admin API using AAD App {$ApplicationID}" -Source $source
                Connect-MSCloudLoginPowerPlatformRESTWithCertificateThumbprint
            }
            else
            {
                throw 'Specified authentication method is not supported.'
            }

            $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.ConnectedDateTime = [System.DateTime]::Now.ToString()
            $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.Connected = $true
            $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.MultiFactorAuthentication = $false
            Add-MSCloudLoginAssistantEvent -Message "Successfully connected to Admin API using AAD App {$ApplicationID}" -Source $source
        }
        catch
        {
            throw $_
        }
    }
}

function Connect-MSCloudLoginPowerPlatformRESTWithUser
{
    [CmdletBinding()]
    param()

    $source = 'Connect-MSCloudLoginPowerPlatformRESTWithUser'

    if ([System.String]::IsNullOrEmpty($Script:MSCloudLoginConnectionProfile.PowerPlatformREST.TenantId))
    {
        $tenantId = $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.Credentials.UserName.Split('@')[1]
    }
    else
    {
        $tenantId = $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.TenantId
    }

    try
    {
        $managementToken = Get-AuthToken -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AuthorizationUrl `
            -Credentials $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.Credentials `
            -TenantId $tenantId `
            -ClientId $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.ClientId `
            -Scope $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.Scope

        $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AccessToken = $managementToken.token_type.ToString() + ' ' + $managementToken.access_token.ToString()
        $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.Connected = $true
        $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.ConnectedDateTime = [System.DateTime]::Now.ToString()
    }
    catch
    {
        if ($_.ErrorDetails.Message -like '*AADSTS50076*')
        {
            Add-MSCloudLoginAssistantEvent -Message 'Account used required MFA' -Source $source
            Connect-MSCloudLoginPowerPlatformRESTWithUserMFA
        }
        else
        {
            $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.Connected = $false
            throw $_
        }
    }
}
function Connect-MSCloudLoginPowerPlatformRESTWithUserMFA
{
    [CmdletBinding()]
    param()

    if ([System.String]::IsNullOrEmpty($Script:MSCloudLoginConnectionProfile.PowerPlatformREST.TenantId))
    {
        $tenantId = $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.Credentials.UserName.Split('@')[1]
    }
    else
    {
        $tenantId = $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.TenantId
    }

    $managementToken = Get-AuthToken -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AuthorizationUrl `
        -Credentials $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.Credentials `
        -TenantId $tenantId `
        -ClientId $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.ClientId `
        -Scope $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.Scope `
        -DeviceCode

    $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AccessToken = $managementToken.token_type.ToString() + ' ' + $managementToken.access_token.ToString()
    $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.Connected = $true
    $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.MultiFactorAuthentication = $true
    $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.ConnectedDateTime = [System.DateTime]::Now.ToString()
}

function Connect-MSCloudLoginPowerPlatformRESTWithCertificateThumbprint
{
    [CmdletBinding()]
    param()

    $ProgressPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginPowerPlatformRESTWithCertificateThumbprint'

    Add-MSCloudLoginAssistantEvent -Message 'Attempting to connect to PowerPlatformREST using CertificateThumbprint' -Source $source
    $tenantId = $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.TenantId

    try
    {
        $request = Get-AuthToken -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AuthorizationUrl `
            -CertificateThumbprint $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.CertificateThumbprint `
            -TenantId $tenantId `
            -ClientId $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.ApplicationId `
            -Scope $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.Scope

        $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AccessToken = 'Bearer ' + $Request.access_token
        Add-MSCloudLoginAssistantEvent -Message 'Successfully connected to the Admin API API using Certificate Thumbprint' -Source $source

        $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.Connected = $true
        $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.ConnectedDateTime = [System.DateTime]::Now.ToString()
    }
    catch
    {
        throw $_
    }
}

function Disconnect-MSCloudLoginPowerPlatformREST
{
    [CmdletBinding()]
    param()

    $source = 'Disconnect-MSCloudLoginPowerPlatformREST'

    if ($Script:MSCloudLoginConnectionProfile.PowerPlatformREST.Connected)
    {
        Add-MSCloudLoginAssistantEvent -Message 'Attempting to disconnect from PowerPlatformREST API' -Source $source
        $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.Connected = $false
        Add-MSCloudLoginAssistantEvent -Message 'Successfully disconnected from PowerPlatformREST API' -Source $source
    }
    else
    {
        Add-MSCloudLoginAssistantEvent -Message 'No connections to PowerPlatformREST API were found' -Source $source
    }
}
