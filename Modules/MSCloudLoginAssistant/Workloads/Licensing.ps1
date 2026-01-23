function Connect-MSCloudLoginLicensing
{
    [CmdletBinding()]
    param()

    $InformationPreference = 'SilentlyContinue'
    $ProgressPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginLicensing'

    if (-not $Script:MSCloudLoginConnectionProfile.Licensing.AccessToken)
    {
        try
        {
            if ($Script:MSCloudLoginConnectionProfile.Licensing.AuthenticationType -eq 'CredentialsWithApplicationId' -or
                $Script:MSCloudLoginConnectionProfile.Licensing.AuthenticationType -eq 'Credentials' -or
                $Script:MSCloudLoginConnectionProfile.Licensing.AuthenticationType -eq 'CredentialsWithTenantId')
            {
                Add-MSCloudLoginAssistantEvent -Message 'Will try connecting with user credentials' -Source $source
                Connect-MSCloudLoginLicensingWithUser
            }
            elseif ($Script:MSCloudLoginConnectionProfile.Licensing.AuthenticationType -eq 'ServicePrincipalWithThumbprint')
            {
                Add-MSCloudLoginAssistantEvent -Message "Attempting to connect to Licensing API using AAD App {$ApplicationID}" -Source $source
                Connect-MSCloudLoginLicensingWithCertificateThumbprint
            }
            else
            {
                throw 'Specified authentication method is not supported.'
            }

            $Script:MSCloudLoginConnectionProfile.Licensing.ConnectedDateTime = [System.DateTime]::Now.ToString()
            $Script:MSCloudLoginConnectionProfile.Licensing.Connected = $true
            $Script:MSCloudLoginConnectionProfile.Licensing.MultiFactorAuthentication = $false
            Add-MSCloudLoginAssistantEvent -Message "Successfully connected to Licensing API using AAD App {$ApplicationID}" -Source $source
        }
        catch
        {
            throw $_
        }
    }
}

function Connect-MSCloudLoginLicensingWithUser
{
    [CmdletBinding()]
    param()

    $source = 'Connect-MSCloudLoginLicensingWithUser'

    if ([System.String]::IsNullOrEmpty($Script:MSCloudLoginConnectionProfile.Licensing.TenantId))
    {
        $tenantId = $Script:MSCloudLoginConnectionProfile.Licensing.Credentials.UserName.Split('@')[1]
    }
    else
    {
        $tenantId = $Script:MSCloudLoginConnectionProfile.Licensing.TenantId
    }

    try
    {
        $managementToken = Get-AuthToken -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.Licensing.AuthorizationUrl `
            -Credentials $Script:MSCloudLoginConnectionProfile.Licensing.Credentials `
            -TenantId $tenantId `
            -ClientId $Script:MSCloudLoginConnectionProfile.Licensing.ApplicationId `
            -Resource $Script:MSCloudLoginConnectionProfile.Licensing.Resource

        $Script:MSCloudLoginConnectionProfile.Licensing.AccessToken = $managementToken.token_type.ToString() + ' ' + $managementToken.access_token.ToString()
        $Script:MSCloudLoginConnectionProfile.Licensing.Connected = $true
        $Script:MSCloudLoginConnectionProfile.Licensing.ConnectedDateTime = [System.DateTime]::Now.ToString()
    }
    catch
    {
        if ($_.ErrorDetails.Message -like '*AADSTS50076*')
        {
            Add-MSCloudLoginAssistantEvent -Message 'Account used required MFA' -Source $source
            Connect-MSCloudLoginLicensingWithUserMFA
        }
        else
        {
            $Script:MSCloudLoginConnectionProfile.Licensing.Connected = $false
            throw $_
        }
    }
}

function Connect-MSCloudLoginLicensingWithUserMFA
{
    [CmdletBinding()]
    param()

    if ([System.String]::IsNullOrEmpty($Script:MSCloudLoginConnectionProfile.Licensing.TenantId))
    {
        $tenantId = $Script:MSCloudLoginConnectionProfile.Licensing.Credentials.UserName.Split('@')[1]
    }
    else
    {
        $tenantId = $Script:MSCloudLoginConnectionProfile.Licensing.TenantId
    }

    $managementToken = Get-AuthToken -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.Licensing.AuthorizationUrl `
        -Credentials $Script:MSCloudLoginConnectionProfile.Licensing.Credentials `
        -TenantId $tenantId `
        -ClientId $Script:MSCloudLoginConnectionProfile.Licensing.ApplicationId `
        -Resource $Script:MSCloudLoginConnectionProfile.Licensing.Resource `
        -DeviceCode

    $Script:MSCloudLoginConnectionProfile.Licensing.AccessToken = $managementToken.token_type.ToString() + ' ' + $managementToken.access_token.ToString()
    $Script:MSCloudLoginConnectionProfile.Licensing.Connected = $true
    $Script:MSCloudLoginConnectionProfile.Licensing.MultiFactorAuthentication = $true
    $Script:MSCloudLoginConnectionProfile.Licensing.ConnectedDateTime = [System.DateTime]::Now.ToString()
}

function Connect-MSCloudLoginLicensingWithCertificateThumbprint
{
    [CmdletBinding()]
    param()

    $ProgressPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginLicensingWithCertificateThumbprint'

    Add-MSCloudLoginAssistantEvent -Message 'Attempting to connect to Licensing using CertificateThumbprint' -Source $source
    $tenantId = $Script:MSCloudLoginConnectionProfile.Licensing.TenantId

    try
    {
        $request = Get-AuthToken -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.Licensing.AuthorizationUrl `
            -CertificateThumbprint $Script:MSCloudLoginConnectionProfile.Licensing.CertificateThumbprint `
            -TenantId $tenantId `
            -ClientId $Script:MSCloudLoginConnectionProfile.Licensing.ApplicationId `
            -Resource $Script:MSCloudLoginConnectionProfile.Licensing.Resource

        $Script:MSCloudLoginConnectionProfile.Licensing.AccessToken = 'Bearer ' + $Request.access_token
        Add-MSCloudLoginAssistantEvent -Message 'Successfully connected to the Licensing API using Certificate Thumbprint' -Source $source

        $Script:MSCloudLoginConnectionProfile.Licensing.Connected = $true
        $Script:MSCloudLoginConnectionProfile.Licensing.ConnectedDateTime = [System.DateTime]::Now.ToString()
    }
    catch
    {
        throw $_
    }
}

function Disconnect-MSCloudLoginLicensing
{
    [CmdletBinding()]
    param()

    $source = 'Disconnect-MSCloudLoginLicensing'

    if ($Script:MSCloudLoginConnectionProfile.Licensing.Connected)
    {
        Add-MSCloudLoginAssistantEvent -Message 'Attempting to disconnect from Licensing API' -Source $source
        $Script:MSCloudLoginConnectionProfile.Licensing.Connected = $false
        Add-MSCloudLoginAssistantEvent -Message 'Successfully disconnected from Licensing API' -Source $source
    }
    else
    {
        Add-MSCloudLoginAssistantEvent -Message 'No connections to Licensing API were found' -Source $source
    }
}
