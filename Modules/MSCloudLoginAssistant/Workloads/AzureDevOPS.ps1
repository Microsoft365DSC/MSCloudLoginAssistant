function Connect-MSCloudLoginAzureDevOPS
{
    [CmdletBinding()]
    param()

    $InformationPreference = 'SilentlyContinue'
    $ProgressPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginAzureDevOPS'

    if ($Script:MSCloudLoginConnectionProfile.AzureDevOPS.AuthenticationType -eq 'ServicePrincipalWithThumbprint')
    {
        Add-MSCloudLoginAssistantEvent -Message "Attempting to connect to Azure DevOPS using AAD App {$ApplicationID}" -Source $source
        try
        {
            Connect-MSCloudLoginAzureDevOPSWithCertificateThumbprint
            $Script:MSCloudLoginConnectionProfile.AzureDevOPS.ConnectedDateTime = [System.DateTime]::Now.ToString()
            $Script:MSCloudLoginConnectionProfile.AzureDevOPS.Connected = $true
            $Script:MSCloudLoginConnectionProfile.AzureDevOPS.MultiFactorAuthentication = $false
            Add-MSCloudLoginAssistantEvent -Message "Successfully connected to Azure DevOPS using AAD App {$ApplicationID}" -Source $source
        }
        catch
        {
            throw $_
        }
    }
    elseif ($Script:MSCloudLoginConnectionProfile.AzureDevOPS.AuthenticationType -eq 'CredentialsWithApplicationId' -or
        $Script:MSCloudLoginConnectionProfile.AzureDevOPS.AuthenticationType -eq 'Credentials' -or
        $Script:MSCloudLoginConnectionProfile.AzureDevOPS.AuthenticationType -eq 'CredentialsWithTenantId')
    {
        Add-MSCloudLoginAssistantEvent -Message 'Attempting to connect to Azure DevOPS using Credentials.' -Source $source
        Connect-MSCloudAzureDevOPSWithUser
        Add-MSCloudLoginAssistantEvent -Message 'Successfully connected to Azure DevOPS using Credentials' -Source $source
    }
    else
    {
        throw 'Specified authentication method is not supported.'
    }
}
function Connect-MSCloudAzureDevOPSWithUser
{
    [CmdletBinding()]
    param()

    $source = 'Connect-MSCloudAzureDevOPSWithUser'

    if ([System.String]::IsNullOrEmpty($Script:MSCloudLoginConnectionProfile.AzureDevOPS.TenantId))
    {
        $tenantId = $Script:MSCloudLoginConnectionProfile.AzureDevOPS.Credentials.UserName.Split('@')[1]
    }
    else
    {
        $tenantId = $Script:MSCloudLoginConnectionProfile.AzureDevOPS.TenantId
    }

    try
    {
        $managementToken = Get-AuthToken -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.AzureDevOPS.AuthorizationUrl `
            -Credentials $Script:MSCloudLoginConnectionProfile.AzureDevOPS.Credentials `
            -TenantId $tenantId `
            -Resource $Script:MSCloudLoginConnectionProfile.AzureDevOPS.Resource `
            -ClientId $Script:MSCloudLoginConnectionProfile.AzureDevOPS.ApplicationId

        $Script:MSCloudLoginConnectionProfile.AzureDevOPS.AccessToken = $managementToken.token_type.ToString() + ' ' + $managementToken.access_token.ToString()
        $Script:MSCloudLoginConnectionProfile.AzureDevOPS.Connected = $true
        $Script:MSCloudLoginConnectionProfile.AzureDevOPS.ConnectedDateTime = [System.DateTime]::Now.ToString()
    }
    catch
    {
        if ($_.ErrorDetails.Message -like '*AADSTS50076*')
        {
            Add-MSCloudLoginAssistantEvent -Message 'Account used required MFA' -Source $source
            Connect-MSCloudLoginAzureDevOPSWithUserMFA
        }
        else
        {
            $Script:MSCloudLoginConnectionProfile.AzureDevOPS.Connected = $false
            throw $_
        }
    }
}

function Connect-MSCloudAzureDevOPSWithUserMFA
{
    [CmdletBinding()]
    param()

    if ([System.String]::IsNullOrEmpty($Script:MSCloudLoginConnectionProfile.AzureDevOPS.TenantId))
    {
        $tenantId = $Script:MSCloudLoginConnectionProfile.AzureDevOPS.Credentials.UserName.Split('@')[1]
    }
    else
    {
        $tenantId = $Script:MSCloudLoginConnectionProfile.AzureDevOPS.TenantId
    }

    $managementToken = Get-AuthToken -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.AzureDevOPS.AuthorizationUrl `
        -Credentials $Script:MSCloudLoginConnectionProfile.AzureDevOPS.Credentials `
        -TenantId $tenantId `
        -ClientId $Script:MSCloudLoginConnectionProfile.AzureDevOPS.ApplicationId `
        -Resource $Script:MSCloudLoginConnectionProfile.AzureDevOPS.Resource `
        -DeviceCode

    $Script:MSCloudLoginConnectionProfile.AzureDevOPS.AccessToken = $managementToken.token_type.ToString() + ' ' + $managementToken.access_token.ToString()
    $Script:MSCloudLoginConnectionProfile.AzureDevOPS.Connected = $true
    $Script:MSCloudLoginConnectionProfile.AzureDevOPS.MultiFactorAuthentication = $true
    $Script:MSCloudLoginConnectionProfile.AzureDevOPS.ConnectedDateTime = [System.DateTime]::Now.ToString()
}

function Connect-MSCloudLoginAzureDevOPSWithCertificateThumbprint
{
    [CmdletBinding()]
    param()

    $source = 'Connect-MSCloudLoginAzureDevOPSWithCertificateThumbprint'

    Add-MSCloudLoginAssistantEvent -Message 'Attempting to connect to Azure DevOPS using CertificateThumbprint' -Source $source
    $tenantId = $Script:MSCloudLoginConnectionProfile.AzureDevOPS.TenantId

    try
    {
        $request = Get-AuthToken -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.AzureDevOPS.AuthorizationUrl `
            -CertificateThumbprint $Script:MSCloudLoginConnectionProfile.AzureDevOPS.CertificateThumbprint `
            -TenantId $tenantId `
            -ClientId $Script:MSCloudLoginConnectionProfile.AzureDevOPS.ApplicationId `
            -Resource $Script:MSCloudLoginConnectionProfile.AzureDevOPS.Resource

        $Script:MSCloudLoginConnectionProfile.AzureDevOPS.AccessToken = 'Bearer ' + $Request.access_token
        Add-MSCloudLoginAssistantEvent -Message 'Successfully connected to the Azure DevOPS API using Certificate Thumbprint' -Source $source
    }
    catch
    {
        throw $_
    }
}
