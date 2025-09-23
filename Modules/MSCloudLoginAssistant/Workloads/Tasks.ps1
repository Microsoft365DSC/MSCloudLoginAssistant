function Connect-MSCloudLoginTasks
{
    [CmdletBinding()]
    param()

    $ProgressPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginTasks'

    if ($Script:MSCloudLoginConnectionProfile.Tasks.AuthenticationType -eq 'CredentialsWithApplicationId' -or
        $Script:MSCloudLoginConnectionProfile.Tasks.AuthenticationType -eq 'Credentials' -or
        $Script:MSCloudLoginConnectionProfile.Tasks.AuthenticationType -eq 'CredentialsWithTenantId')
    {
        Add-MSCloudLoginAssistantEvent -Message 'Will try connecting with user credentials' -Source $source
        Connect-MSCloudLoginTasksWithUser
    }
    elseif ($Script:MSCloudLoginConnectionProfile.Tasks.AuthenticationType -eq 'ServicePrincipalWithSecret')
    {
        Add-MSCloudLoginAssistantEvent -Message 'Will try connecting with Application Secret' -Source $source
        Connect-MSCloudLoginTasksWithAppSecret
    }
    elseif ($Script:MSCloudLoginConnectionProfile.Tasks.AuthenticationType -eq 'ServicePrincipalWithThumbprint')
    {
        Add-MSCloudLoginAssistantEvent -Message 'Will try connecting with Application Secret' -Source $source
        Connect-MSCloudLoginTasksWithCertificateThumbprint
    }
    elseif ($Script:MSCloudLoginConnectionProfile.Tasks.AuthenticationType -eq 'AccessToken')
    {
        Add-MSCloudLoginAssistantEvent -Message 'Will try connecting with Access Token' -Source $source
        $Ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToCoTaskMemUnicode($Script:MSCloudLoginConnectionProfile.Tasks.AccessTokens[0])
        $AccessTokenValue = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($Ptr)
        [System.Runtime.InteropServices.Marshal]::ZeroFreeCoTaskMemUnicode($Ptr)
        $Script:MSCloudLoginConnectionProfile.Tasks.AccessToken = $AccessTokenValue
    }
}

function Connect-MSCloudLoginTasksWithUser
{
    [CmdletBinding()]
    param()

    $source = 'Connect-MSCloudLoginTasksWithUser'

    if ([System.String]::IsNullOrEmpty($Script:MSCloudLoginConnectionProfile.Tasks.TenantId))
    {
        $tenantId = $Script:MSCloudLoginConnectionProfile.Tasks.Credentials.UserName.Split('@')[1]
    }
    else
    {
        $tenantId = $Script:MSCloudLoginConnectionProfile.Tasks.TenantId
    }

    try
    {
        $managementToken = Get-AuthToken -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.Tasks.AuthorizationUrl `
            -Credentials $Script:MSCloudLoginConnectionProfile.Tasks.Credentials `
            -TenantId $tenantId `
            -ClientId $Script:MSCloudLoginConnectionProfile.Tasks.ApplicationId `
            -Scope $Script:MSCloudLoginConnectionProfile.Tasks.Scope

        $Script:MSCloudLoginConnectionProfile.Tasks.AccessToken = $managementToken.token_type.ToString() + ' ' + $managementToken.access_token.ToString()
    }
    catch
    {
        if ($_.ErrorDetails.Message -like '*AADSTS50076*')
        {
            Add-MSCloudLoginAssistantEvent -Message 'Account used required MFA' -Source $source
            Connect-MSCloudLoginTasksWithUserMFA
        }
    }
}

function Connect-MSCloudLoginTasksWithUserMFA
{
    [CmdletBinding()]
    param()

    if ([System.String]::IsNullOrEmpty($Script:MSCloudLoginConnectionProfile.Tasks.TenantId))
    {
        $tenantId = $Script:MSCloudLoginConnectionProfile.Tasks.Credentials.UserName.Split('@')[1]
    }
    else
    {
        $tenantId = $Script:MSCloudLoginConnectionProfile.Tasks.TenantId
    }

    $managementToken = Get-AuthToken -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.Tasks.AuthorizationUrl `
        -Credentials $Script:MSCloudLoginConnectionProfile.Tasks.Credentials `
        -TenantId $tenantId `
        -ClientId $Script:MSCloudLoginConnectionProfile.Tasks.ApplicationId `
        -Scope $Script:MSCloudLoginConnectionProfile.Tasks.Scope

    $Script:MSCloudLoginConnectionProfile.Tasks.AccessToken = $managementToken.token_type.ToString() + ' ' + $managementToken.access_token.ToString()
}

function Connect-MSCloudLoginTasksWithAppSecret
{
    [CmdletBinding()]
    param()

    $managementToken = Get-AuthToken -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.Tasks.AuthorizationUrl `
        -ClientSecret $Script:MSCloudLoginConnectionProfile.Tasks.ApplicationSecret `
        -TenantId $tenantId `
        -ClientId $Script:MSCloudLoginConnectionProfile.Tasks.ApplicationId `
        -Scope $Script:MSCloudLoginConnectionProfile.Tasks.Scope

    $Script:MSCloudLoginConnectionProfile.Tasks.AccessToken = $managementToken.token_type.ToString() + ' ' + $managementToken.access_token.ToString()
}

function Connect-MSCloudLoginTasksWithCertificateThumbprint
{
    [CmdletBinding()]
    param()

    $ProgressPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginTasksWithCertificateThumbprint'

    Add-MSCloudLoginAssistantEvent -Message 'Attempting to connect to Whiteboard using CertificateThumbprint' -Source $source
    $tenantId = $Script:MSCloudLoginConnectionProfile.Tasks.TenantId

    try
    {
        $request = Get-AuthToken -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.Tasks.AuthorizationUrl `
            -CertificateThumbprint $Script:MSCloudLoginConnectionProfile.Tasks.CertificateThumbprint `
            -TenantId $tenantId `
            -ClientId $Script:MSCloudLoginConnectionProfile.Tasks.ApplicationId `
            -Scope $Script:MSCloudLoginConnectionProfile.Tasks.Scope

        $Script:MSCloudLoginConnectionProfile.Tasks.AccessToken = 'Bearer ' + $Request.access_token
        Add-MSCloudLoginAssistantEvent -Message 'Successfully connected to the Tasks API using Certificate Thumbprint' -Source $source
    }
    catch
    {
        throw $_
    }
}
