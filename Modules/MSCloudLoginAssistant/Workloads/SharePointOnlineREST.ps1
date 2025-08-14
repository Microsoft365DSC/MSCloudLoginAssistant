function Connect-MSCloudLoginSharePointOnlineREST
{
    [CmdletBinding()]
    param()

    $InformationPreference = 'SilentlyContinue'
    $ProgressPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginSharePointOnlineREST'

    if (-not $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.AccessToken)
    {
        try
        {
            if ($Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.AuthenticationType -eq 'CredentialsWithApplicationId' -or
                $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.AuthenticationType -eq 'Credentials' -or
                $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.AuthenticationType -eq 'CredentialsWithTenantId')
            {
                Add-MSCloudLoginAssistantEvent -Message 'Will try connecting with user credentials' -Source $source
                Connect-MSCloudLoginSharePointOnlineRESTWithUser
            }
            elseif ($Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.AuthenticationType -eq 'ServicePrincipalWithThumbprint')
            {
                Add-MSCloudLoginAssistantEvent -Message "Attempting to connect to SharePoint Online REST using AAD App {$ApplicationID}" -Source $source
                Connect-MSCloudLoginSharePointOnlineRESTWithCertificateThumbprint
            }
            else
            {
                throw 'Specified authentication method is not supported.'
            }

            $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.ConnectedDateTime = [System.DateTime]::Now.ToString()
            $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.Connected = $true
            $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.MultiFactorAuthentication = $false
            Add-MSCloudLoginAssistantEvent -Message "Successfully connected to SharePoint Online REST using AAD App {$ApplicationID}" -Source $source
        }
        catch
        {
            throw $_
        }
    }
}

function Connect-MSCloudLoginSharePointOnlineRESTWithUser
{
    [CmdletBinding()]
    param()

    $source = 'Connect-MSCloudLoginSharePointOnlineRESTWithUser'
    if ([System.String]::IsNullOrEmpty($Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.TenantId))
    {
        $tenantId = $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.Credentials.UserName.Split('@')[1]
    }
    else
    {
        $tenantId = $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.TenantId
    }

    try
    {
        $managementToken = Get-AuthToken -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.AuthorizationUrl `
            -Credential $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.Credentials `
            -TenantId $tenantId `
            -ClientId $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.ApplicationId `
            -Scope $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.Scope `

        $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.AccessToken = $managementToken.token_type.ToString() + ' ' + $managementToken.access_token.ToString()
        $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.Connected = $true
        $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.ConnectedDateTime = [System.DateTime]::Now.ToString()
    }
    catch
    {
        if ($_.ErrorDetails.Message -like '*AADSTS50076*')
        {
            Add-MSCloudLoginAssistantEvent -Message 'Account used required MFA' -Source $source
            Connect-MSCloudLoginSharePointOnlineRESTWithUserMFA
        }
    }
}

function Connect-MSCloudLoginSharePointOnlineRESTWithUserMFA
{
    [CmdletBinding()]
    param()

    if ([System.String]::IsNullOrEmpty($Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.TenantId))
    {
        $tenantId = $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.Credentials.UserName.Split('@')[1]
    }
    else
    {
        $tenantId = $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.TenantId
    }

    $managementToken = Get-AuthToken -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.AuthorizationUrl `
        -Credentials $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.Credentials `
        -TenantId $tenantId `
        -ClientId $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.ApplicationId `
        -Scope $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.Scope `
        -DeviceCode

    $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.AccessToken = $managementToken.token_type.ToString() + ' ' + $managementToken.access_token.ToString()
    $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.Connected = $true
    $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.MultiFactorAuthentication = $true
    $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.ConnectedDateTime = [System.DateTime]::Now.ToString()
}

function Connect-MSCloudLoginSharePointOnlineRESTWithCertificateThumbprint
{
    [CmdletBinding()]
    param()

    $ProgressPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginSharePointOnlineRESTWithCertificateThumbprint'

    Add-MSCloudLoginAssistantEvent -Message 'Attempting to connect to SharePointOnlineREST using CertificateThumbprint' -Source $source

    try
    {
        $request = Get-AuthToken -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.AuthorizationUrl `
            -CertificateThumbprint $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.CertificateThumbprint `
            -TenantId $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.TenantId `
            -ClientId $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.ApplicationId `
            -Scope $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.Scope

        Add-MSCloudLoginAssistantEvent -Message 'Successfully connected to the SharePoint Online REST API using Certificate Thumbprint' -Source $source
        $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.AccessToken = 'Bearer ' + $Request.access_token
        $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.Connected = $true
        $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.ConnectedDateTime = [System.DateTime]::Now.ToString()
    }
    catch
    {
        throw $_
    }
}
