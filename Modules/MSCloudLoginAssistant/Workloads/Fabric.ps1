function Connect-MSCloudLoginFabric
{
    [CmdletBinding()]
    param()

    $InformationPreference = 'SilentlyContinue'
    $ProgressPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginFabric'

    if ($Script:MSCloudLoginConnectionProfile.Fabric.AuthenticationType -eq 'ServicePrincipalWithThumbprint')
    {
        Add-MSCloudLoginAssistantEvent -Message "Attempting to connect to Fabric using AAD App {$ApplicationID}" -Source $source
        try
        {
            Connect-MSCloudLoginFabricWithCertificateThumbprint
            $Script:MSCloudLoginConnectionProfile.Fabric.ConnectedDateTime = [System.DateTime]::Now.ToString()
            $Script:MSCloudLoginConnectionProfile.Fabric.Connected = $true
            $Script:MSCloudLoginConnectionProfile.Fabric.MultiFactorAuthentication = $false
            Add-MSCloudLoginAssistantEvent -Message "Successfully connected to Fabric using AAD App {$ApplicationID}" -Source $source
        }
        catch
        {
            throw $_
        }
    }
    else
    {
        throw 'Specified authentication method is not supported.'
    }
}

function Connect-MSCloudLoginFabricWithCertificateThumbprint
{
    [CmdletBinding()]
    param()

    $ProgressPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginFabricWithCertificateThumbprint'

    try
    {
        Add-MSCloudLoginAssistantEvent -Message 'Attempting to connect to Fabric using CertificateThumbprint' -Source $source
        $request = Get-AuthToken -AuthorizationUrl $Script:MSCloudLoginConnectionProfile.Fabric.AuthorizationUrl `
            -CertificateThumbprint $Script:MSCloudLoginConnectionProfile.Fabric.CertificateThumbprint `
            -TenantId $Script:MSCloudLoginConnectionProfile.Fabric.TenantId `
            -ClientId $Script:MSCloudLoginConnectionProfile.Fabric.ApplicationId `
            -Scope $Script:MSCloudLoginConnectionProfile.Fabric.Scope

        $Script:MSCloudLoginConnectionProfile.Fabric.AccessToken = 'Bearer ' + $request.access_token
        Add-MSCloudLoginAssistantEvent -Message 'Successfully connected to the Fabric API using Certificate Thumbprint' -Source $source
    }
    catch
    {
        throw $_
    }
}
