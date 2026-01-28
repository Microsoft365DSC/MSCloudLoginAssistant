function Connect-MSCloudLoginPowerPlatform
{
    [CmdletBinding()]
    param()

    $source = 'Connect-MSCloudLoginPowerPlatform'

    if ($Script:MSCloudLoginConnectionProfile.PowerPlatform.Connected)
    {
        return
    }

    try
    {
        if ($PSVersionTable.PSVersion.Major -ge 7)
        {
            Add-MSCloudLoginAssistantEvent -Message 'Using PowerShell 7 or above. Loading the Microsoft.PowerApps.Administration.PowerShell module using Windows PowerShell.' -Source $source
            Import-Module Microsoft.PowerApps.Administration.PowerShell -UseWindowsPowerShell -Global -DisableNameChecking | Out-Null
        }
        if ($Script:MSCloudLoginConnectionProfile.PowerPlatform.EnvironmentName -eq 'AzureGermany')
        {
            Write-Warning 'Microsoft PowerPlatform is not supported in the Germany Cloud'
            return
        }

        switch ($Script:CloudEnvironmentInfo.tenant_region_sub_scope)
        {
            'DODCON'
            {
                $Script:MSCloudLoginConnectionProfile.PowerPlatform.Endpoint = 'usgovhigh'
            }
            'DOD'
            {
                $Script:MSCloudLoginConnectionProfile.PowerPlatform.Endpoint = 'dod'
            }
            'GCC'
            {
                $Script:MSCloudLoginConnectionProfile.PowerPlatform.Endpoint = 'usgov'
            }
            default
            {
                $Script:MSCloudLoginConnectionProfile.PowerPlatform.Endpoint = 'prod'
            }
        }

        if ($Script:MSCloudLoginConnectionProfile.PowerPlatform.AuthenticationType -eq 'ServicePrincipalWithThumbprint')
        {
            Add-PowerAppsAccount -ApplicationId $Script:MSCloudLoginConnectionProfile.PowerPlatform.ApplicationId `
                -TenantID $Script:MSCloudLoginConnectionProfile.PowerPlatform.TenantId `
                -CertificateThumbprint $Script:MSCloudLoginConnectionProfile.PowerPlatform.CertificateThumbprint `
                -Endpoint $Script:MSCloudLoginConnectionProfile.PowerPlatform.Endpoint `
                -ErrorAction Stop | Out-Null
            $tokenValue = "Bearer $(($Global:currentSession.resourceTokens.'https://service.powerapps.com/'.accessToken).ToString())"
            $Script:MSCloudLoginConnectionProfile.PowerPlatform.AccessTokens = $tokenValue
            $Script:MSCloudLoginConnectionProfile.PowerPlatform.ConnectedDateTime = [System.DateTime]::Now.ToString()
            $Script:MSCloudLoginConnectionProfile.PowerPlatform.MultiFactorAuthentication = $false
            $Script:MSCloudLoginConnectionProfile.PowerPlatform.Connected = $true
        }
        elseif ($Script:MSCloudLoginConnectionProfile.PowerPlatform.AuthenticationType -eq 'ServicePrincipalWithSecret')
        {
            Add-PowerAppsAccount -ApplicationId $Script:MSCloudLoginConnectionProfile.PowerPlatform.ApplicationId `
                -TenantID $Script:MSCloudLoginConnectionProfile.PowerPlatform.TenantId `
                -ClientSecret $Script:MSCloudLoginConnectionProfile.PowerPlatform.ApplicationSecret `
                -Endpoint $Script:MSCloudLoginConnectionProfile.PowerPlatform.Endpoint `
                -ErrorAction Stop | Out-Null
            $Script:MSCloudLoginConnectionProfile.PowerPlatform.ConnectedDateTime = [System.DateTime]::Now.ToString()
            $Script:MSCloudLoginConnectionProfile.PowerPlatform.MultiFactorAuthentication = $false
            $Script:MSCloudLoginConnectionProfile.PowerPlatform.Connected = $true
        }
        elseif ($Script:MSCloudLoginConnectionProfile.PowerPlatform.AuthenticationType -eq 'CredentialsWithTenantId')
        {
            throw 'You cannot specify TenantId with Credentials when connecting to PowerPlatforms.'
        }
        else
        {
            Add-PowerAppsAccount -Username $Script:MSCloudLoginConnectionProfile.PowerPlatform.Credentials.UserName `
                -Password $Script:MSCloudLoginConnectionProfile.PowerPlatform.Credentials.Password `
                -Endpoint $Script:MSCloudLoginConnectionProfile.PowerPlatform.Endpoint `
                -ErrorAction Stop | Out-Null
            $Script:MSCloudLoginConnectionProfile.PowerPlatform.ConnectedDateTime = [System.DateTime]::Now.ToString()
            $Script:MSCloudLoginConnectionProfile.PowerPlatform.MultiFactorAuthentication = $false
            $Script:MSCloudLoginConnectionProfile.PowerPlatform.Connected = $true
        }
    }
    catch
    {
        if ($_.Exception.Message -like '*unknown_user_type: Unknown User Type*')
        {
            try
            {
                if ($Script:MSCloudLoginConnectionProfile.PowerPlatform.AuthenticationType -eq 'ServicePrincipalWithThumbprint')
                {
                    Add-PowerAppsAccount -ApplicationId $Script:MSCloudLoginConnectionProfile.PowerPlatform.ApplicationId `
                        -TenantID Global:MSCloudLoginConnectionProfile.PowerPlatform.$TenantId `
                        -CertificateThumbprint $Script:MSCloudLoginConnectionProfile.PowerPlatform.CertificateThumbprint `
                        -Endpoint 'preview' `
                        -ErrorAction Stop | Out-Null
                    $Script:MSCloudLoginConnectionProfile.PowerPlatform.ConnectedDateTime = [System.DateTime]::Now.ToString()
                    $Script:MSCloudLoginConnectionProfile.PowerPlatform.MultiFactorAuthentication = $false
                    $Script:MSCloudLoginConnectionProfile.PowerPlatform.Connected = $true
                }
                else
                {
                    Add-PowerAppsAccount -Username $Script:MSCloudLoginConnectionProfile.PowerPlatform.Credentials.UserName `
                        -Password $Script:MSCloudLoginConnectionProfile.PowerPlatform.Credentials.Password `
                        -Endpoint 'preview' `
                        -ErrorAction Stop | Out-Null

                    $Script:MSCloudLoginConnectionProfile.PowerPlatform.ConnectedDateTime = [System.DateTime]::Now.ToString()
                    $Script:MSCloudLoginConnectionProfile.PowerPlatform.MultiFactorAuthentication = $false
                    $Script:MSCloudLoginConnectionProfile.PowerPlatform.Connected = $true
                }
            }
            catch
            {
                Connect-MSCloudLoginPowerPlatformMFA
            }
        }
        elseif ($_.Exception.Message -like '*AADSTS50076: Due to a configuration change made by your administrator*')
        {
            Connect-MSCloudLoginPowerPlatformMFA
        }
        elseif ($_.Exception.Message -like '*Cannot find an overload for "UserCredential"*')
        {
            Connect-MSCloudLoginPowerPlatformMFA
        }
        else
        {
            $Script:MSCloudLoginConnectionProfile.PowerPlatform.Connected = $false
            throw $_
        }
    }
    return
}

function Connect-MSCloudLoginPowerPlatformMFA
{
    [CmdletBinding()]
    param()
    try
    {
        # Test-PowerAppsAccount This is failing in PowerApps admin module for GCCH MFA
        Add-PowerAppsAccount -Endpoint $Script:MSCloudLoginConnectionProfile.PowerPlatform.Endpoint
        $Script:MSCloudLoginConnectionProfile.PowerPlatform.ConnectedDateTime = [System.DateTime]::Now.ToString()
        $Script:MSCloudLoginConnectionProfile.PowerPlatform.MultiFactorAuthentication = $true
        $Script:MSCloudLoginConnectionProfile.PowerPlatform.Connected = $true
    }
    catch
    {
        $Script:MSCloudLoginConnectionProfile.PowerPlatform.Connected = $false
        throw $_
    }
    return
}
