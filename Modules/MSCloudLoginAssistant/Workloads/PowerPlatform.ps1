function Connect-MSCloudLoginPowerPlatform
{
    [CmdletBinding()]
    param()

    $ProgressPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginPowerPlatform'

    if (Test-MSCloudLoginConnectionReusable -WorkloadProfile $Script:MSCloudLoginConnectionProfile.PowerPlatform -Source $source)
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
            $Script:MSCloudLoginConnectionProfile.PowerPlatform.CompleteConnection()
        }
        elseif ($Script:MSCloudLoginConnectionProfile.PowerPlatform.AuthenticationType -eq 'ServicePrincipalWithSecret')
        {
            Add-PowerAppsAccount -ApplicationId $Script:MSCloudLoginConnectionProfile.PowerPlatform.ApplicationId `
                -TenantID $Script:MSCloudLoginConnectionProfile.PowerPlatform.TenantId `
                -ClientSecret $Script:MSCloudLoginConnectionProfile.PowerPlatform.ApplicationSecret `
                -Endpoint $Script:MSCloudLoginConnectionProfile.PowerPlatform.Endpoint `
                -ErrorAction Stop | Out-Null
            $Script:MSCloudLoginConnectionProfile.PowerPlatform.CompleteConnection()
        }
        elseif ($Script:MSCloudLoginConnectionProfile.PowerPlatform.AuthenticationType -eq 'CredentialsWithTenantId')
        {
            throw 'You cannot specify TenantId with Credentials when connecting to PowerPlatforms.'
        }
        elseif ($Script:MSCloudLoginConnectionProfile.PowerPlatform.AuthenticationType -in @('Credentials', 'CredentialsWithApplicationId'))
        {
            Add-PowerAppsAccount -Username $Script:MSCloudLoginConnectionProfile.PowerPlatform.Credentials.UserName `
                -Password $Script:MSCloudLoginConnectionProfile.PowerPlatform.Credentials.Password `
                -Endpoint $Script:MSCloudLoginConnectionProfile.PowerPlatform.Endpoint `
                -ErrorAction Stop | Out-Null
            $Script:MSCloudLoginConnectionProfile.PowerPlatform.CompleteConnection()
        }
        else
        {
            throw "Authentication type '$($Script:MSCloudLoginConnectionProfile.PowerPlatform.AuthenticationType)' is not supported for workload 'PowerPlatform'."
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
                        -TenantID $Script:MSCloudLoginConnectionProfile.PowerPlatform.TenantId `
                        -CertificateThumbprint $Script:MSCloudLoginConnectionProfile.PowerPlatform.CertificateThumbprint `
                        -Endpoint 'preview' `
                        -ErrorAction Stop | Out-Null
                    $Script:MSCloudLoginConnectionProfile.PowerPlatform.CompleteConnection()
                }
                else
                {
                    Add-PowerAppsAccount -Username $Script:MSCloudLoginConnectionProfile.PowerPlatform.Credentials.UserName `
                        -Password $Script:MSCloudLoginConnectionProfile.PowerPlatform.Credentials.Password `
                        -Endpoint 'preview' `
                        -ErrorAction Stop | Out-Null

                    $Script:MSCloudLoginConnectionProfile.PowerPlatform.CompleteConnection()
                }
            }
            catch
            {
                if (Assert-IsNonInteractiveShell)
                {
                    $Script:MSCloudLoginConnectionProfile.PowerPlatform.Connected = $false
                    Add-MSCloudLoginAssistantEvent -Message "Failed to connect to PowerPlatform against the preview endpoint: $($_.Exception.Message)" -Source $source -EntryType 'Error'
                    throw
                }
                Connect-MSCloudLoginPowerPlatformMFA
            }
        }
        elseif ((Test-MSCloudLoginMFARequiredError -ErrorRecord $_ -AdditionalPatterns @('*Cannot find an overload for "UserCredential"*')) -and -not (Assert-IsNonInteractiveShell))
        {
            Add-MSCloudLoginAssistantEvent -Message "Account requires MFA: $($_.Exception.Message)" -Source $source
            Connect-MSCloudLoginPowerPlatformMFA
        }
        else
        {
            $Script:MSCloudLoginConnectionProfile.PowerPlatform.Connected = $false
            Add-MSCloudLoginAssistantEvent -Message "Failed to connect to PowerPlatform: $($_.Exception.Message)" -Source $source -EntryType 'Error'
            throw
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
        $Script:MSCloudLoginConnectionProfile.PowerPlatform.CompleteConnection($true)
    }
    catch
    {
        $Script:MSCloudLoginConnectionProfile.PowerPlatform.Connected = $false
        throw
    }
    return
}
