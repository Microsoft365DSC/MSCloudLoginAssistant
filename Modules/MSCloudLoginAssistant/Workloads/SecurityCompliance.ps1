function Connect-MSCloudLoginSecurityCompliance
{
    [CmdletBinding()]
    param()

    $InformationPreference = 'SilentlyContinue'
    $ProgressPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginSecurityCompliance'

    Add-MSCloudLoginAssistantEvent -Message 'Trying to get the Get-ComplianceSearch command from within MSCloudLoginAssistant' -Source $source
    if ($Script:MSCloudLoginCurrentLoadedModule -eq 'SC')
    {
        try
        {
            Get-ComplianceSearch -ErrorAction Stop
            Add-MSCloudLoginAssistantEvent -Message 'Succeeded' -Source $source
            $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Connected = $true
            return
        }
        catch
        {
            Add-MSCloudLoginAssistantEvent -Message 'Failed' -Source $source
        }
    }

    Add-MSCloudLoginAssistantEvent -Message "Connection Profile: $($Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter | Out-String)" -Source $source
    if ($Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Connected)
    {
        return
    }

    $loadedModules = Get-Module
    Add-MSCloudLoginAssistantEvent -Message "The following modules are already loaded: $loadedModules" -Source $source

    $AlreadyLoadedSCProxyModules = $loadedModules | Where-Object -FilterScript { $_.ExportedCommands.Keys.Contains('Get-ComplianceSearch') }
    foreach ($loadedModule in $AlreadyLoadedSCProxyModules)
    {
        Add-MSCloudLoginAssistantEvent -Message "Removing module {$($loadedModule.Name)} from current S+C session" -Source $source
        # Temporarily set ErrorAction to SilentlyContinue to make sure the Remove-Module doesn't throw an error if some files are still in use.
        # Using the ErrorAction preference parameter doesn't work because within the Remove-Module cmdlet, that preference is not passed to
        # the underlying cmdlets.
        $currErrorActionPreference = $ErrorActionPreference
        $ErrorActionPreference = 'SilentlyContinue'
        Remove-Module $loadedModule.Name -Force -Verbose:$false | Out-Null
        $ErrorActionPreference = $currErrorActionPreference
    }

    [array]$activeSessions = Get-PSSession | Where-Object -FilterScript { $_.ComputerName -like '*ps.compliance.protection*' -and $_.State -eq 'Opened' }

    if ($activeSessions.Length -ge 1)
    {
        Add-MSCloudLoginAssistantEvent -Message "Found {$($activeSessions.Length)} existing Security and Compliance Session" -Source $source
        $ProxyModule = Import-PSSession $activeSessions[0] `
            -DisableNameChecking `
            -AllowClobber `
            -Verbose:$false
        Add-MSCloudLoginAssistantEvent -Message "Imported session into $ProxyModule" -Source $source
        Import-Module $ProxyModule -Global `
            -Verbose:$false | Out-Null
        $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Connected = $true
        Add-MSCloudLoginAssistantEvent 'Reloaded the Security & Compliance Module' -Source $source
        return
    }
    Add-MSCloudLoginAssistantEvent -Message 'No Active Connections to Security & Compliance were found.' -Source $source
    #endregion

    if ($Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.AuthenticationType -eq 'ServicePrincipalWithThumbprint')
    {
        Add-MSCloudLoginAssistantEvent -Message "Attempting to connect to Security and Compliance using AAD App {$($Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ApplicationID)}" -Source $source
        try
        {
            Add-MSCloudLoginAssistantEvent -Message 'Connecting to Security & Compliance with Service Principal and Certificate Thumbprint' -Source $source
            if ($null -ne $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ConnectionUrl -and `
                $null -ne $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.AzureADAuthorizationEndpointUri)
            {
                Add-MSCloudLoginAssistantEvent -Message 'Connecting by endpoints URI' -Source $source
                Connect-IPPSSession -AppId $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ApplicationId `
                    -Organization $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.TenantId `
                    -CertificateThumbprint $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.CertificateThumbprint `
                    -ShowBanner:$false `
                    -ConnectionUri $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ConnectionUrl `
                    -AzureADAuthorizationEndpointUri $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.AzureADAuthorizationEndpointUri `
                    -Verbose:$false
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ConnectedDateTime = [System.DateTime]::Now.ToString()
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.MultiFactorAuthentication = $false
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Connected = $true
            }
            else
            {
                Add-MSCloudLoginAssistantEvent -Message 'Connecting by environment name' -Source $source
                switch ($Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.EnvironmentName)
                {
                    { $_ -eq 'AzureUSGovernment' -or $_ -eq 'AzureDOD' }
                    {
                        Connect-IPPSSession -AppId $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ApplicationId `
                            -CertificateThumbprint $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.CertificateThumbprint `
                            -Organization $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.TenantId `
                            -ConnectionUri $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ConnectionUrl `
                            -AzureADAuthorizationEndpointUri $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.AzureADAuthorizationEndpointUri `
                            -ErrorAction Stop  `
                            -ShowBanner:$false | Out-Null
                        $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ConnectedDateTime = [System.DateTime]::Now.ToString()
                        $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.MultiFactorAuthentication = $false
                        $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Connected = $true
                    }
                    Default
                    {
                        Connect-IPPSSession -AppId $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ApplicationId `
                            -CertificateThumbprint $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.CertificateThumbprint `
                            -Organization $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.TenantId `
                            -ErrorAction Stop  `
                            -ShowBanner:$false | Out-Null
                        $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ConnectedDateTime = [System.DateTime]::Now.ToString()
                        $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.MultiFactorAuthentication = $false
                        $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Connected = $true
                    }
                }
            }
        }
        catch
        {
            $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Connected = $false
            throw $_
        }
    }
    elseif ($Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.AuthenticationType -eq 'ServicePrincipalWithPath')
    {
        try
        {
            Add-MSCloudLoginAssistantEvent -Message 'Connecting to Security & Compliance with Service Principal and Certificate Path' -Source $source
            switch ($Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.EnvironmentName)
            {
                { $_ -eq 'AzureUSGovernment' -or $_ -eq 'AzureDOD' }
                {
                    Connect-IPPSSession -AppId $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ApplicationId `
                        -CertificateFilePath $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.CertificatePath `
                        -Organization $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.TenantId `
                        -CertificatePassword $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.CertificatePassword `
                        -ConnectionUri $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ConnectionUrl `
                        -AzureADAuthorizationEndpointUri $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.AzureADAuthorizationEndpointUri  `
                        -ShowBanner:$false | Out-Null
                    $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ConnectedDateTime = [System.DateTime]::Now.ToString()
                    $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.MultiFactorAuthentication = $false
                    $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Connected = $true
                }
                Default
                {
                    Connect-IPPSSession -AppId $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ApplicationId `
                        -CertificateFilePath $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.CertificatePath `
                        -Organization $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.TenantId `
                        -CertificatePassword $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.CertificatePassword `
                        -ShowBanner:$false | Out-Null
                    $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ConnectedDateTime = [System.DateTime]::Now.ToString()
                    $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.MultiFactorAuthentication = $false
                    $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Connected = $true
                }
            }
        }
        catch
        {
            $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Connected = $false
            throw $_
        }
    }
    elseif ($Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.AuthenticationType -eq 'CredentialsWithTenantId')
    {
        try
        {
            $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.AuthorizationUrl = `
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.AuthorizationUrl.Replace('/organizations', "/$($Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.TenantId)")
            Add-MSCloudLoginAssistantEvent -Message 'Connecting to Security & Compliance with Credentials & TenantId' -Source $source
            Connect-IPPSSession -Credential $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Credentials `
                -ConnectionUri $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ConnectionUrl `
                -AzureADAuthorizationEndpointUri $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.AuthorizationUrl `
                -Verbose:$false -ErrorAction Stop  `
                -DelegatedOrganization $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.TenantId `
                -ShowBanner:$false | Out-Null
            $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ConnectedDateTime = [System.DateTime]::Now.ToString()
            $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.MultiFactorAuthentication = $false
            $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Connected = $true
        }
        catch
        {
            Add-MSCloudLoginAssistantEvent -Message "Could not connect connect IPPSSession with Credentials & TenantId: {$($_.Exception)}" -Source $source
            Connect-MSCloudLoginSecurityComplianceMFA -TenantId $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.TenantId
        }
    }
    elseif ($Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.AuthenticationType -eq 'AccessTokens')
    {
        Add-MSCloudLoginAssistantEvent -Message 'Connecting to Security & Compliance with Access Token' -Source $source
        Connect-M365Tenant -Workload 'ExchangeOnline' `
            -AccessTokens $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.AccessTokens `
            -TenantId $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.TenantId
        $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ConnectedDateTime = [System.DateTime]::Now.ToString()
        $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.MultiFactorAuthentication = $false
        $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Connected = $true
    }
    else
    {
        try
        {
            Add-MSCloudLoginAssistantEvent -Message 'Connecting to Security & Compliance with Credentials' -Source $source
            Connect-IPPSSession -Credential $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Credentials `
                -ConnectionUri $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ConnectionUrl `
                -AzureADAuthorizationEndpointUri $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.AuthorizationUrl `
                -Verbose:$false -ErrorAction Stop  `
                -ShowBanner:$false | Out-Null
            $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ConnectedDateTime = [System.DateTime]::Now.ToString()
            $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.MultiFactorAuthentication = $false
            $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Connected = $true
        }
        catch
        {
            Add-MSCloudLoginAssistantEvent -Message "Could not connect connect IPPSSession with Credentials: {$($_.Exception)}" -Source $source -EntryType Error
            Connect-MSCloudLoginSecurityComplianceMFA
        }
    }

    $Script:MSCloudLoginCurrentLoadedModule = 'SC'
}

function Connect-MSCloudLoginSecurityComplianceMFA
{
    [CmdletBinding()]
    param(
        [Parameter()]
        [System.String]
        $TenantId
    )

    $ProgressPreference = 'SilentlyContinue'
    $InformationPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginSecurityComplianceMFA'

    try
    {
        Add-MSCloudLoginAssistantEvent -Message 'Creating a new Security and Compliance Session using MFA' -Source $source
        if ($Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.EnvironmentName -eq 'AzureCloud')
        {
            if ([System.String]::IsNullOrEmpty($TenantId))
            {
                Connect-IPPSSession -UserPrincipalName $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Credentials.UserName `
                    -Verbose:$false  `
                    -ShowBanner:$false | Out-Null
            }
            else
            {
                Connect-IPPSSession -UserPrincipalName $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Credentials.UserName `
                    -Verbose:$false  `
                    -DelegatedOrganization $TenantId `
                    -ShowBanner:$false | Out-Null
            }
        }
        else
        {
            if ([System.String]::IsNullOrEmpty($TenantId))
            {
                Connect-IPPSSession -UserPrincipalName $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Credentials.UserName `
                    -ConnectionUri $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ConnectionUrl `
                    -Verbose:$false  `
                    -ShowBanner:$false | Out-Null
            }
            else
            {
                Connect-IPPSSession -UserPrincipalName $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Credentials.UserName `
                    -ConnectionUri $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ConnectionUrl `
                    -Verbose:$false `
                    -DelegatedOrganization $TenantId `
                    -ShowBanner:$false | Out-Null
            }
        }
        Add-MSCloudLoginAssistantEvent -Message 'New Session with MFA created successfully' -Source $source
        $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ConnectedDateTime = [System.DateTime]::Now.ToString()
        $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.MultiFactorAuthentication = $false
        $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Connected = $true
    }
    catch
    {
        $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Connected = $false
        throw $_
    }
}
