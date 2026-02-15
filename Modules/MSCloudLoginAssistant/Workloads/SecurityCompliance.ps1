function Connect-MSCloudLoginSecurityCompliance
{
    [CmdletBinding()]
    param()

    $InformationPreference = 'SilentlyContinue'
    $ProgressPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginSecurityCompliance'

    Add-MSCloudLoginAssistantEvent -Message "Connection Profile: $($Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter | Out-String)" -Source $source
    if ($Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Connected)
    {
        return
    }

    $loadedModules = Get-Module
    Add-MSCloudLoginAssistantEvent -Message "The following modules are already loaded: $loadedModules" -Source $source

    $alreadyLoadedSCProxyModules = $loadedModules | Where-Object -FilterScript { $_.ExportedCommands.Keys.Contains('Get-ComplianceSearch') }
    foreach ($loadedModule in $alreadyLoadedSCProxyModules)
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
            Connect-IPPSSession -AppId $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ApplicationId `
                -Organization $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.TenantId `
                -CertificateThumbprint $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.CertificateThumbprint `
                -EnableSearchOnlySession:$Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.EnableSearchOnlySession `
                -ShowBanner:$false `
                -ConnectionUri $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ConnectionUrl `
                -AzureADAuthorizationEndpointUri $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.AzureADAuthorizationEndpointUri `
                -ErrorAction Stop | Out-Null
            $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.CompleteConnection()
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
            Connect-IPPSSession -AppId $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ApplicationId `
                -CertificateFilePath $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.CertificatePath `
                -Organization $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.TenantId `
                -EnableSearchOnlySession:$Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.EnableSearchOnlySession `
                -CertificatePassword $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.CertificatePassword `
                -ConnectionUri $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ConnectionUrl `
                -AzureADAuthorizationEndpointUri $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.AzureADAuthorizationEndpointUri  `
                -ShowBanner:$false `
                -ErrorAction Stop | Out-Null
            $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.CompleteConnection()
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
            $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.AzureADAuthorizationEndpointUri = `
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.AzureADAuthorizationEndpointUri.Replace('/organizations', "/$($Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.TenantId)")
            Add-MSCloudLoginAssistantEvent -Message 'Connecting to Security & Compliance with Credentials & TenantId' -Source $source
            Connect-IPPSSession -Credential $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Credentials `
                -ConnectionUri $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ConnectionUrl `
                -AzureADAuthorizationEndpointUri $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.AzureADAuthorizationEndpointUri `
                -DelegatedOrganization $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.TenantId `
                -EnableSearchOnlySession:$Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.EnableSearchOnlySession `
                -ShowBanner:$false `
                -ErrorAction Stop | Out-Null
            $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.CompleteConnection()
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
            -TenantId $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.TenantId `
            -ErrorAction Stop
        $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.CompleteConnection()
    }
    elseif ($Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.AuthenticationType -eq 'Identity')
    {
        Add-MSCloudLoginAssistantEvent -Message 'Connecting to Security & Compliance with Managed Identity' -Source $source
        Connect-IPPSSession -ManagedIdentity `
            -EnableSearchOnlySession:$Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.EnableSearchOnlySession `
            -ConnectionUri $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ConnectionUrl `
            -AzureADAuthorizationEndpointUri $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.AzureADAuthorizationEndpointUri `
            -ShowBanner:$false `
            -ErrorAction Stop | Out-Null
        $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.CompleteConnection()
    }
    else
    {
        try
        {
            Add-MSCloudLoginAssistantEvent -Message 'Connecting to Security & Compliance with Credentials' -Source $source
            Connect-IPPSSession -Credential $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Credentials `
                -ConnectionUri $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ConnectionUrl `
                -AzureADAuthorizationEndpointUri $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.AzureADAuthorizationEndpointUri `
                -EnableSearchOnlySession:$Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.EnableSearchOnlySession `
                -ShowBanner:$false `
                -ErrorAction Stop | Out-Null
            $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.CompleteConnection()
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
        if ([System.String]::IsNullOrEmpty($TenantId))
        {
            Connect-IPPSSession -UserPrincipalName $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Credentials.UserName `
                -ConnectionUri $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ConnectionUrl `
                -EnableSearchOnlySession:$Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.EnableSearchOnlySession `
                -ErrorAction Stop `
                -Verbose:$false  `
                -ShowBanner:$false | Out-Null
        }
        else
        {
            Connect-IPPSSession -UserPrincipalName $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Credentials.UserName `
                -ConnectionUri $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ConnectionUrl `
                -EnableSearchOnlySession:$Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.EnableSearchOnlySession `
                -ErrorAction Stop `
                -Verbose:$false `
                -DelegatedOrganization $TenantId `
                -ShowBanner:$false | Out-Null
        }
        Add-MSCloudLoginAssistantEvent -Message 'New Session with MFA created successfully' -Source $source
        $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.CompleteConnection($true)
    }
    catch
    {
        $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Connected = $false
        throw $_
    }
}
