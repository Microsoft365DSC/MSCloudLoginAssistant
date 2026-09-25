function Connect-MSCloudLoginSecurityCompliance
{
    [CmdletBinding()]
    param()

    $InformationPreference = 'SilentlyContinue'
    $ProgressPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginSecurityCompliance'

    Add-MSCloudLoginAssistantEvent -Message "Connection Profile: $($Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter | Out-String)" -Source $source

    # Only a known token expiry ends the connection, certificate and credential sessions renew their tokens.
    if ($Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Connected)
    {
        $null = Test-MSCloudLoginConnectionReusable -WorkloadProfile $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter `
            -TokenBasedAuthTypes @() `
            -Source $source
    }

    if ($Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Connected)
    {
        if ($Script:MSCloudLoginCurrentLoadedModule -eq 'SC' -and
            $null -ne (Get-Command -Name 'Get-ComplianceSearch' -ErrorAction SilentlyContinue))
        {
            return
        }

        # Shared commands such as Get-Group must resolve to the Security & Compliance proxy module.
        if (Restore-MSCloudLoginProxyModule -ProbeCommand 'Get-ComplianceSearch' -Source $source)
        {
            $Script:MSCloudLoginCurrentLoadedModule = 'SC'
            return
        }

        Add-MSCloudLoginAssistantEvent -Message 'Security & Compliance proxy module is no longer loaded, reconnecting' -Source $source
        $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Connected = $false
    }

    $loadedModules = Get-Module
    Add-MSCloudLoginAssistantEvent -Message "The following modules are already loaded: $loadedModules" -Source $source

    Remove-MSCloudLoginProxyModule -ProbeCommand 'Get-ComplianceSearch' -Source $source

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
        $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.CompleteConnection($Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.MultiFactorAuthentication)
        $Script:MSCloudLoginCurrentLoadedModule = 'SC'
        Add-MSCloudLoginAssistantEvent -Message 'Reloaded the Security & Compliance Module' -Source $source
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
            Add-MSCloudLoginAssistantEvent -Message "Failed to connect to Security & Compliance with Certificate Thumbprint: $($_.Exception.Message)" -Source $source -EntryType 'Error'
            throw
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
            Add-MSCloudLoginAssistantEvent -Message "Failed to connect to Security & Compliance with Certificate Path: $($_.Exception.Message)" -Source $source -EntryType 'Error'
            throw
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
            if ((Test-MSCloudLoginMFARequiredError -ErrorRecord $_) -and -not (Assert-IsNonInteractiveShell))
            {
                Add-MSCloudLoginAssistantEvent -Message "Could not connect IPPSSession with Credentials & TenantId, account requires MFA: {$($_.Exception.Message)}" -Source $source
                Connect-MSCloudLoginSecurityComplianceMFA -TenantId $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.TenantId
            }
            else
            {
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Connected = $false
                Add-MSCloudLoginAssistantEvent -Message "Failed to connect to Security & Compliance with Credentials & TenantId: $($_.Exception.Message)" -Source $source -EntryType 'Error'
                throw
            }
        }
    }
    elseif ($Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.AuthenticationType -eq 'AccessTokens')
    {
        Add-MSCloudLoginAssistantEvent -Message 'Connecting to Security & Compliance with Access Token' -Source $source
        try
        {
            $accessToken = Get-MSCloudLoginAccessTokenValue -Token $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.AccessTokens[0]
            Connect-MSCloudLoginSecurityComplianceWithAccessToken -AccessToken $accessToken
        }
        catch
        {
            $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Connected = $false
            Add-MSCloudLoginAssistantEvent -Message "Failed to connect to Security & Compliance with Access Token: $($_.Exception.Message)" -Source $source -EntryType 'Error'
            throw
        }
    }
    elseif ($Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.AuthenticationType -eq 'Identity')
    {
        Add-MSCloudLoginAssistantEvent -Message 'Connecting to Security & Compliance with Managed Identity' -Source $source
        try
        {
            if ([System.String]::IsNullOrEmpty($Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ResourceUrl))
            {
                throw "No Security & Compliance resource URL is defined for environment '$($Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.EnvironmentName)'."
            }

            $accessToken = Get-AuthToken -Resource $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ResourceUrl -Identity
            Connect-MSCloudLoginSecurityComplianceWithAccessToken -AccessToken $accessToken
        }
        catch
        {
            $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Connected = $false
            Add-MSCloudLoginAssistantEvent -Message "Failed to connect to Security & Compliance with Managed Identity: $($_.Exception.Message)" -Source $source -EntryType 'Error'
            throw
        }
    }
    elseif ($Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.AuthenticationType -in @('Credentials', 'CredentialsWithApplicationId'))
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
            if ((Test-MSCloudLoginMFARequiredError -ErrorRecord $_) -and -not (Assert-IsNonInteractiveShell))
            {
                Add-MSCloudLoginAssistantEvent -Message "Could not connect IPPSSession with Credentials, account requires MFA: {$($_.Exception.Message)}" -Source $source
                Connect-MSCloudLoginSecurityComplianceMFA
            }
            else
            {
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Connected = $false
                Add-MSCloudLoginAssistantEvent -Message "Failed to connect to Security & Compliance with Credentials: $($_.Exception.Message)" -Source $source -EntryType 'Error'
                throw
            }
        }
    }
    else
    {
        throw "Authentication type '$($Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.AuthenticationType)' is not supported for workload 'SecurityComplianceCenter'."
    }

    $Script:MSCloudLoginCurrentLoadedModule = 'SC'
}

<#
.SYNOPSIS
    Connects to Security & Compliance with an access token.

.DESCRIPTION
    Uses the TenantId of the workload as the organization. The workload is marked as
    connected with the expiry of the token only after Connect-IPPSSession succeeds.

.PARAMETER AccessToken
    The access token for the Security & Compliance resource, with or without the 'Bearer ' prefix.
#>
function Connect-MSCloudLoginSecurityComplianceWithAccessToken
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [System.String]
        $AccessToken
    )

    $workloadProfile = $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter

    $tenantGuid = [System.Guid]::Empty
    if ([System.String]::IsNullOrEmpty($workloadProfile.TenantId) -or [System.Guid]::TryParse($workloadProfile.TenantId, [ref]$tenantGuid))
    {
        throw "TenantId must be the initial domain of the tenant (e.g. contoso.onmicrosoft.com) to connect to Security & Compliance with an access token."
    }

    $AccessToken = $AccessToken -replace '^Bearer\s+', ''
    $tokenExpiresOn = Get-MSCloudLoginAccessTokenExpiry -Token $AccessToken
    if ($null -ne $tokenExpiresOn -and $tokenExpiresOn -le [System.DateTime]::Now)
    {
        throw "The access token for Security & Compliance expired at {$tokenExpiresOn}. Provide a new access token."
    }

    Connect-IPPSSession -AccessToken $AccessToken `
        -Organization $workloadProfile.TenantId `
        -ConnectionUri $workloadProfile.ConnectionUrl `
        -AzureADAuthorizationEndpointUri $workloadProfile.AzureADAuthorizationEndpointUri `
        -EnableSearchOnlySession:$workloadProfile.EnableSearchOnlySession `
        -ShowBanner:$false `
        -ErrorAction Stop | Out-Null
    $workloadProfile.CompleteConnection($false, $tokenExpiresOn)
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
        Add-MSCloudLoginAssistantEvent -Message "Failed to connect to Security & Compliance using MFA: $($_.Exception.Message)" -Source $source -EntryType 'Error'
        throw
    }
}

function Disconnect-MSCloudLoginSecurityCompliance
{
    [CmdletBinding()]
    param()

    $source = 'Disconnect-MSCloudLoginSecurityCompliance'

    if ($Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Connected)
    {
        Add-MSCloudLoginAssistantEvent -Message 'Attempting to disconnect from Security & Compliance Center' -Source $source
        Disconnect-MSCloudLoginExchangeConnection -SecurityCompliance -Source $source
        $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Connected = $false
        Add-MSCloudLoginAssistantEvent -Message 'Successfully disconnected from Security & Compliance Center' -Source $source
    }
    else
    {
        Add-MSCloudLoginAssistantEvent -Message 'No connections to Security & Compliance Center were found.' -Source $source
    }
}
