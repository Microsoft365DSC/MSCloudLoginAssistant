function Connect-MSCloudLoginTeams
{
    [CmdletBinding()]
    param()

    $ProgressPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginTeams'

    Add-MSCloudLoginAssistantEvent -Message 'Trying to get the Get-CsTeamsCallingPolicy command from within MSCloudLoginAssistant' -Source $source
    try
    {
        $results = Get-CsTeamsCallingPolicy -ErrorAction Stop
        if ($null -ne $results)
        {
            Add-MSCloudLoginAssistantEvent -Message 'Succeeded' -Source $source
            $Script:MSCloudLoginConnectionProfile.Teams.CompleteConnection($Script:MSCloudLoginConnectionProfile.Teams.MultiFactorAuthentication)
            return
        }
    }
    catch
    {
        # Liveness probe: a failure only means that there is no usable Teams session yet.
        Add-MSCloudLoginAssistantEvent -Message "Probe for existing Microsoft Teams session failed: $($_.Exception.Message)" -Source $source
        $Script:MSCloudLoginConnectionProfile.Teams.Connected = $false
    }

    if ($Script:MSCloudLoginConnectionProfile.Teams.Connected)
    {
        Add-MSCloudLoginAssistantEvent -Message 'Already connected to Microsoft Teams. Not attempting to re-connect.' -Source $source
        return
    }
    Add-MSCloudLoginAssistantEvent -Message 'No Active Connections to Microsoft Teams were found.' -Source $source

    if ($Script:MSCloudLoginConnectionProfile.Teams.AuthenticationType -eq 'ServicePrincipalWithThumbprint')
    {
        Add-MSCloudLoginAssistantEvent -Message "Connecting to Microsoft Teams using AzureAD Application {$($Script:MSCloudLoginConnectionProfile.Teams.ApplicationId)}" -Source $source
        if ($null -ne $Script:MSCloudLoginConnectionProfile.Teams.GraphScope -and `
            $null -ne $Script:MSCloudLoginConnectionProfile.Teams.TeamsScope -and `
            $null -ne $Script:MSCloudLoginConnectionProfile.Teams.TokenUrl -and `
            $null -eq $Script:CustomEnvConfig.CustomTeamsEndpoints)
        {
            $graphAccessToken = Get-MSCloudLoginAccessToken -ConnectionUri $Script:MSCloudLoginConnectionProfile.Teams.GraphScope `
                -AzureADAuthorizationEndpointUri $Script:MSCloudLoginConnectionProfile.Teams.TokenUrl `
                -ApplicationId $Script:MSCloudLoginConnectionProfile.Teams.ApplicationId `
                -TenantId $Script:MSCloudLoginConnectionProfile.Teams.TenantId `
                -CertificateThumbprint $Script:MSCloudLoginConnectionProfile.Teams.CertificateThumbprint
            $Script:MSCloudLoginConnectionProfile.Teams.AccessTokens += $graphAccessToken

            $teamsAccessToken = Get-MSCloudLoginAccessToken -ConnectionUri $Script:MSCloudLoginConnectionProfile.Teams.TeamsScope `
                -AzureADAuthorizationEndpointUri $Script:MSCloudLoginConnectionProfile.Teams.TokenUrl `
                -ApplicationId $Script:MSCloudLoginConnectionProfile.Teams.ApplicationId `
                -TenantId $Script:MSCloudLoginConnectionProfile.Teams.TenantId `
                -CertificateThumbprint $Script:MSCloudLoginConnectionProfile.Teams.CertificateThumbprint
            $Script:MSCloudLoginConnectionProfile.Teams.AccessTokens += $teamsAccessToken

            Connect-MicrosoftTeams -AccessTokens @($graphAccessToken, $teamsAccessToken)
            Add-MSCloudLoginAssistantEvent -Message 'Successfully connected to the Microsoft Graph API using Certificate Thumbprint' -Source $source
        }
        elseif ($null -ne $Script:CustomEnvConfig.CustomTeamsEndpoints -and $Script:CustomEnvConfig.CustomEnvironment)
        {
            if ($PSVersionTable.PSVersion.Major -gt 5)
            {
                throw 'Custom Environment connections to Microsoft Teams are only supported in PowerShell 5. Please run this module in PowerShell 5 to connect to Microsoft Teams in a custom environment.'
            }
            Set-TeamsEnvironmentConfig -EndpointUris $Script:CustomEnvConfig.CustomTeamsEndpoints
            Connect-MicrosoftTeams -ApplicationId $Script:MSCloudLoginConnectionProfile.Teams.ApplicationId `
                -TenantId $Script:MSCloudLoginConnectionProfile.Teams.TenantId `
                -CertificateThumbprint $Script:MSCloudLoginConnectionProfile.Teams.CertificateThumbprint
        }
        else
        {
            try
            {
                $ConnectionParams = @{
                    ApplicationId         = $Script:MSCloudLoginConnectionProfile.Teams.ApplicationId
                    TenantId              = $Script:MSCloudLoginConnectionProfile.Teams.TenantId
                    CertificateThumbprint = $Script:MSCloudLoginConnectionProfile.Teams.CertificateThumbprint
                }

                if ($Script:MSCloudLoginConnectionProfile.Teams.EnvironmentName -eq 'AzureUSGovernment')
                {
                    $ConnectionParams.Add('TeamsEnvironmentName', 'TeamsGCCH')
                }
                elseif ($Script:MSCloudLoginConnectionProfile.Teams.EnvironmentName -eq 'USGovernmentDoD' -or `
                        $Script:MSCloudLoginConnectionProfile.Teams.EnvironmentName -eq 'AzureDOD')
                {
                    $ConnectionParams.Add('TeamsEnvironmentName', 'TeamsDOD')
                }
                elseif ($Script:MSCloudLoginConnectionProfile.Teams.EnvironmentName -eq 'AzureChinaCloud')
                {
                    $ConnectionParams.Add('TeamsEnvironmentName', 'TeamsChina')
                }

                Connect-MicrosoftTeams @ConnectionParams | Out-Null
            }
            catch
            {
                $Script:MSCloudLoginConnectionProfile.Teams.Connected = $false
                Add-MSCloudLoginAssistantEvent -Message "Failed to connect to Microsoft Teams with Certificate Thumbprint: $($_.Exception.Message)" -Source $source -EntryType 'Error'
                throw
            }
        }

        $Script:MSCloudLoginConnectionProfile.Teams.CompleteConnection()
    }
    elseif ($Script:MSCloudLoginConnectionProfile.Teams.AuthenticationType -eq 'ServicePrincipalWithPath')
    {
        Add-MSCloudLoginAssistantEvent -Message "Connecting to Microsoft Teams using AzureAD Application {$($Script:MSCloudLoginConnectionProfile.Teams.ApplicationId)}" -Source $source
        $certificate = Get-MSCloudLoginCertificate -CertificatePath $Script:MSCloudLoginConnectionProfile.Teams.CertificatePath `
            -CertificatePassword $Script:MSCloudLoginConnectionProfile.Teams.CertificatePassword
        Connect-MicrosoftTeams -ApplicationId $Script:MSCloudLoginConnectionProfile.Teams.ApplicationId `
            -TenantId $Script:MSCloudLoginConnectionProfile.Teams.TenantId `
            -Certificate $certificate
        $Script:MSCloudLoginConnectionProfile.Teams.CompleteConnection()
    }
    elseif ($Script:MSCloudLoginConnectionProfile.Teams.AuthenticationType -eq 'Credentials' -or
        $Script:MSCloudLoginConnectionProfile.Teams.AuthenticationType -eq 'CredentialsWithTenantId')
    {
        if ($Script:MSCloudLoginConnectionProfile.Teams.EnvironmentName -eq 'AzureGermany')
        {
            $Script:MSCloudLoginConnectionProfile.Teams.Connected = $false
            throw 'Microsoft Teams is not supported in the Germany Cloud'
        }

        try
        {
            $ConnectionParams = @{
                Credential = $Script:MSCloudLoginConnectionProfile.Teams.Credentials
            }

            if ($Script:MSCloudLoginConnectionProfile.Teams.EnvironmentName -eq 'AzureUSGovernment')
            {
                $ConnectionParams.Add('TeamsEnvironmentName', 'TeamsGCCH')
            }

            if ($Script:MSCloudLoginConnectionProfile.Teams.EnvironmentName -eq 'USGovernmentDoD' -or `
                $Script:MSCloudLoginConnectionProfile.Teams.EnvironmentName -eq 'AzureDOD')
            {
                $ConnectionParams.Add('TeamsEnvironmentName', 'TeamsDOD')
            }

            if ($Script:MSCloudLoginConnectionProfile.Teams.EnvironmentName -eq 'AzureChinaCloud')
            {
                $ConnectionParams.Add('TeamsEnvironmentName', 'TeamsChina')
            }

            if (-not [System.String]::IsNullOrEmpty($Script:MSCloudLoginConnectionProfile.Teams.TenantId))
            {
                $ConnectionParams.Add('TenantId', $Script:MSCloudLoginConnectionProfile.Teams.TenantId)
            }

            Add-MSCloudLoginAssistantEvent -Message 'Connecting to Microsoft Teams using credentials.' -Source $source
            Add-MSCloudLoginAssistantEvent -Message "Params: $($ConnectionParams | Out-String)" -Source $source
            Add-MSCloudLoginAssistantEvent -Message "User: $($Script:MSCloudLoginConnectionProfile.Teams.Credentials.Username)" -Source $source
            Connect-MicrosoftTeams @ConnectionParams -ErrorAction Stop
            $Script:MSCloudLoginConnectionProfile.Teams.CompleteConnection()
        }
        catch
        {
            # TODO: 'One or more errors occurred.' is the generic AggregateException message and is
            # far too broad as an MFA indicator; kept for backwards compatibility.
            if ((Test-MSCloudLoginMFARequiredError -ErrorRecord $_ -AdditionalPatterns @('One or more errors occurred.')) -and -not (Assert-IsNonInteractiveShell))
            {
                Add-MSCloudLoginAssistantEvent -Message "Account requires MFA: $($_.Exception.Message)" -Source $source
                Connect-MSCloudLoginTeamsMFA
            }
            else
            {
                $Script:MSCloudLoginConnectionProfile.Teams.Connected = $false
                Add-MSCloudLoginAssistantEvent -Message "Failed to connect to Microsoft Teams with Credentials: $($_.Exception.Message)" -Source $source -EntryType 'Error'
                throw
            }
        }
    }
    elseif ($Script:MSCloudLoginConnectionProfile.Teams.AuthenticationType -eq 'Identity')
    {
        $ConnectionParams = @{
            Identity = $true
        }
        Add-MSCloudLoginAssistantEvent -Message 'Connecting to Microsoft Teams using Managed Identity' -Source $source
        Connect-MicrosoftTeams @ConnectionParams -ErrorAction Stop
        $Script:MSCloudLoginConnectionProfile.Teams.CompleteConnection()
    }
    elseif ($Script:MSCloudLoginConnectionProfile.Teams.AuthenticationType -eq 'AccessTokens')
    {
        $tokenValues = @()
        foreach ($tokenInfo in $Script:MSCloudLoginConnectionProfile.Teams.AccessTokens)
        {
            if ($null -ne $tokenInfo)
            {
                $tokenValues += Get-MSCloudLoginAccessTokenValue -Token $tokenInfo
            }
        }
        $ConnectionParams = @{
            AccessTokens = $tokenValues
        }
        Add-MSCloudLoginAssistantEvent -Message 'Connecting to Microsoft Teams using Access Token' -Source $source
        Connect-MicrosoftTeams @ConnectionParams -ErrorAction Stop
        $Script:MSCloudLoginConnectionProfile.Teams.CompleteConnection()
    }
    else
    {
        throw "Authentication type '$($Script:MSCloudLoginConnectionProfile.Teams.AuthenticationType)' is not supported for workload 'MicrosoftTeams'."
    }

    return
}

function Connect-MSCloudLoginTeamsMFA
{
    [CmdletBinding()]
    param()

    $source = 'Connect-MSCloudLoginTeamsMFA'

    try
    {
        $ConnectionParams = @{}
        if ($Script:MSCloudLoginConnectionProfile.Teams.EnvironmentName -eq 'AzureUSGovernment')
        {
            $ConnectionParams.Add('TeamsEnvironmentName', 'TeamsGCCH')
        }
        if ($Script:MSCloudLoginConnectionProfile.Teams.EnvironmentName -eq 'USGovernmentDoD' -or `
                $Script:MSCloudLoginConnectionProfile.Teams.EnvironmentName -eq 'AzureDOD')
        {
            $ConnectionParams.Add('TeamsEnvironmentName', 'TeamsDOD')
        }
        if (-not [System.String]::IsNullOrEmpty($Script:MSCloudLoginConnectionProfile.Teams.TenantId))
        {
            $ConnectionParams.Add('TenantId', $Script:MSCloudLoginConnectionProfile.Teams.TenantId)
        }
        Add-MSCloudLoginAssistantEvent -Message 'Disconnecting from Microsoft Teams' -Source $source
        Disconnect-MicrosoftTeams | Out-Null

        Add-MSCloudLoginAssistantEvent -Message 'Connecting to Microsoft Teams using MFA credentials' -Source $source
        Connect-MicrosoftTeams @ConnectionParams -ErrorAction Stop | Out-Null
        $Script:MSCloudLoginConnectionProfile.Teams.CompleteConnection($true)
    }
    catch
    {
        Add-MSCloudLoginAssistantEvent -Message "Error from MFA logic Path: $_" -Source $source -EntryType 'Error'
        $Script:MSCloudLoginConnectionProfile.Teams.Connected = $false
        throw
    }
}

function Disconnect-MSCloudLoginTeams
{
    [CmdletBinding()]
    param()

    $source = 'Disconnect-MSCloudLoginTeams'

    if ($Script:MSCloudLoginConnectionProfile.Teams.Connected)
    {
        Add-MSCloudLoginAssistantEvent -Message 'Attempting to disconnect from Microsoft Teams' -Source $source
        Disconnect-MicrosoftTeams | Out-Null
        $Script:MSCloudLoginConnectionProfile.Teams.Connected = $false
        Add-MSCloudLoginAssistantEvent -Message 'Successfully disconnected from Microsoft Teams' -Source $source
    }
    else
    {
        Add-MSCloudLoginAssistantEvent -Message 'No connections to Microsoft Teams were found.' -Source $source
    }
}
