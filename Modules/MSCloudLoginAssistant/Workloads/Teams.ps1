function Connect-MSCloudLoginTeams
{
    [CmdletBinding()]
    param()

    $ProgressPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginTeams'

    Add-MSCloudLoginAssistantEvent -Message 'Trying to get the Get-CsTeamsCallingPolicy command from within MSCloudLoginAssistant' -Source $source
    $currentErrorPreference = $ErrorActionPreference
    $Script:ErrorActionPreference = 'SilentlyContinue'
    try
    {
        if ($PSVersionTable.PSVersion.Major -ge 7)
        {
            Add-MSCloudLoginAssistantEvent -Message 'Using PowerShell 7 or above. Loading the MicrosoftTeams module using Windows PowerShell.' -Source $source
            Import-Module MicrosoftTeams -UseWindowsPowerShell -Global | Out-Null
        }

        $results = Get-CsTeamsCallingPolicy

        if ($null -ne $results)
        {
            Add-MSCloudLoginAssistantEvent -Message 'Succeeded' -Source $source
            $Script:MSCloudLoginConnectionProfile.Teams.Connected = $true
            return
        }
    }
    catch
    {
        Add-MSCloudLoginAssistantEvent -Message 'Failed' -Source $source -EntryType 'Error'
        $Script:MSCloudLoginConnectionProfile.Teams.Connected = $false
    }
    $Script:ErrorActionPreference = $currentErrorPreference

    if ($Script:MSCloudLoginConnectionProfile.Teams.Connected)
    {
        Add-MSCloudLoginAssistantEvent -Message 'Already connected to Microsoft Teams. Not attempting to re-connect.' -Source $source
        return
    }

    [array]$activeSessions = Get-PSSession | Where-Object -FilterScript { $_.Name -like '*SfBPowerShellSessionViaTeamsModule*' -and $_.State -eq 'Opened' }

    if ($activeSessions.Length -ge 1)
    {
        Add-MSCloudLoginAssistantEvent -Message "Found {$($activeSessions.Length)} existing Microsoft Teams Session" -Source $source
        Add-MSCloudLoginAssistantEvent -Message ($activeSessions | Out-String) -Source $source
        $ProxyModule = Import-PSSession $activeSessions[0] `
            -DisableNameChecking `
            -AllowClobber
        Add-MSCloudLoginAssistantEvent -Message "Imported session into $ProxyModule" -Source $source
        Import-Module $ProxyModule -Global | Out-Null
        $Script:MSCloudLoginConnectionProfile.Teams.Connected = $true
        Add-MSCloudLoginAssistantEvent 'Reloaded the Microsoft Teams Module' -Source $source
        return
    }
    Add-MSCloudLoginAssistantEvent -Message 'No Active Connections to Microsoft Teams were found.' -Source $source

    if ($Script:MSCloudLoginConnectionProfile.Teams.AuthenticationType -eq 'ServicePrincipalWithThumbprint')
    {
        Add-MSCloudLoginAssistantEvent -Message "Connecting to Microsoft Teams using AzureAD Application {$($Script:MSCloudLoginConnectionProfile.Teams.ApplicationId)}" -Source $source
        if ($null -ne $Script:MSCloudLoginConnectionProfile.Teams.Endpoints -and `
            $null -ne $Script:MSCloudLoginConnectionProfile.Teams.Endpoints.ConnectionUri -and `
            $null -ne $Script:MSCloudLoginConnectionProfile.Teams.Endpoints.AzureADAuthorizationEndpointUri)
        {
            $graphAccessToken = Get-MSCloudLoginAccessToken -ConnectionUri $Script:MSCloudLoginConnectionProfile.Teams.Endpoints.ConnectionUri `
                -AzureADAuthorizationEndpointUri $Script:MSCloudLoginConnectionProfile.Teams.Endpoints.AzureADAuthorizationEndpointUri `
                -ApplicationId $Script:MSCloudLoginConnectionProfile.Teams.ApplicationId `
                -TenantId $Script:MSCloudLoginConnectionProfile.Teams.TenantId `
                -CertificateThumbprint $Script:MSCloudLoginConnectionProfile.Teams.CertificateThumbprint
            $Script:MSCloudLoginConnectionProfile.Teams.AccessTokens += $graphAccessToken

            $teamsAccessToken = Get-MSCloudLoginAccessToken -ConnectionUri '48ac35b8-9aa8-4d74-927d-1f4a14a0b239/.default' `
                -AzureADAuthorizationEndpointUri $Script:MSCloudLoginConnectionProfile.Teams.Endpoints.AzureADAuthorizationEndpointUri `
                -ApplicationId $Script:MSCloudLoginConnectionProfile.Teams.ApplicationId `
                -TenantId $Script:MSCloudLoginConnectionProfile.Teams.TenantId `
                -CertificateThumbprint $Script:MSCloudLoginConnectionProfile.Teams.CertificateThumbprint
            $Script:MSCloudLoginConnectionProfile.Teams.AccessTokens += $teamsAccessToken

            Connect-MicrosoftTeams -AccessTokens @($graphAccessToken, $teamsAccessToken)
            Add-MSCloudLoginAssistantEvent -Message 'Successfully connected to the Microsoft Graph API using Certificate Thumbprint' -Source $source
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
                elseif ($Script:MSCloudLoginConnectionProfile.Teams.EnvironmentName -eq 'USGovernmentDoD')
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
                throw $_
            }
        }

        $Script:MSCloudLoginConnectionProfile.Teams.ConnectedDateTime = [System.DateTime]::Now.ToString()
        $Script:MSCloudLoginConnectionProfile.Teams.MultiFactorAuthentication = $false
        $Script:MSCloudLoginConnectionProfile.Teams.Connected = $true
    }
    elseif ($Script:MSCloudLoginConnectionProfile.Teams.AuthenticationType -eq 'Credentials' -or
        $Script:MSCloudLoginConnectionProfile.Teams.AuthenticationType -eq 'CredentialsWithTenantId')
    {
        if ($Script:MSCloudLoginConnectionProfile.Teams.EnvironmentName -eq 'AzureGermany')
        {
            Write-Warning 'Microsoft Teams is not supported in the Germany Cloud'
            $Script:MSCloudLoginConnectionProfile.Teams.Connected = $false
            return
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

            if ($Script:MSCloudLoginConnectionProfile.Teams.EnvironmentName -eq 'USGovernmentDoD')
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
            $Script:MSCloudLoginConnectionProfile.Teams.ConnectedDateTime = [System.DateTime]::Now.ToString()
            $Script:MSCloudLoginConnectionProfile.Teams.MultiFactorAuthentication = $false
            $Script:MSCloudLoginConnectionProfile.Teams.Connected = $true
        }
        catch
        {
            Add-MSCloudLoginAssistantEvent -Message "Error from Non-MFA Logic Path: $_" -Source $source -EntryType 'Error'
            if ($_.Exception -like '*AADSTS50076*' -or $_.Exception -eq 'One or more errors occurred.')
            {
                Connect-MSCloudLoginTeamsMFA
            }
            else
            {
                $Script:MSCloudLoginConnectionProfile.Teams.Connected = $false
                Add-MSCloudLoginAssistantEvent -Message $_ -Source $source -EntryType 'Error'
                throw $_
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
        $Script:MSCloudLoginConnectionProfile.Teams.ConnectedDateTime = [System.DateTime]::Now.ToString()
        $Script:MSCloudLoginConnectionProfile.Teams.MultiFactorAuthentication = $false
        $Script:MSCloudLoginConnectionProfile.Teams.Connected = $true
    }
    elseif ($Script:MSCloudLoginConnectionProfile.Teams.AuthenticationType -eq 'AccessToken')
    {
        $tokenValues = @()
        foreach ($tokenInfo in $Script:MSCloudLoginConnectionProfile.Teams.AccessTokens)
        {
            if ($null -ne $tokenInfo)
            {
                $Ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToCoTaskMemUnicode($tokenInfo)
                $AccessTokenValue = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($Ptr)
                [System.Runtime.InteropServices.Marshal]::ZeroFreeCoTaskMemUnicode($Ptr)
                $tokenValues += $AccessTokenValue
            }
        }
        $ConnectionParams = @{
            AccessTokens = $tokenValues
        }
        Add-MSCloudLoginAssistantEvent -Message 'Connecting to Microsoft Teams using Access Token' -Source $source
        Connect-MicrosoftTeams @ConnectionParams -ErrorAction Stop
        $Script:MSCloudLoginConnectionProfile.Teams.ConnectedDateTime = [System.DateTime]::Now.ToString()
        $Script:MSCloudLoginConnectionProfile.Teams.MultiFactorAuthentication = $false
        $Script:MSCloudLoginConnectionProfile.Teams.Connected = $true
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
        if ($Script:MSCloudLoginConnectionProfile.Teams.EnvironmentName -eq 'USGovernmentDoD')
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
        $Script:MSCloudLoginConnectionProfile.Teams.ConnectedDateTime = [System.DateTime]::Now.ToString()
        $Script:MSCloudLoginConnectionProfile.Teams.MultiFactorAuthentication = $true
        $Script:MSCloudLoginConnectionProfile.Teams.Connected = $true
    }
    catch
    {
        Add-MSCloudLoginAssistantEvent -Message "Error from MFA logic Path: $_" -Source $source -EntryType 'Error'
        $Script:MSCloudLoginConnectionProfile.Teams.Connected = $false
        throw $_
    }
}

function Disconnect-MSCloudLoginTeams
{
    [CmdletBinding()]
    param()

    $source = 'Disconnect-MSCloudLoginTeams'

    if ($Script:MSCloudLoginAssistant.Teams.Connected)
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
