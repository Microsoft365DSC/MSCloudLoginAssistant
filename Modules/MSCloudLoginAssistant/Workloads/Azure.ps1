function Connect-MSCloudLoginAzure
{
    [CmdletBinding()]
    param()

    $ProgressPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginAzure'
    # If the current profile is not the same we expect, make the switch.
    if ($Script:MSCloudLoginConnectionProfile.Azure.Connected)
    {
        if (($Script:MSCloudLoginConnectionProfile.Azure.AuthenticationType -eq 'ServicePrincipalWithSecret' `
                    -or $Script:MSCloudLoginConnectionProfile.Azure.AuthenticationType -eq 'Identity') `
                -and (Get-Date -Date $Script:MSCloudLoginConnectionProfile.Azure.ConnectedDateTime) -lt [System.DateTime]::Now.AddMinutes(-50))
        {
            Add-MSCloudLoginAssistantEvent -Message 'Token is about to expire, renewing' -Source $source
            $Script:MSCloudLoginConnectionProfile.Azure.Connected = $false
        }
        elseif ($null -eq (Get-AzContext))
        {
            $Script:MSCloudLoginConnectionProfile.Azure.Connected = $false
        }
        else
        {
            return
        }
    }

    if ($Script:MSCloudLoginConnectionProfile.Azure.AuthenticationType -eq 'ServicePrincipalWithThumbprint')
    {
        Add-MSCloudLoginAssistantEvent -Message 'Connecting to Azure using AAD App with Certificate Thumbprint' -Source $source
        Connect-AzAccount -ServicePrincipal `
            -ApplicationId $Script:MSCloudLoginConnectionProfile.Azure.ApplicationId `
            -TenantId $Script:MSCloudLoginConnectionProfile.Azure.TenantId `
            -CertificateThumbprint $Script:MSCloudLoginConnectionProfile.Azure.CertificateThumbprint `
            -Environment $Script:MSCloudLoginConnectionProfile.Azure.EnvironmentName | Out-Null
        $Script:MSCloudLoginConnectionProfile.Azure.ConnectedDateTime = [System.DateTime]::Now.ToString()
        $Script:MSCloudLoginConnectionProfile.Azure.Connected = $true
        $Script:MSCloudLoginConnectionProfile.Azure.MultiFactorAuthentication = $false
    }
    elseif ($Script:MSCloudLoginConnectionProfile.Azure.AuthenticationType -eq 'ServicePrincipalWithSecret')
    {
        Add-MSCloudLoginAssistantEvent -Message 'Connecting to Azure using AAD App with Client Secret' -Source $source
        $secStringPassword = $Script:MSCloudLoginConnectionProfile.Azure.ApplicationSecret | ConvertTo-SecureString -AsPlainText -Force
        $credential = [System.Management.Automation.PSCredential]::new($Script:MSCloudLoginConnectionProfile.Azure.ApplicationId, $secStringPassword)
        Connect-AzAccount -ServicePrincipal `
            -Credential $credential `
            -TenantId $Script:MSCloudLoginConnectionProfile.Azure.TenantId `
            -Environment $Script:MSCloudLoginConnectionProfile.Azure.EnvironmentName | Out-Null
        $Script:MSCloudLoginConnectionProfile.Azure.ConnectedDateTime = [System.DateTime]::Now.ToString()
        $Script:MSCloudLoginConnectionProfile.Azure.Connected = $true
        $Script:MSCloudLoginConnectionProfile.Azure.MultiFactorAuthentication = $false
    }
    elseif ($Script:MSCloudLoginConnectionProfile.Azure.AuthenticationType -eq 'ServicePrincipalWithPath')
    {
        Add-MSCloudLoginAssistantEvent -Message 'Connecting to Azure using AAD App with Certificate Path' -Source $source
        Connect-AzAccount -ServicePrincipal `
            -ApplicationId $Script:MSCloudLoginConnectionProfile.Azure.ApplicationId `
            -TenantId $Script:MSCloudLoginConnectionProfile.Azure.TenantId `
            -CertificatePath $Script:MSCloudLoginConnectionProfile.Azure.CertificatePath `
            -CertificatePassword $Script:MSCloudLoginConnectionProfile.Azure.CertificatePassword `
            -Environment $Script:MSCloudLoginConnectionProfile.Azure.EnvironmentName | Out-Null
        $Script:MSCloudLoginConnectionProfile.Azure.ConnectedDateTime = [System.DateTime]::Now.ToString()
        $Script:MSCloudLoginConnectionProfile.Azure.Connected = $true
        $Script:MSCloudLoginConnectionProfile.Azure.MultiFactorAuthentication = $false
    }
    elseif ($Script:MSCloudLoginConnectionProfile.Azure.AuthenticationType -eq 'CredentialsWithApplicationId' -or
        $Script:MSCloudLoginConnectionProfile.Azure.AuthenticationType -eq 'Credentials' -or
        $Script:MSCloudLoginConnectionProfile.Azure.AuthenticationType -eq 'CredentialsWithTenantId')
    {
        Add-MSCloudLoginAssistantEvent -Message 'Connecting to Azure using Credentials' -Source $source
        try
        {
            if ([System.String]::IsNullOrEmpty($Script:MSCloudLoginConnectionProfile.Azure.TenantId))
            {
                $Script:MSCloudLoginConnectionProfile.Azure.TenantId = $Script:MSCloudLoginConnectionProfile.Azure.Credentials.UserName.Split('@')[1]
            }
            Connect-AzAccount -Credential $Script:MSCloudLoginConnectionProfile.Azure.Credentials `
                -TenantId $Script:MSCloudLoginConnectionProfile.Azure.TenantId `
                -Environment $Script:MSCloudLoginConnectionProfile.Azure.EnvironmentName `
                -ErrorAction Stop | Out-Null
            $Script:MSCloudLoginConnectionProfile.Azure.ConnectedDateTime = [System.DateTime]::Now.ToString()
            $Script:MSCloudLoginConnectionProfile.Azure.Connected = $true
            $Script:MSCloudLoginConnectionProfile.Azure.MultiFactorAuthentication = $false
        }
        catch
        {
            if ($_.Exception.Message -like '*AADSTS50076*')
            {
                Add-MSCloudLoginAssistantEvent -Message 'MFA is required. Fallback to interactive login.' -Source $source -EntryType 'Warning'
                Connect-AzAccount -TenantId $Script:MSCloudLoginConnectionProfile.Azure.TenantId `
                    -Environment $Script:MSCloudLoginConnectionProfile.Azure.EnvironmentName | Out-Null
                $Script:MSCloudLoginConnectionProfile.Azure.ConnectedDateTime = [System.DateTime]::Now.ToString()
                $Script:MSCloudLoginConnectionProfile.Azure.Connected = $true
                $Script:MSCloudLoginConnectionProfile.Azure.MultiFactorAuthentication = $true
            }
            else
            {
                throw $_
            }
        }
    }
    elseif ($Script:MSCloudLoginConnectionProfile.Azure.AuthenticationType -eq 'AccessTokens')
    {
        Add-MSCloudLoginAssistantEvent -Message 'Connecting to Azure using Access Token' -Source $source
        Connect-AzAccount -AccessToken $Script:MSCloudLoginConnectionProfile.Azure.AccessTokens[0]`
            -AccountId $Script:MSCloudLoginConnectionProfile.Azure.TenantId `
            -Environment $Script:MSCloudLoginConnectionProfile.Azure.EnvironmentName `
            -AccountId "MSCloudLoginAssistant" | Out-Null
        $Script:MSCloudLoginConnectionProfile.Azure.ConnectedDateTime = [System.DateTime]::Now.ToString()
        $Script:MSCloudLoginConnectionProfile.Azure.Connected = $true
        $Script:MSCloudLoginConnectionProfile.Azure.MultiFactorAuthentication = $false
    }
    elseif ($Script:MSCloudLoginConnectionProfile.Azure.AuthenticationType -eq 'Identity')
    {
        Add-MSCloudLoginAssistantEvent -Message 'Connecting to Azure using Managed Identity' -Source $source
        Connect-AzAccount -Identity `
            -Environment $Script:MSCloudLoginConnectionProfile.Azure.EnvironmentName | Out-Null
        $Script:MSCloudLoginConnectionProfile.Azure.ConnectedDateTime = [System.DateTime]::Now.ToString()
        $Script:MSCloudLoginConnectionProfile.Azure.Connected = $true
        $Script:MSCloudLoginConnectionProfile.Azure.MultiFactorAuthentication = $false
    }
    else
    {
        throw 'Specified authentication method is not supported.'
    }

    # If the connection to Azure was successful update the management URL
    if ($Script:MSCloudLoginConnectionProfile.Azure.Connected)
    {
        $managementUrl = (Get-AzContext).Environment.ResourceManagerUrl
        Add-MSCloudLoginAssistantEvent -Message "Setting Azure Management URL to $managementUrl" -Source $source
        $Script:MSCloudLoginConnectionProfile.Azure.ManagementUrl = $managementUrl
    }

    Add-MSCloudLoginAssistantEvent -Message 'Successfully connected to Azure' -Source $source
}

function Disconnect-MSCloudLoginAzure
{
    [CmdletBinding()]
    param()

    $source = 'Disconnect-MSCloudLoginAzure'

    if ($Script:MSCloudLoginConnectionProfile.Azure.Connected)
    {
        Add-MSCloudLoginAssistantEvent -Message 'Attempting to disconnect from Azure' -Source $source
        Disconnect-AzAccount | Out-Null
        $Script:MSCloudLoginConnectionProfile.Azure.Connected = $false
        Add-MSCloudLoginAssistantEvent -Message 'Successfully disconnected from Azure' -Source $source
    }
    else
    {
        Add-MSCloudLoginAssistantEvent -Message 'No connections to Azure were found' -Source $source
    }
}
