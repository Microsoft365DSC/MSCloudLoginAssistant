function Connect-MSCloudLoginAzure
{
    [CmdletBinding()]
    param()

    $ProgressPreference = 'SilentlyContinue'
    $source = 'Connect-MSCloudLoginAzure'
    $workloadProfile = $Script:MSCloudLoginConnectionProfile.Azure
    if (Test-MSCloudLoginConnectionReusable -WorkloadProfile $workloadProfile `
            -ProbeScript $Script:MSCloudLoginConnectionProbes.Azure `
            -Source $source)
    {
        return
    }

    $additionalParameters = @{
        Scope = 'Process'
    }
    if ($workloadProfile.SubscriptionId)
    {
        $additionalParameters['Subscription'] = $workloadProfile.SubscriptionId
    }

    try
    {
        if ($workloadProfile.AuthenticationType -eq 'ServicePrincipalWithThumbprint')
        {
            Add-MSCloudLoginAssistantEvent -Message 'Connecting to Azure using AAD App with Certificate Thumbprint' -Source $source
            Connect-AzAccount -ServicePrincipal `
                -ApplicationId $workloadProfile.ApplicationId `
                -TenantId $workloadProfile.TenantId `
                -CertificateThumbprint $workloadProfile.CertificateThumbprint `
                -Environment $workloadProfile.EnvironmentName `
                @additionalParameters `
                -ErrorAction Stop | Out-Null
            $workloadProfile.CompleteConnection()
        }
        elseif ($workloadProfile.AuthenticationType -eq 'ServicePrincipalWithSecret')
        {
            Add-MSCloudLoginAssistantEvent -Message 'Connecting to Azure using AAD App with Client Secret' -Source $source
            $secStringPassword = $workloadProfile.ApplicationSecret | ConvertTo-SecureString -AsPlainText -Force
            $credential = [System.Management.Automation.PSCredential]::new($workloadProfile.ApplicationId, $secStringPassword)
            Connect-AzAccount -ServicePrincipal `
                -Credential $credential `
                -TenantId $workloadProfile.TenantId `
                -Environment $workloadProfile.EnvironmentName `
                @additionalParameters `
                -ErrorAction Stop | Out-Null
            $workloadProfile.CompleteConnection()
        }
        elseif ($workloadProfile.AuthenticationType -eq 'ServicePrincipalWithPath')
        {
            Add-MSCloudLoginAssistantEvent -Message 'Connecting to Azure using AAD App with Certificate Path' -Source $source
            Connect-AzAccount -ServicePrincipal `
                -ApplicationId $workloadProfile.ApplicationId `
                -TenantId $workloadProfile.TenantId `
                -CertificatePath $workloadProfile.CertificatePath `
                -CertificatePassword $workloadProfile.CertificatePassword `
                -Environment $workloadProfile.EnvironmentName `
                @additionalParameters `
                -ErrorAction Stop | Out-Null
            $workloadProfile.CompleteConnection()
        }
        elseif ($workloadProfile.AuthenticationType -eq 'CredentialsWithApplicationId' -or
            $workloadProfile.AuthenticationType -eq 'Credentials' -or
            $workloadProfile.AuthenticationType -eq 'CredentialsWithTenantId')
        {
            Add-MSCloudLoginAssistantEvent -Message 'Connecting to Azure using Credentials' -Source $source
            if ([System.String]::IsNullOrEmpty($workloadProfile.TenantId))
            {
                $workloadProfile.TenantId = Get-MSCloudLoginTenantDomainFromCredentials -Credentials $workloadProfile.Credentials
            }

            try
            {
                Connect-AzAccount -Credential $workloadProfile.Credentials `
                    -TenantId $workloadProfile.TenantId `
                    -Environment $workloadProfile.EnvironmentName `
                    @additionalParameters `
                    -ErrorAction Stop | Out-Null
                $workloadProfile.CompleteConnection()
            }
            catch
            {
                if (-not (Test-MSCloudLoginMFARequiredError -ErrorRecord $_) -or (Assert-IsNonInteractiveShell))
                {
                    throw
                }

                Add-MSCloudLoginAssistantEvent -Message 'MFA is required. Fallback to interactive login.' -Source $source -EntryType 'Warning'
                Connect-AzAccount -TenantId $workloadProfile.TenantId `
                    -Environment $workloadProfile.EnvironmentName `
                    @additionalParameters `
                    -ErrorAction Stop | Out-Null
                $workloadProfile.CompleteConnection($true)
            }
        }
        elseif ($workloadProfile.AuthenticationType -eq 'AccessTokens')
        {
            Add-MSCloudLoginAssistantEvent -Message 'Connecting to Azure using Access Token' -Source $source
            Connect-AzAccount -AccessToken $workloadProfile.AccessTokens[0] `
                -TenantId $workloadProfile.TenantId `
                -Environment $workloadProfile.EnvironmentName `
                -AccountId 'MSCloudLoginAssistant' `
                @additionalParameters `
                -ErrorAction Stop | Out-Null
            $workloadProfile.CompleteConnection()
        }
        elseif ($workloadProfile.AuthenticationType -eq 'Identity')
        {
            Add-MSCloudLoginAssistantEvent -Message 'Connecting to Azure using Managed Identity' -Source $source
            Connect-AzAccount -Identity `
                -Environment $workloadProfile.EnvironmentName `
                @additionalParameters `
                -ErrorAction Stop | Out-Null
            $workloadProfile.CompleteConnection()
        }
        else
        {
            throw 'Specified authentication method is not supported.'
        }

        $managementUrl = (Get-AzContext -ErrorAction Stop).Environment.ResourceManagerUrl
        Add-MSCloudLoginAssistantEvent -Message "Setting Azure Management URL to $managementUrl" -Source $source
        $workloadProfile.ManagementUrl = $managementUrl
    }
    catch
    {
        $workloadProfile.Connected = $false
        Add-MSCloudLoginAssistantEvent -Message "Failed to connect to Azure: $($_.Exception.Message)" -Source $source -EntryType 'Error'
        throw
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
        Disconnect-AzAccount -Scope Process | Out-Null
        $Script:MSCloudLoginConnectionProfile.Azure.Connected = $false
        Add-MSCloudLoginAssistantEvent -Message 'Successfully disconnected from Azure' -Source $source
    }
    else
    {
        Add-MSCloudLoginAssistantEvent -Message 'No connections to Azure were found' -Source $source
    }
}
