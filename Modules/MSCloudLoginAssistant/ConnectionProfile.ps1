class MSCloudLoginConnectionProfile
{
    [string]
    $CreatedTime

    [string]
    $OrganizationName

    [AdminAPI]
    $AdminAPI

    [Azure]
    $Azure

    [AzureDevOPS]
    $AzureDevOPS

    [DefenderForEndpoint]
    $DefenderForEndpoint

    [EngageHub]
    $EngageHub

    [ExchangeOnline]
    $ExchangeOnline

    [Fabric]
    $Fabric

    [Licensing]
    $Licensing

    [MicrosoftGraph]
    $MicrosoftGraph

    [PnP]
    $PnP

    [PowerPlatform]
    $PowerPlatform

    [PowerPlatformREST]
    $PowerPlatformREST

    [SecurityComplianceCenter]
    $SecurityComplianceCenter

    [SharePointOnlineREST]
    $SharePointOnlineREST

    [Tasks]
    $Tasks

    [Teams]
    $Teams

    MSCloudLoginConnectionProfile()
    {
        $this.CreatedTime = [System.DateTime]::Now.ToString()

        # Workloads Object Creation
        $this.AdminAPI                 = New-Object AdminAPI
        $this.Azure                    = New-Object Azure
        $this.AzureDevOPS              = New-Object AzureDevOPS
        $this.DefenderForEndpoint      = New-Object DefenderForEndpoint
        $this.EngageHub                = New-Object EngageHub
        $this.ExchangeOnline           = New-Object ExchangeOnline
        $this.Fabric                   = New-Object Fabric
        $this.Licensing                = New-Object Licensing
        $this.MicrosoftGraph           = New-Object MicrosoftGraph
        $this.PnP                      = New-Object PnP
        $this.PowerPlatform            = New-Object PowerPlatform
        $this.PowerPlatformREST        = New-Object PowerPlatformREST
        $this.SecurityComplianceCenter = New-Object SecurityComplianceCenter
        $this.SharePointOnlineREST     = New-Object SharePointOnlineREST
        $this.Tasks                    = New-Object Tasks
        $this.Teams                    = New-Object Teams
    }
}

class Workload : ICloneable
{
    [string]
    [ValidateSet('Credentials', 'CredentialsWithApplicationId', 'CredentialsWithTenantId', 'ServicePrincipalWithSecret', 'ServicePrincipalWithThumbprint', 'ServicePrincipalWithPath', 'Interactive', 'Identity', 'AccessTokens')]
    $AuthenticationType

    [boolean]
    $Connected = $false

    [string]
    $ConnectedDateTime

    [PSCredential]
    $Credentials

    [string]
    [ValidateSet('AzureCloud', 'AzureChinaCloud', 'AzureGermanyCloud', 'AzureUSGovernment', 'AzureDOD', 'Custom')]
    $EnvironmentName

    [boolean]
    $MultiFactorAuthentication

    [string]
    $ApplicationId

    [string]
    $ApplicationSecret

    [string]
    $TenantId

    [string]
    $TenantGUID

    [securestring]
    $CertificatePassword

    [string]
    $CertificatePath

    [string]
    $CertificateThumbprint

    [String[]]
    $AccessTokens

    [switch]
    $Identity

    [System.Collections.Hashtable]
    $Endpoints

    [object] Clone()
    {
        return $this.MemberwiseClone()
    }

    Setup()
    {
        $source = "Workload"
        Add-MSCloudLoginAssistantEvent -Message "Starting the Setup() logic" -Source $source
        Add-MSCloudLoginAssistantEvent -Message "`$this.EnvironmentName = '$($this.EnvironmentName)'" -Source $source
        Add-MSCloudLoginAssistantEvent -Message "`$Script:MSCloudLoginTriedGetEnvironment = '$($Script:MSCloudLoginTriedGetEnvironment)'" -Source $source
        # Determine the environment name based on email
        if ($null -eq $this.EnvironmentName -and -not $Script:MSCloudLoginTriedGetEnvironment)
        {
            $Script:MSCloudLoginTriedGetEnvironment = $true
            if ($null -ne $this.Credentials)
            {
                $Script:CloudEnvironmentInfo = Get-CloudEnvironmentInfo -Credentials $this.Credentials
            }
            elseif ($this.ApplicationID -and $this.CertificateThumbprint)
            {
                Add-MSCloudLoginAssistantEvent -Message "Trying to retrieve the Cloud Environment using Certificate Thumbprint." -Source $source
                $Script:CloudEnvironmentInfo = Get-CloudEnvironmentInfo -ApplicationId $this.ApplicationId -TenantId $this.TenantId -CertificateThumbprint $this.CertificateThumbprint
            }
            elseif ($this.ApplicationID -and $this.ApplicationSecret)
            {
                $Script:CloudEnvironmentInfo = Get-CloudEnvironmentInfo -ApplicationId $this.ApplicationId -TenantId $this.TenantId -ApplicationSecret $this.ApplicationSecret
            }
            elseif ($this.Identity.IsPresent)
            {
                $Script:CloudEnvironmentInfo = Get-CloudEnvironmentInfo -Identity -TenantId $this.TenantId
            }
            elseif ($this.AccessTokens)
            {
                $Script:CloudEnvironmentInfo = Get-CloudEnvironmentInfo -TenantId $this.TenantId
            }

            Add-MSCloudLoginAssistantEvent "Set environment to {$($Script:CloudEnvironmentInfo.tenant_region_sub_scope)}" -Source $source
        }
        #if ($null -eq $this.Endpoints)
        #{
        ##DLW
            switch ($Script:CloudEnvironmentInfo.tenant_region_sub_scope)
            {
                'AzureGermanyCloud'
                {
                    $this.EnvironmentName = 'O365GermanyCloud'
                }
                'DOD'
                {
                    $this.EnvironmentName = 'AzureDOD'
                }
                'DODCON'
                {
                    $this.EnvironmentName = 'AzureUSGovernment'
                }
                'USGov'
                {
                    $this.EnvironmentName = 'AzureUSGovernment'
                }
                default
                {
                    if ($null -ne $Script:CloudEnvironmentInfo -and $Script:CloudEnvironmentInfo.token_endpoint.StartsWith('https://login.partner.microsoftonline.cn'))
                    {
                        $this.EnvironmentName = 'AzureChinaCloud'

                        # Converting tenant to GUID. This is a limitation of the PnP module which
                        # can't recognize the tenant when FQDN is provided.
                        $tenantGUIDValue = $Script:CloudEnvironmentInfo.token_endpoint.Split('/')[3]
                        $this.TenantGUID = $tenantGUIDValue
                    }
                    elseif ($Global:CustomEnvironment)
                    {
                        $this.EnvironmentName = 'Custom'
                    }
                    else
                    {
                        $this.EnvironmentName = 'AzureCloud'
                    }
                }
            }
        #}
        #else
        #{
        #    $this.EnvironmentName = 'Custom'
        #}
        Add-MSCloudLoginAssistantEvent -Message "`$this.EnvironmentName was detected to be {$($this.EnvironmentName)}" -Source $source
        if ([System.String]::IsNullOrEmpty($this.EnvironmentName))
        {
            if ($null -ne $this.TenantId -and $this.TenantId.EndsWith('.cn'))
            {
                $this.EnvironmentName = 'AzureChinaCloud'
            }
            else
            {
                $this.EnvironmentName = 'AzureCloud'
            }
        }

        # Determine the Authentication Type
        if ($this.ApplicationId -and $this.TenantId -and $this.CertificateThumbprint)
        {
            $this.AuthenticationType = 'ServicePrincipalWithThumbprint'
        }
        elseif ($this.ApplicationId -and $this.TenantId -and $this.ApplicationSecret)
        {
            $this.AuthenticationType = 'ServicePrincipalWithSecret'
        }
        elseif ($this.ApplicationId -and $this.TenantId -and $this.CertificatePath -and $this.CertificatePassword)
        {
            $this.AuthenticationType = 'ServicePrincipalWithPath'
        }
        elseif ($this.Credentials -and $this.ApplicationId)
        {
            $this.AuthenticationType = 'CredentialsWithApplicationId'
        }
        elseif ($this.Credentials -and $this.TenantId)
        {
            $this.AuthenticationType = 'CredentialsWithTenantId'
        }
        elseif ($this.Credentials)
        {
            $this.AuthenticationType = 'Credentials'
        }
        elseif ($this.Identity)
        {
            $this.AuthenticationType = 'Identity'
        }
        elseif ($this.AccessTokens -and -not [System.String]::IsNullOrEmpty($this.TenantId))
        {
            $this.AuthenticationType = 'AccessTokens'
        }
        else
        {
            $this.AuthenticationType = 'Interactive'
        }
        Add-MSCloudLoginAssistantEvent -Message "`$this.AuthenticationType determined to be {$($this.AuthenticationType)}" -Source $source
    }
}

class AdminAPI:Workload
{
    [string]
    $AuthorizationUrl

    [string]
    $Scope

    [string]
    $AccessToken

    AdminAPI()
    {
    }

    [void] Connect()
    {
        ([Workload]$this).Setup()

        switch ($this.EnvironmentName)
        {
            'AzureDOD'
            {
                $this.Scope            = "6a8b4b39-c021-437c-b060-5a14a3fd65f3/.default"
                $this.AuthorizationUrl = "https://login.microsoftonline.us"
            }
            'AzureUSGovernment'
            {
                $this.Scope            = "6a8b4b39-c021-437c-b060-5a14a3fd65f3/.default"
                $this.AuthorizationUrl = "https://login.microsoftonline.us"
            }
            'Custom'
            {
                $this.Scope            = $Global:CustomAdminApiScope
                $this.AuthorizationUrl = $Global:CustomAdminApiAuthorizationUrl
            }
            default
            {
                $this.Scope            = "6a8b4b39-c021-437c-b060-5a14a3fd65f3/.default"
                $this.AuthorizationUrl = "https://login.microsoftonline.com"
            }
        }

        $Script:MSCloudLoginConnectionProfile.AdminAPI = $this
        Connect-MSCloudLoginAdminAPI
    }
}

class Azure:Workload
{
    Azure()
    {
    }

    [void] Connect()
    {
        $Script:MSCloudLoginTriedGetEnvironment = $false
        ([Workload]$this).Setup()

        $Script:MSCloudLoginConnectionProfile.Azure = $this
        Connect-MSCloudLoginAzure
    }

    [void] Disconnect()
    {
        Disconnect-MSCloudLoginAzure
    }
}

class AzureDevOPS:Workload
{
    [string]
    $HostUrl

    [string]
    $AuthorizationUrl

    [string]
    $Scope

    [string]
    $AccessToken

    AzureDevOPS()
    {
    }

    [void] Connect()
    {
        ([Workload]$this).Setup()
        switch ($this.EnvironmentName)
        {
            'AzureDOD'
            {
                $this.HostUrl          = "https://dev.azure.us"
                $this.Scope            = "499b84ac-1321-427f-aa17-267ca6975798/.default"
                $this.AuthorizationUrl = "https://login.microsoftonline.us"
            }
            'AzureUSGovernment'
            {
                $this.HostUrl          = "https://dev.azure.com"
                $this.Scope            = "499b84ac-1321-427f-aa17-267ca6975798/.default"
                $this.AuthorizationUrl = "https://login.microsoftonline.us"
            }
            'Custom'
            {
                $this.HostUrl          = $Global:CustomAzureDevopsHostUrl
                $this.Scope            = $Global:CustomAzureDevopsScope
                $this.AuthorizationUrl = $Global:CustomAzureDevopsAuthorizationUrl
            }
            default
            {
                $this.HostUrl          = "https://dev.azure.com"
                $this.Scope            = "499b84ac-1321-427f-aa17-267ca6975798/.default"
                $this.AuthorizationUrl = "https://login.microsoftonline.com"
            }
        }


        $Script:MSCloudLoginConnectionProfile.AzureDevOPS = $this
        Connect-MSCloudLoginAzureDevOPS
    }
}

class DefenderForEndpoint:Workload
{
    [string]
    $HostUrl

    [string]
    $AuthorizationUrl

    [string]
    $Scope

    [string]
    $AccessToken

    DefenderForEndpoint()
    {
    }

    [void] Connect()
    {
        ([Workload]$this).Setup()

        switch ($this.EnvironmentName)
        {
            'AzureDOD'
            {
                $this.HostUrl          = 'https://api-gov.securitycenter.microsoft.us'
                $this.Scope            = 'https://api.securitycenter.microsoft.com/.default'
                $this.AuthorizationUrl = 'https://login.microsoftonline.us'
            }
            'AzureUSGovernment'
            {
                $this.HostUrl          = 'https://api-gcc.securitycenter.microsoft.us'
                $this.Scope            = 'https://api.securitycenter.microsoft.com/.default'
                $this.AuthorizationUrl = 'https://login.microsoftonline.com'
            }
            'Custom'
            {
                $this.HostUrl          = $Global:CustomDefenderForEndpointHostUrl
                $this.Scope            = $Global:CustomDefenderForEndpointScope
                $this.AuthorizationUrl = $Global:CustomDefenderForEndpointAuthorizationUrl
            }
            default
            {
                $this.HostUrl          = 'https://api.security.microsoft.com'
                $this.Scope            = 'https://api.securitycenter.microsoft.com/.default'
                $this.AuthorizationUrl = 'https://login.microsoftonline.com'
            }
        }

        $Script:MSCloudLoginConnectionProfile.DefenderForEndpoint = $this
        Connect-MSCloudLoginDefenderForEndpoint
    }

}

class EngageHub:Workload
{
    [string]
    $AuthorizationUrl

    [string]
    $Scope

    [string]
    $AccessToken

    [string]
    $APIUrl

    EngageHub()
    {
    }

    [void] Connect()
    {
        ([Workload]$this).Setup()

        switch ($this.EnvironmentName)
        {
            'AzureDOD'
            {
                $this.Scope            = "https://engagehub.microsoft.us/.default"
                $this.AuthorizationUrl = "https://login.microsoftonline.us"
                $this.APIUrl           = "https://api.dev.engagecenter.microsoft.us"

            }
            'AzureUSGovernment'
            {
                $this.Scope            = "https://engagehub.microsoft.us/.default"
                $this.AuthorizationUrl = "https://login.microsoftonline.us"
                $this.APIUrl           = "https://api.dev.engagecenter.microsoft.us"
            }
            'Custom'
            {
                $this.Scope            = $Global:CustomEngageHubScope
                $this.AuthorizationUrl = $Global:CustomEngageHubAuthorizationUrl
                $this.APIUrl           = $Global:CustomEngageHubAPIUrl
            }
            default
            {
                $this.Scope            = "https://engagehub.microsoft.com/.default"
                $this.AuthorizationUrl = "https://login.microsoftonline.com"
                $this.APIUrl           = "https://api.dev.engagecenter.microsoft.com"
            }
        }
        $Script:MSCloudLoginConnectionProfile.EngageHub = $this
        Connect-MSCloudLoginEngageHub
    }
}

class ExchangeOnline:Workload
{
    [string]
    [ValidateSet('O365Default', 'O365GermanyCloud', 'O365China', 'O365USGovGCCHigh', 'O365USGovDod')]
    $ExchangeEnvironmentName = 'O365Default'

    [string]
    $ConnectionUri

    [string]
    $AzureADAuthorizationEndpointUri

    [boolean]
    $SkipModuleReload = $false

    [System.String[]]
    $CmdletsToLoad = @()

    [System.String[]]
    $LoadedCmdlets = @()

    [boolean]
    $LoadedAllCmdlets = $false

    ExchangeOnline()
    {
    }

    [void] Connect()
    {
        ([Workload]$this).Setup()

        switch ($this.EnvironmentName)
        {
            'AzureCloud'
            {
                $this.ExchangeEnvironmentName = 'O365Default'
            }
            'AzureGermanyCloud'
            {
                $this.ExchangeEnvironmentName = 'O365GermanyCloud'
            }
            'AzureDOD'
            {
                $this.ExchangeEnvironmentName = 'O365USGovDoD'
            }
            'AzureUSGovernment'
            {
                $this.ExchangeEnvironmentName = 'O365USGovGCCHigh'
            }
            'AzureChinaCloud'
            {
                $this.ExchangeEnvironmentName = 'O365China'
            }
            'Custom'
            {
                $this.ConnectionUri                   = $Global:CustomEXOConnectionUri
                $this.AzureADAuthorizationEndpointUri = $Global:CustomEXOAzureADAuthorizationEndpointUri
            }
        }
        $Script:MSCloudLoginConnectionProfile.ExchangeOnline = $this
        Connect-MSCloudLoginExchangeOnline -Verbose
    }

    [void] Disconnect()
    {
        $source = 'ExchangeOnline-Disconnect()'
        Add-MSCloudLoginAssistantEvent -Message 'Disconnecting from Exchange Online Connection' -Source $source
        Disconnect-ExchangeOnline -Confirm:$false
        $this.Connected = $false
        $this.LoadedAllCmdlets = $false
        $this.LoadedCmdlets = @()
        $this.CmdletsToLoad = @()
    }
}

class Fabric:Workload
{
    [string]
    $HostUrl

    [string]
    $AuthorizationUrl

    [string]
    $Scope

    [string]
    $AccessToken

    Fabric()
    {
    }

    [void] Connect()
    {
        ([Workload]$this).Setup()
        switch ($this.EnvironmentName)
        {
            'AzureDOD'
            {
                $this.HostUrl          = "https://api.fabric.microsoft.us"
                $this.Scope            = "https://api.fabric.microsoft.us/.default"
                $this.AuthorizationUrl = "https://login.microsoftonline.us"
            }
            'AzureUSGovernment'
            {
                $this.HostUrl          = "https://api.fabric.microsoft.us"
                $this.Scope            = "https://api.fabric.microsoft.us/.default"
                $this.AuthorizationUrl = "https://login.microsoftonline.us"
            }
            'Custom'
            {
                $this.HostUrl          = $Global:CustomFabricHostUrl
                $this.Scope            = $Global:CustomFabricScope
                $this.AuthorizationUrl = $Global:CustomFabricAuthorizationUrl
            }
            default
            {
                $this.HostUrl          = "https://api.fabric.microsoft.com"
                $this.Scope            = "https://api.fabric.microsoft.com/.default"
                $this.AuthorizationUrl = "https://login.microsoftonline.com"
            }
        }

        $Script:MSCloudLoginConnectionProfile.Fabric = $this
        Connect-MSCloudLoginFabric
    }
}

class Licensing:Workload
{
    [string]
    $HostUrl

    [string]
    $AuthorizationUrl

    [string]
    $Scope

    [string]
    $AccessToken

    Licensing()
    {
    }

    [void] Connect()
    {
        ([Workload]$this).Setup()
        switch ($this.EnvironmentName)
        {
            'AzureDOD'
            {
                $this.HostUrl          = "https://licensing.m365.microsoft.com"
                $this.Scope            = "aeb86249-8ea3-49e2-900b-54cc8e308f85/.default"
                $this.AuthorizationUrl = "hhttps://login.microsoftonline.com"
            }
            'AzureUSGovernment'
            {
                $this.HostUrl          = "https://licensing.m365.microsoft.com"
                $this.Scope            = "aeb86249-8ea3-49e2-900b-54cc8e308f85/.default"
                $this.AuthorizationUrl = "hhttps://login.microsoftonline.com"
            }
            'Custom'
            {
                $this.HostUrl          = $Global:CustomLicensingHostUrl
                $this.Scope            = $Global:CustomLicensingScope
                $this.AuthorizationUrl = $Global:CustomLicensingAuthorizationUrl
            }
            default
            {
                $this.HostUrl          = "https://licensing.m365.microsoft.com"
                $this.Scope            = "aeb86249-8ea3-49e2-900b-54cc8e308f85/.default"
                $this.AuthorizationUrl = "https://login.microsoftonline.com"
            }
        }

        $Script:MSCloudLoginConnectionProfile.Licensing = $this
        Connect-MSCloudLoginLicensing
    }
}

class MicrosoftGraph:Workload
{
    [string]
    [ValidateSet('China', 'Global', 'USGov', 'USGovDoD', 'Germany', 'Custom')]
    $GraphEnvironment = 'Global'

    [string]
    [ValidateSet('v1.0', 'beta')]
    $ProfileName = 'v1.0'

    [string]
    $ResourceUrl

    [string]
    $Scope

    [string]
    $TokenUrl

    MicrosoftGraph()
    {
    }

    [void] Connect()
    {
        ([Workload]$this).Setup()

        if ($null -ne $this.Credentials -and [System.String]::IsNullOrEmpty($this.TenantId))
        {
            $this.TenantId = $this.Credentials.Username.Split('@')[1]
        }

        switch ($this.EnvironmentName)
        {
            'AzureCloud'
            {
                $this.GraphEnvironment = 'Global'
                $this.ResourceUrl      = 'https://graph.microsoft.com/'
                $this.Scope            = 'https://graph.microsoft.com/.default'
                $this.TokenUrl         = "https://login.microsoftonline.com/$($this.TenantId)/oauth2/v2.0/token"
            }
            'AzureUSGovernment'
            {
                $this.GraphEnvironment = 'USGov'
                $this.ResourceUrl      = 'https://graph.microsoft.us/'
                $this.Scope            = 'https://graph.microsoft.us/.default'
                $this.TokenUrl         = "https://login.microsoftonline.us/$($this.TenantId)/oauth2/v2.0/token"
            }
            'AzureDOD'
            {
                $this.GraphEnvironment = 'USGovDoD'
                $this.ResourceUrl      = 'https://dod-graph.microsoft.us/'
                $this.Scope            = 'https://dod-graph.microsoft.us/.default'
                $this.TokenUrl         = "https://login.microsoftonline.us/$($this.TenantId)/oauth2/v2.0/token"
            }
            'AzureChinaCloud'
            {
                $this.GraphEnvironment = 'China'
                $this.ResourceUrl      = 'https://microsoftgraph.chinacloudapi.cn/'
                $this.Scope            = 'https://microsoftgraph.chinacloudapi.cn/.default'
                $this.TokenUrl         = "https://login.chinacloudapi.cn/$($this.TenantId)/oauth2/v2.0/token"
            }
            'Custom'
            {
                $this.GraphEnvironment = 'Custom'
                $this.ResourceUrl      = $Global:CustomGraphResourceUrl
                $this.Scope            = $Global:CustomGraphScope
                $this.TokenUrl         = "$($Global:CustomGraphTokenUrl)/$($this.TenantId)/oauth2/v2.0/token"
            }
        }
        $Script:MSCloudLoginConnectionProfile.MicrosoftGraph = $this
        Connect-MSCloudLoginMicrosoftGraph
    }

    [void] Disconnect()
    {
        Disconnect-MSCloudLoginMicrosoftGraph
    }
}

class PnP:Workload
{
    [string]
    $Scope

    [string]
    $TokenUrl

    [string]
    $ConnectionUrl

    [string]
    $ClientId = '9bc3ab49-b65d-410a-85ad-de819febfddc' # Microsoft Sharepoint Online Management Shell

    [string]
    $RedirectURI = 'https://oauth.spops.microsoft.com/'

    [string]
    $AdminUrl

    [string]
    [ValidateSet('Production', 'PPE', 'China', 'Germany', 'USGovernment', 'USGovernmentHigh', 'USGovernmentDoD', 'Custom')]
    $PnPAzureEnvironment

    PnP()
    {
        if (-not [String]::IsNullOrEmpty($this.CertificateThumbprint) -and (-not[String]::IsNullOrEmpty($this.CertificatePassword) -or
                -not[String]::IsNullOrEmpty($this.CertificatePath))
        )
        {
            throw 'Cannot specify both a Certificate Thumbprint and Certificate Path and Password'
        }
    }

    [void] Connect([boolean]$ForceRefresh)
    {
        ([Workload]$this).Setup()

        # PnP uses Production instead of AzureCloud to designate the Public Azure Cloud * AzureUSGovernment to USGovernmentHigh
        if ($null -ne $this.Endpoints)
        {
            $this.PnPAzureEnvironment = 'Custom'
            $this.Scope               = $Global:CustomPnPScope
            $this.TokenUrl            = "$($Global:CustomPnPTokenUrl)/$($this.TenantId)/oauth2/v2.0/token"
        }
        elseif ($this.EnvironmentName -eq 'AzureCloud')
        {
            $this.PnPAzureEnvironment = 'Production'
        }
        elseif ($this.EnvironmentName -eq 'AzureUSGovernment')
        {
            $this.PnPAzureEnvironment = 'USGovernmentHigh'
        }
        elseif ($this.EnvironmentName -eq 'AzureDOD')
        {
            $this.PnPAzureEnvironment = 'USGovernmentDoD'
        }
        elseif ($this.EnvironmentName -eq 'AzureGermany')
        {
            $this.PnPAzureEnvironment = 'Germany'
        }
        elseif ($this.EnvironmentName -eq 'AzureChinaCloud')
        {
            $this.PnPAzureEnvironment = 'China'
        }
        $Script:MSCloudLoginConnectionProfile.PnP = $this
        Connect-MSCloudLoginPnP -ForceRefreshConnection $ForceRefresh
    }

    [void] Disconnect()
    {
        Disconnect-MSCloudLoginPnP
    }
}

class PowerPlatform:Workload
{
    [string]
    $Endpoint = 'prod'

    PowerPlatform()
    {
    }

    [void] Connect()
    {
        ([Workload]$this).Setup()
        Connect-MSCloudLoginPowerPlatform
    }
    [void] Disconnect()
    {
        $this.Connected = $false
    }
}

class PowerPlatformREST:Workload
{
    [string]
    $AuthorizationUrl

    [string]
    $Audience

    [string]
    $BapEndpoint

    [string]
    $ClientId

    [string]
    $Scope

    [string]
    $AccessToken

    PowerPlatformREST()
    {
    }

    [void] Connect()
    {
        ([Workload]$this).Setup()

        switch ($this.EnvironmentName)
        {
            <# AUDIENCE
            "prod" = "https://service.powerapps.com/";
            "preview" = "https://service.powerapps.com/";
            "tip1"= "https://service.powerapps.com/";
            "tip2"= "https://service.powerapps.com/";
            "usgov"= "https://gov.service.powerapps.us/";
            "usgovhigh"= "https://high.service.powerapps.us/";
            "dod" = "https://service.apps.appsplatform.us/";
            "china" = "https://service.powerapps.cn/";
            #>

            <# BAP
                        "prod"      { "api.bap.microsoft.com" }
                        "usgov"     { "gov.api.bap.microsoft.us" }
                        "usgovhigh" { "high.api.bap.microsoft.us" }
                        "dod"       { "api.bap.appsplatform.us" }
                        "china"     { "api.bap.partner.microsoftonline.cn" }
                        "preview"   { "preview.api.bap.microsoft.com" }
                        "tip1"      { "tip1.api.bap.microsoft.com"}
                        "tip2"      { "tip2.api.bap.microsoft.com" }
            #>
            'AzureDOD'
            {
                $this.Scope            = "6a8b4b39-c021-437c-b060-5a14a3fd65f3/.default"
                $this.AuthorizationUrl = "https://login.microsoftonline.us"
                $this.Audience         = "https://service.apps.appsplatform.us/"
                $this.ClientId         = "1950a258-227b-4e31-a9cf-717495945fc2"
                $this.BapEndpoint      = "api.bap.appsplatform.us"

            }
            'AzureUSGovernment'
            {
                $this.Scope            = "6a8b4b39-c021-437c-b060-5a14a3fd65f3/.default"
                $this.AuthorizationUrl = "https://login.microsoftonline.us"
                $this.Audience         = "https://gov.service.powerapps.us/"
                $this.ClientId         = "1950a258-227b-4e31-a9cf-717495945fc2"
                $this.BapEndpoint      = "gov.api.bap.microsoft.us"
            }
            'Custom'
            {
                $this.Scope            = $Global:CustomPowerPlatformRESTScope
                $this.AuthorizationUrl = $Global:CustomPowerPlatformRESTAuthorizationUrl
                $this.Audience         = $Global:CustomPowerPlatformRESTAudience
                $this.ClientId         = $Global:CustomPowerPlatformRESTClientId
                $this.BapEndpoint      = $Global:CustomPowerPlatformRESTBapEndpoint
            }
            default
            {
                $this.Scope            = "6a8b4b39-c021-437c-b060-5a14a3fd65f3/.default"
                $this.AuthorizationUrl = "https://login.microsoftonline.com"
                $this.Audience         = "https://service.powerapps.com/"
                $this.ClientId         = "1950a258-227b-4e31-a9cf-717495945fc2"
                $this.BapEndpoint      = "api.bap.microsoft.com"
            }
        }
        $Script:MSCloudLoginConnectionProfile.PowerPlatformREST = $this
        Connect-MSCloudLoginPowerPlatformREST
    }
}

class SecurityComplianceCenter:Workload
{
    [boolean]
    $SkipModuleReload = $false

    [string]
    $ConnectionUrl

    [string]
    $AuthorizationUrl

    [string]
    $AzureADAuthorizationEndpointUri

    SecurityComplianceCenter()
    {
    }

    [void] Connect()
    {
        ([Workload]$this).Setup()

        switch ($this.EnvironmentName)
        {
            'AzureCloud'
            {
                $this.ConnectionUrl    = 'https://ps.compliance.protection.outlook.com/powershell-liveid/'
                $this.AuthorizationUrl = 'https://login.microsoftonline.com/organizations'
            }
            'AzureUSGovernment'
            {
                $this.ConnectionUrl                   = 'https://ps.compliance.protection.office365.us/powershell-liveid/'
                $this.AuthorizationUrl                = 'https://login.microsoftonline.us/organizations'
                $this.AzureADAuthorizationEndpointUri = 'https://login.microsoftonline.us/common'
            }
            'AzureDOD'
            {
                $this.ConnectionUrl                   = 'https://l5.ps.compliance.protection.office365.us/powershell-liveid/'
                $this.AuthorizationUrl                = 'https://login.microsoftonline.us/organizations'
                $this.AzureADAuthorizationEndpointUri = 'https://login.microsoftonline.us/common'
            }
            'AzureGermany'
            {
                $this.ConnectionUrl    = 'https://ps.compliance.protection.outlook.de/powershell-liveid/'
                $this.AuthorizationUrl = 'https://login.microsoftonline.de/organizations'
            }
            'AzureChinaCloud'
            {
                $this.ConnectionUrl    = 'https://ps.compliance.protection.partner.outlook.cn/powershell-liveid/'
                $this.AuthorizationUrl = 'https://login.chinacloudapi.cn/organizations'
            }
            'Custom'
            {
                $this.ConnectionUrl                   = $Global:CustomSCCConnectionUrl
                $this.AuthorizationUrl                = $Global:CustomSCCAuthorizationUrl
                $this.AzureADAuthorizationEndpointUri = $Global:CustomSCCAzureADAuthorizationEndpointUri
            }
        }
        $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter = $this
        Connect-MSCloudLoginSecurityCompliance
    }
}

class SharePointOnlineREST:Workload
{
    [string]
    $AdminUrl

    [string]
    $ConnectionUrl

    [string]
    $HostUrl

    [string]
    $AuthorizationUrl

    [string]
    $Scope

    [string]
    $AccessToken

    SharePointOnlineREST()
    {
    }

    [void] Connect()
    {
        ([Workload]$this).Setup()

        # Retrieve the SPO Admin URL
        if ($Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.AuthenticationType -eq 'Credentials' -and `
            -not $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.AdminUrl)
        {
            $this.AdminUrl = Get-SPOAdminUrl -Credential $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.Credentials
            if ([String]::IsNullOrEmpty($this.AdminUrl) -eq $false)
            {
                $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.AdminUrl = $this.AdminUrl
                $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.ConnectionUrl = $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.AdminUrl
            }
            else
            {
                throw 'Unable to retrieve SharePoint Admin Url. Check if the Graph can be contacted successfully.'
            }
        }
        elseif (-not $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.AdminUrl -and `
                -not [System.String]::IsNullOrEmpty($Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.TenantId))
        {
            if ($Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.TenantId.Contains('onmicrosoft'))
            {
                $domain = $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.TenantId.Replace('.onmicrosoft.', '-admin.sharepoint.')
                if (-not $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.AdminUrl)
                {
                    $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.AdminUrl = "https://$domain"
                }
                $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.ConnectionUrl = ("https://$domain").Replace('-admin', '')
            }
            elseif ($Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.TenantId.Contains('.onmschina.'))
            {
                $domain = $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.TenantId.Replace('.partner.onmschina.', '-admin.sharepoint.')
                if (-not $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.AdminUrl)
                {
                    $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.AdminUrl = "https://$domain"
                }
                $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.ConnectionUrl = ("https://$domain").Replace('-admin', '')
            }
            elseif ($Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.TenantId.Contains('.onms.'))
            {
                $domain = $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.TenantId.Replace('.onms.', '-admin.sharepoint.')
                if (-not $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.AdminUrl)
                {
                    $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.AdminUrl = "https://$domain"
                }
                $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.ConnectionUrl = ("https://$domain").Replace('-admin', '')
            }
            else
            {
                throw 'TenantId must be in format contoso.onmicrosoft.com'
            }
        }

        switch ($this.EnvironmentName)
        {
            'AzureDOD'
            {
                $this.HostUrl          = $this.AdminUrl
                $this.Scope            = "$($this.AdminUrl)/.default"
                $this.AuthorizationUrl = "https://login.microsoftonline.us"
            }
            'AzureUSGovernment'
            {
                $this.HostUrl          = $this.AdminUrl
                $this.Scope            = "$($this.AdminUrl)/.default"
                $this.AuthorizationUrl = "https://login.microsoftonline.us"
            }
            'Custom'
            {
                $this.HostUrl          = $this.Endpoints.HostUrl
                $this.Scope            = $this.Endpoints.Scope
                $this.AuthorizationUrl = $this.Endpoints.AuthorizationUrl
            }
            default
            {
                $this.HostUrl          = $this.AdminUrl
                $this.Scope            = "$($this.AdminUrl)/.default"
                $this.AuthorizationUrl = "https://login.microsoftonline.com"
            }
        }
        $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST = $this
        Connect-MSCloudLoginSharePointOnlineREST
    }
}

class Tasks:Workload
{
    [string]
    $HostUrl

    [string]
    $AuthorizationUrl

    [string]
    $ResourceUrl

    [string]
    $Scope

    [string]
    $AccessToken

    Tasks()
    {
    }

    [void] Connect()
    {
        ([Workload]$this).Setup()
        switch ($this.EnvironmentName)
        {
            'AzureDOD'
            {
                $this.HostUrl          = "https://tasks.office.us"
                $this.Scope            = "https://tasks.office.us/.default"
                $this.AuthorizationUrl = "https://login.microsoftonline.us"
                $this.ResourceUrl      = "https://tasks.osi.apps.mil"
            }
            'AzureUSGovernment'
            {
                $this.HostUrl          = "https://tasks.office.us"
                $this.Scope            = "https://tasks.office365.us/.default"
                $this.AuthorizationUrl = "https://login.microsoftonline.us"
                $this.ResourceUrl      = "https://tasks.office365.us"
            }
            'Custom'
            {
                $this.HostUrl          = $this.Endpoints.HostUrl
                $this.Scope            = $this.Endpoints.Scope
                $this.AuthorizationUrl = $this.Endpoints.AuthorizationUrl
                $this.ResourceUrl      = $this.Endpoints.ResourceUrl
            }
            default
            {
                $this.HostUrl          = "https://tasks.office.com"
                $this.Scope            = "https://tasks.office.com/.default"
                $this.AuthorizationUrl = "https://login.microsoftonline.com"
                $this.ResourceUrl      = "https://tasks.office.com"
            }
        }
        $Script:MSCloudLoginConnectionProfile.Tasks = $this
        Connect-MSCloudLoginTasks
    }
}

class Teams:Workload
{
    [string]
    $TokenUrl

    [string]
    $GraphScope

    [string]
    $TeamsScope

    Teams()
    {
    }

    [void] Connect()
    {
        ([Workload]$this).Setup()
        switch ($this.EnvironmentName)
        {
            "Custom"
            {
                $this.TokenUrl   = "$($Global:CustomTeamsTokenUrl)/$($this.TenantId)/oauth2/v2.0/token"
                $this.GraphScope = $Global:CustomGraphScope
                $this.TeamsScope = $Global:CustomTeamsScope
            }
        }
        $Script:MSCloudLoginConnectionProfile.Teams = $this
        Connect-MSCloudLoginTeams
    }

    [void] Disconnect()
    {
        Disconnect-MSCloudLoginTeams
    }
}
