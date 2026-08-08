@{
    # Per-workload, per-environment endpoint data used by the workload classes in
    # ConnectionProfile.ps1 and resolved through Get-MSCloudLoginEndpointInfo.
    #
    # Conventions:
    # - 'default' is used for every environment that has no explicit entry.
    # - The values of a 'Custom' entry are NOT literal values but the names of the
    #   corresponding keys in the custom environment configuration file
    #   (CustomEnvironment.psd1), resolved at runtime.
    # - '{Resource}', '{AdminUrl}' and '{TenantId}' placeholders are replaced at runtime
    #   with the corresponding workload profile values.
    # - AzureFranceCloud entries follow the pattern of the pre-existing France entries:
    #   the sovereign authorization endpoint (login.sovcloud-identity.fr) combined with
    #   the commercial service endpoints unless a dedicated sovereign endpoint is known.

    AdminAPI              = @{
        AzureDOD          = @{ Scope = '{Resource}/.default'; AuthorizationUrl = 'https://login.microsoftonline.us' }
        AzureUSGovernment = @{ Scope = '{Resource}/.default'; AuthorizationUrl = 'https://login.microsoftonline.us' }
        AzureFranceCloud  = @{ Scope = '{Resource}/.default'; AuthorizationUrl = 'https://login.sovcloud-identity.fr' }
        Custom            = @{ Scope = 'CustomAdminApiScope'; AuthorizationUrl = 'CustomAdminApiAuthorizationUrl' }
        default           = @{ Scope = '{Resource}/.default'; AuthorizationUrl = 'https://login.microsoftonline.com' }
    }

    AzureDevOPS           = @{
        AzureDOD          = @{ HostUrl = 'https://dev.azure.us'; Scope = '{Resource}/.default'; AuthorizationUrl = 'https://login.microsoftonline.us' }
        AzureUSGovernment = @{ HostUrl = 'https://dev.azure.com'; Scope = '{Resource}/.default'; AuthorizationUrl = 'https://login.microsoftonline.us' }
        AzureFranceCloud  = @{ HostUrl = 'https://dev.azure.com'; Scope = '{Resource}/.default'; AuthorizationUrl = 'https://login.sovcloud-identity.fr' }
        Custom            = @{ HostUrl = 'CustomAzureDevopsHostUrl'; Scope = 'CustomAzureDevopsScope'; AuthorizationUrl = 'CustomAzureDevopsAuthorizationUrl' }
        default           = @{ HostUrl = 'https://dev.azure.com'; Scope = '{Resource}/.default'; AuthorizationUrl = 'https://login.microsoftonline.com' }
    }

    DefenderForEndpoint   = @{
        AzureDOD          = @{ HostUrl = 'https://api-gov.securitycenter.microsoft.us'; Scope = 'https://api.securitycenter.microsoft.com/.default'; AuthorizationUrl = 'https://login.microsoftonline.us' }
        AzureUSGovernment = @{ HostUrl = 'https://api-gcc.securitycenter.microsoft.us'; Scope = 'https://api.securitycenter.microsoft.com/.default'; AuthorizationUrl = 'https://login.microsoftonline.com' }
        AzureFranceCloud  = @{ HostUrl = 'https://api.securitycenter.microsoft.com/'; Scope = 'https://api.securitycenter.microsoft.com/.default'; AuthorizationUrl = 'https://login.sovcloud-identity.fr' }
        Custom            = @{ HostUrl = 'CustomDefenderForEndpointHostUrl'; Scope = 'CustomDefenderForEndpointScope'; AuthorizationUrl = 'CustomDefenderForEndpointAuthorizationUrl' }
        default           = @{ HostUrl = 'https://api.securitycenter.microsoft.com/'; Scope = 'https://api.securitycenter.microsoft.com/.default'; AuthorizationUrl = 'https://login.microsoftonline.com' }
    }

    EngageHub             = @{
        AzureDOD          = @{ ClientId = ''; Scope = 'https://engagehub.microsoft.us/.default'; AuthorizationUrl = 'https://login.microsoftonline.us'; APIUrl = 'https://api.engagecenter.microsoft.us' }
        AzureUSGovernment = @{ ClientId = ''; Scope = 'https://engagehub.microsoft.us/.default'; AuthorizationUrl = 'https://login.microsoftonline.us'; APIUrl = 'https://api.engagecenter.microsoft.us' }
        AzureFranceCloud  = @{ ClientId = ''; Scope = 'https://engagehub.microsoft.com/.default'; AuthorizationUrl = 'https://login.sovcloud-identity.fr'; APIUrl = 'https://api.engagecenter.microsoft.com' }
        Custom            = @{ ClientId = 'CustomEngageHubClientId'; Scope = 'CustomEngageHubScope'; AuthorizationUrl = 'CustomEngageHubAuthorizationUrl'; APIUrl = 'CustomEngageHubAPIUrl' }
        default           = @{ ClientId = ''; Scope = 'https://engagehub.microsoft.com/.default'; AuthorizationUrl = 'https://login.microsoftonline.com'; APIUrl = 'https://api.engagecenter.microsoft.com' }
    }

    Fabric                = @{
        AzureDOD          = @{ HostUrl = 'https://api.fabric.microsoft.us'; Scope = 'https://api.fabric.microsoft.us/.default'; AuthorizationUrl = 'https://login.microsoftonline.us' }
        AzureUSGovernment = @{ HostUrl = 'https://api.fabric.microsoft.us'; Scope = 'https://api.fabric.microsoft.us/.default'; AuthorizationUrl = 'https://login.microsoftonline.us' }
        AzureFranceCloud  = @{ HostUrl = 'https://api.fabric.microsoft.com'; Scope = 'https://api.fabric.microsoft.com/.default'; AuthorizationUrl = 'https://login.sovcloud-identity.fr' }
        Custom            = @{ HostUrl = 'CustomFabricHostUrl'; Scope = 'CustomFabricScope'; AuthorizationUrl = 'CustomFabricAuthorizationUrl' }
        default           = @{ HostUrl = 'https://api.fabric.microsoft.com'; Scope = 'https://api.fabric.microsoft.com/.default'; AuthorizationUrl = 'https://login.microsoftonline.com' }
    }

    Licensing             = @{
        AzureDOD          = @{ HostUrl = 'https://licensing.m365.microsoft.com'; Scope = '{Resource}/.default'; AuthorizationUrl = 'https://login.microsoftonline.us' }
        AzureUSGovernment = @{ HostUrl = 'https://licensing.m365.microsoft.com'; Scope = '{Resource}/.default'; AuthorizationUrl = 'https://login.microsoftonline.us' }
        AzureFranceCloud  = @{ HostUrl = 'https://licensing.m365.microsoft.com'; Scope = '{Resource}/.default'; AuthorizationUrl = 'https://login.sovcloud-identity.fr' }
        Custom            = @{ HostUrl = 'CustomLicensingHostUrl'; Scope = 'CustomLicensingScope'; AuthorizationUrl = 'CustomLicensingAuthorizationUrl' }
        default           = @{ HostUrl = 'https://licensing.m365.microsoft.com'; Scope = '{Resource}/.default'; AuthorizationUrl = 'https://login.microsoftonline.com' }
    }

    O365Portal            = @{
        AzureDOD          = @{ HostUrl = 'https://portal.apps.mil'; Scope = 'https://portal.apps.mil/.default'; AuthorizationUrl = 'https://login.microsoftonline.us' }
        AzureUSGovernment = @{ HostUrl = 'https://portal.office365.us'; Scope = 'https://portal.office365.us/.default'; AuthorizationUrl = 'https://login.microsoftonline.us' }
        AzureFranceCloud  = @{ HostUrl = 'https://admin.microsoft.com'; Scope = 'https://admin.microsoft.com/.default'; AuthorizationUrl = 'https://login.sovcloud-identity.fr' }
        Custom            = @{ HostUrl = 'CustomO365PortalHostUrl'; Scope = 'CustomO365PortalScope'; AuthorizationUrl = 'CustomO365PortalAuthorizationUrl' }
        default           = @{ HostUrl = 'https://admin.microsoft.com'; Scope = 'https://admin.microsoft.com/.default'; AuthorizationUrl = 'https://login.microsoftonline.com' }
    }

    PowerPlatformREST     = @{
        AzureDOD          = @{ Scope = 'https://service.apps.appsplatform.us/.default'; AuthorizationUrl = 'https://login.microsoftonline.us'; Audience = 'https://service.apps.appsplatform.us/'; BapEndpoint = 'api.bap.appsplatform.us' }
        AzureUSGovernment = @{ Scope = 'https://gov.service.powerapps.us/.default'; AuthorizationUrl = 'https://login.microsoftonline.us'; Audience = 'https://gov.service.powerapps.us/'; BapEndpoint = 'gov.api.bap.microsoft.us' }
        AzureFranceCloud  = @{ Scope = 'https://service.powerapps.com/.default'; AuthorizationUrl = 'https://login.sovcloud-identity.fr'; Audience = 'https://service.powerapps.com/'; BapEndpoint = 'api.bap.microsoft.com' }
        Custom            = @{ Scope = 'CustomPowerPlatformRESTScope'; AuthorizationUrl = 'CustomPowerPlatformRESTAuthorizationUrl'; Audience = 'CustomPowerPlatformRESTAudience'; ClientId = 'CustomPowerPlatformRESTClientId'; BapEndpoint = 'CustomPowerPlatformRESTBapEndpoint' }
        default           = @{ Scope = 'https://service.powerapps.com/.default'; AuthorizationUrl = 'https://login.microsoftonline.com'; Audience = 'https://service.powerapps.com/'; BapEndpoint = 'api.bap.microsoft.com' }
    }

    SecurityComplianceCenter = @{
        AzureCloud        = @{ ConnectionUrl = 'https://ps.compliance.protection.outlook.com/powershell-liveid/'; AuthorizationUrl = 'https://login.microsoftonline.com/organizations' }
        AzureUSGovernment = @{ ConnectionUrl = 'https://ps.compliance.protection.office365.us/powershell-liveid/'; AuthorizationUrl = 'https://login.microsoftonline.us/organizations' }
        AzureDOD          = @{ ConnectionUrl = 'https://l5.ps.compliance.protection.office365.us/powershell-liveid/'; AuthorizationUrl = 'https://login.microsoftonline.us/organizations' }
        AzureGermany      = @{ ConnectionUrl = 'https://ps.compliance.protection.outlook.de/powershell-liveid/'; AuthorizationUrl = 'https://login.microsoftonline.de/organizations' }
        AzureChinaCloud   = @{ ConnectionUrl = 'https://ps.compliance.protection.partner.outlook.cn/powershell-liveid/'; AuthorizationUrl = 'https://login.chinacloudapi.cn/organizations' }
        AzureFranceCloud  = @{ ConnectionUrl = 'https://ps.compliance.protection.svc.sovcloud.fr/PowerShell-LiveID'; AuthorizationUrl = 'https://login.sovcloud-identity.fr/organizations' }
        Custom            = @{ ConnectionUrl = 'CustomSCCConnectionUrl'; AuthorizationUrl = 'CustomSCCAzureADAuthorizationEndpointUri' }
        default           = @{ ConnectionUrl = 'https://ps.compliance.protection.outlook.com/powershell-liveid/'; AuthorizationUrl = 'https://login.microsoftonline.com/organizations' }
    }

    SharePointOnlineREST  = @{
        AzureDOD          = @{ HostUrl = '{AdminUrl}'; Scope = '{AdminUrl}/.default'; AuthorizationUrl = 'https://login.microsoftonline.us' }
        AzureUSGovernment = @{ HostUrl = '{AdminUrl}'; Scope = '{AdminUrl}/.default'; AuthorizationUrl = 'https://login.microsoftonline.us' }
        AzureFranceCloud  = @{ HostUrl = '{AdminUrl}'; Scope = '{AdminUrl}/.default'; AuthorizationUrl = 'https://login.sovcloud-identity.fr' }
        # Custom has no dedicated Scope key: the class derives Scope from HostUrl + '/.default'.
        Custom            = @{ HostUrl = 'CustomSharePointOnlineRESTHostUrl'; AuthorizationUrl = 'CustomSharePointOnlineRESTAuthorizationUrl' }
        default           = @{ HostUrl = '{AdminUrl}'; Scope = '{AdminUrl}/.default'; AuthorizationUrl = 'https://login.microsoftonline.com' }
    }

    Tasks                 = @{
        AzureDOD          = @{ HostUrl = 'https://tasks.osi.apps.mil'; Scope = 'https://tasks.osi.apps.mil/.default'; AuthorizationUrl = 'https://login.microsoftonline.us'; ResourceUrl = 'https://tasks.osi.apps.mil' }
        AzureUSGovernment = @{ HostUrl = 'https://tasks.office365.us'; Scope = 'https://tasks.office365.us/.default'; AuthorizationUrl = 'https://login.microsoftonline.us'; ResourceUrl = 'https://tasks.office365.us' }
        AzureFranceCloud  = @{ HostUrl = 'https://tasks.office.com'; Scope = 'https://tasks.office.com/.default'; AuthorizationUrl = 'https://login.sovcloud-identity.fr'; ResourceUrl = 'https://tasks.office.com' }
        Custom            = @{ HostUrl = 'CustomTasksHostUrl'; Scope = 'CustomTasksScope'; AuthorizationUrl = 'CustomTasksAuthorizationUrl'; ResourceUrl = 'CustomTasksResourceUrl' }
        default           = @{ HostUrl = 'https://tasks.office.com'; Scope = 'https://tasks.office.com/.default'; AuthorizationUrl = 'https://login.microsoftonline.com'; ResourceUrl = 'https://tasks.office.com' }
    }
}
