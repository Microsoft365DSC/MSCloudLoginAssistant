# IMPORTANT!  If you are using this module for a custom environment, ensure you keep a copy in a secure location. Installation of a newer version of MSCLoudLoginAssistant will overwrite this file.
# After installation of the newest version, restore your backup file to this location.

@{
    # If you are running DSC in a custom environment without access to WW resources, set this value to $true, otherwise leave it set to $false.
    CustomEnvironment = $false

    # If you are running DSC in a custom environment without access to WW resources, edit the endpoints below to match your environment's values.

    CustomAdminApiScope = "6a8b4b39-c021-437c-b060-5a14a3fd65f3/.default"
    CustomAdminApiAuthorizationUrl = "https://login.microsoftonline.com"

    CustomAzureDevopsHostUrl = "https://dev.azure.com"
    CustomAzureDevopsScope = "499b84ac-1321-427f-aa17-267ca6975798/.default"
    CustomAzureDevopsAuthorizationUrl = "https://login.microsoftonline.com"

    CustomDefenderForEndpointHostUrl = "https://api.security.microsoft.com"
    CustomDefenderForEndpointScope = "https://api.securitycenter.microsoft.com/.default"
    CustomDefenderForEndpointAuthorizationUrl = "https://login.microsoftonline.com"

    CustomEngageHubClientId = ""
    CustomEngageHubScope = "https://engagehub.microsoft.com/.default"
    CustomEngageHubAuthorizationUrl = "https://login.microsoftonline.com"
    CustomEngageHubAPIUrl = "https://api.dev.engagecenter.microsoft.com"

    CustomEXOConnectionUri = "https://outlook.prod.microsoft.com/powershell-liveid/"
    CustomEXOAzureADAuthorizationEndpointUri = "https://login.microsoftonline.us/common"

    CustomFabricHostUrl = "https://api.fabric.microsoft.com"
    CustomFabricScope = "https://api.fabric.microsoft.com/.default"
    CustomFabricAuthorizationUrl = "https://login.microsoftonline.com"

    CustomLicensingHostUrl = "https://licensing.m365.microsoft.com"
    CustomLicensingScope = "aeb86249-8ea3-49e2-900b-54cc8e308f85/.default"
    CustomLicensingAuthorizationUrl = "https://login.microsoftonline.com"

    CustomGraphAuthorizationUrl = "https://login.microsoftonline.com"
    CustomGraphResourceUrl = "https://graph.microsoft.com/"
    CustomGraphScope = "https://graph.microsoft.com/.default"
    CustomGraphTokenUrl = "https://login.microsoftonline.com" # No trailing slash!

    CustomO365PortalHostUrl = "https://admin.microsoft.com"
    CustomO365PortalScope = "https://admin.microsoft.com/.default"
    CustomO365PortalAuthorizationUrl = "https://login.microsoftonline.com"

    CustomPnPScope = "https://prod.sharepoint.microsoft.com/.default"
    CustomPnPTokenUrl = "https://login.microsoftonline.com" # No trailing slash!

    CustomPowerPlatformRESTScope = "6a8b4b39-c021-437c-b060-5a14a3fd65f3/.default"
    CustomPowerPlatformRESTAuthorizationUrl = "https://login.microsoftonline.com"
    CustomPowerPlatformRESTAudience = "https://service.powerapps.com/"
    CustomPowerPlatformRESTClientId = "1950a258-227b-4e31-a9cf-717495945fc2"
    CustomPowerPlatformRESTBapEndpoint = "api.bap.microsoft.com"

    CustomSCCConnectionUrl = "https://ps.compliance.protection.outlook.com/powershell-liveid/"
    CustomSCCAuthorizationUrl = "https://login.microsoftonline.com/organizations"
    CustomSCCAzureADAuthorizationEndpointUri = "https://login.microsoftonline.com/common"

    CustomSharePointOnlineRESTHostUrl = "https://customdomain.sharepoint.com" # No trailing slash!
    CustomSharePointOnlineRESTAuthorizationUrl = "https://login.microsoftonline.com"

    CustomTasksHostUrl = "https://tasks.office.com"
    CustomTasksScope = "https://tasks.office.com/.default"
    CustomTasksAuthorizationUrl = "https://login.microsoftonline.com"
    CustomTasksResourceUrl = "https://tasks.office.com"

    CustomTeamsTokenUrl = "https://login.microsoftonline.com" # No trailing slash!
    CustomTeamsScope = "https://api.interfaces.microsoft.com/.default"
    CustomTeamsEndpoints = @{
        ActiveDirectory = "https://login.microsoftonline.com"
        MsGraphEndpointResourceId = "https://graph.microsoft.com/"
        TeamsConfigApiEndpoint = "https://api.interfaces.records.teams.microsoft.com"
    }
}
