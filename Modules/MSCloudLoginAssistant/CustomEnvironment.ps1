# If you are running DSC in a custom environment without access to WW resources, set this value to $true, otherwise leave it set to $false.
$Global:CustomEnvironment = $false

# If you are running DSC in a custom environment without access to WW resources, edit the endpoints below to match your environment's values.

$Global:CustomGraphResourceUrl = "https://graph.microsoft.com/"
$Global:CustomGraphScope = "https://graph.microsoft.com/.default"
$Global:CustomGraphTokenUrl = "https://login.microsoftonline.com/"


$Global:CustomEXOConnectionUri = "https://outlook.prod.microsoft.com/powersehll-liveid/"
$Global:CustomEXOAzureADAuthorizationEndpointUri = "https://login.microsoftonline.us/common"

$Global:CustomAdminApiScope = "6a8b4b39-c021-437c-b060-5a14a3fd65f3/.default"
$Global:CustomAdminApiAuthorizationUrl = "https://login.microsoftonline.com"

$Global:CustomAzureDevopsHostUrl = "https://dev.azure.com"
$Global:CustomAzureDevopsScope = "499b84ac-1321-427f-aa17-267ca6975798/.default"
$Global:CustomAzureDevopsAuthorizationUrl = "https://login.microsoftonline.com"

$Global:CustomDefenderForEndpointHostUrl = "https://api.security.microsoft.com"
$Global:CustomDefenderForEndpointScope = "https://api.securitycenter.microsoft.com/.default"
$Global:CustomDefenderForEndpointAuthorizationUrl = "https://login.microsoftonline.com"

$Global:CustomEngageHubScope = "https://engagehub.microsoft.com/.default"
$Global:CustomEngageHubAuthorizationUrl = "https://login.microsoftonline.com"
$Global:CustomEngageHubAPIUrl = "https://api.dev.engagecenter.microsoft.com"

$Global:CustomFabricHostUrl = "https://api.fabric.microsoft.com"
$Global:CustomFabricScope = "https://api.fabric.microsoft.com/.default"
$Global:CustomFabricAuthorizationUrl = "https://login.microsoftonline.com"

$Global:CustomLicensingHostUrl = "https://licensing.m365.microsoft.com"
$Global:CustomLicensingScope = "aeb86249-8ea3-49e2-900b-54cc8e308f85/.default"
$Global:CustomLicensingAuthorizationUrl = "https://login.microsoftonline.com"

$Global:CustomPnPScope = "https://prod.sharepoint.microsoft.com/.default"
$Global:CustomPnPTokenUrl = "https://login.microsoftonline.com/"

$Global:CustomPowerPlatformRESTScope = "6a8b4b39-c021-437c-b060-5a14a3fd65f3/.default"
$Global:CustomPowerPlatformRESTAuthorizationUrl = "https://login.microsoftonline.com"
$Global:CustomPowerPlatformRESTAudience = "https://service.powerapps.com/"
$Global:CustomPowerPlatformRESTClientId = "1950a258-227b-4e31-a9cf-717495945fc2"
$Global:CustomPowerPlatformRESTBapEndpoint = "api.bap.microsoft.com"

$Global:CustomSCCConnectionUrl = "https://ps.compliance.microsoft.com/powershell-liveid/"
$Global:CustomSCCAuthorizationUrl = "https://login.microsoftonline.com/organiations"
$Global:CustomSCCAzureADAuthorizationEndpointUri = "https://login.microsoftonline.com/common"

$Global:CustomTeamsTokenUrl = "https://login.microsoftonline.com/"
$Global:CustomTeamsScope = "https://api.interfaces.microsoft.com/.default"
$Global:CustomTeamsEndpoints = @{
    ActiveDirectory = "https://login.microsoftonline.com"
    MsGraphEndpointResourceId = "https://graph.microsoft.com/"
    TeamsConfigApiEndpoint = "https://api.interfaces.records.teams.microsoft.com"
}
