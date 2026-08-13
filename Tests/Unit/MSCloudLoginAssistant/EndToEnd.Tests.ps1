#Requires -Modules Pester

BeforeAll {
    $graphModuleName = 'Microsoft.Graph.Beta.Identity.DirectoryManagement'
    if (-not (Get-Module -Name $graphModuleName -ListAvailable))
    {
        $script:tempModuleBase = Join-Path $env:TEMP 'MSCloudLoginTestModules'
        $tempModuleDir = Join-Path $script:tempModuleBase $graphModuleName
        if (-not (Test-Path $tempModuleDir))
        {
            New-Item -Path $tempModuleDir -ItemType Directory -Force | Out-Null
        }
        $manifestPath = Join-Path $tempModuleDir "$graphModuleName.psd1"
        if (-not (Test-Path $manifestPath))
        {
            New-ModuleManifest -Path $manifestPath -ModuleVersion '1.0.0' -Description 'Test stub'
        }
        $env:PSModulePath = $script:tempModuleBase + [IO.Path]::PathSeparator + $env:PSModulePath
    }

    Import-Module (Join-Path $PSScriptRoot '..\Stubs\Stubs.psm1') -Force -Global -WarningAction SilentlyContinue
    $script:moduleRoot = Resolve-Path (Join-Path $PSScriptRoot '..\..\..\Modules\MSCloudLoginAssistant')
    Import-Module (Join-Path $script:moduleRoot 'MSCloudLoginAssistant.psd1') -Force
}

AfterAll {
    if ($script:tempModuleBase -and (Test-Path $script:tempModuleBase))
    {
        Remove-Item -Path $script:tempModuleBase -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Describe 'Connect-M365Tenant end-to-end for REST workloads' {

    BeforeEach {
        InModuleScope 'MSCloudLoginAssistant' {
            $Script:MSCloudLoginConnectionProfile = $null
            $Script:MSCloudLoginTriedGetEnvironment = $false
            $Script:CloudEnvironmentInfo = $null
        }
    }

    Context 'When connecting with a service principal and an application secret' {
        It 'Should acquire a bearer token and expose the commercial <Workload> endpoints on the connection profile' -TestCases @(
            @{ Workload = 'AdminAPI'; ExpectedScope = '6a8b4b39-c021-437c-b060-5a14a3fd65f3/.default'; ExpectedHostProperty = ''; ExpectedHostUrl = '' }
            @{ Workload = 'AzureDevOPS'; ExpectedScope = '499b84ac-1321-427f-aa17-267ca6975798/.default'; ExpectedHostProperty = 'HostUrl'; ExpectedHostUrl = 'https://dev.azure.com' }
            @{ Workload = 'DefenderForEndpoint'; ExpectedScope = 'https://api.securitycenter.microsoft.com/.default'; ExpectedHostProperty = 'HostUrl'; ExpectedHostUrl = 'https://api.securitycenter.microsoft.com/' }
            @{ Workload = 'EngageHub'; ExpectedScope = 'https://engagehub.microsoft.com/.default'; ExpectedHostProperty = 'APIUrl'; ExpectedHostUrl = 'https://api.engagecenter.microsoft.com' }
            @{ Workload = 'Fabric'; ExpectedScope = 'https://api.fabric.microsoft.com/.default'; ExpectedHostProperty = 'HostUrl'; ExpectedHostUrl = 'https://api.fabric.microsoft.com' }
            @{ Workload = 'Licensing'; ExpectedScope = 'aeb86249-8ea3-49e2-900b-54cc8e308f85/.default'; ExpectedHostProperty = 'HostUrl'; ExpectedHostUrl = 'https://licensing.m365.microsoft.com' }
            @{ Workload = 'PowerPlatformREST'; ExpectedScope = 'https://service.powerapps.com/.default'; ExpectedHostProperty = 'BapEndpoint'; ExpectedHostUrl = 'api.bap.microsoft.com' }
            @{ Workload = 'Tasks'; ExpectedScope = 'https://tasks.office.com/.default'; ExpectedHostProperty = 'ResourceUrl'; ExpectedHostUrl = 'https://tasks.office.com' }
        ) {
            param ($Workload, $ExpectedScope, $ExpectedHostProperty, $ExpectedHostUrl)
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{
                Workload             = $Workload
                ExpectedScope        = $ExpectedScope
                ExpectedHostProperty = $ExpectedHostProperty
                ExpectedHostUrl      = $ExpectedHostUrl
            } {
                param ($Workload, $ExpectedScope, $ExpectedHostProperty, $ExpectedHostUrl)

                Mock -CommandName Invoke-WebRequest -MockWith {
                    return @{ Content = '{ "token_endpoint": "https://login.microsoftonline.com/tenant/oauth2/v2.0/token" }' }
                }
                Mock -CommandName Invoke-RestMethod -MockWith {
                    return @{ token_type = 'Bearer'; access_token = 'e2e-token' }
                }

                Connect-M365Tenant -Workload $Workload `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -ApplicationSecret 'super-secret'

                $connection = Get-MSCloudLoginConnectionProfile -Workload $Workload

                $connection.Connected | Should -BeTrue
                $connection.AuthenticationType | Should -Be 'ServicePrincipalWithSecret'
                $connection.EnvironmentName | Should -Be 'AzureCloud'
                $connection.AccessToken | Should -Be 'Bearer e2e-token'
                $connection.AuthorizationUrl | Should -Be 'https://login.microsoftonline.com'
                $connection.Scope | Should -Be $ExpectedScope
                if (-not [System.String]::IsNullOrEmpty($ExpectedHostProperty))
                {
                    $connection.$ExpectedHostProperty | Should -Be $ExpectedHostUrl
                }

                Should -Invoke Invoke-RestMethod -Exactly 1 -ParameterFilter {
                    $Uri -eq 'https://login.microsoftonline.com/contoso.onmicrosoft.com/oauth2/v2.0/token' -and
                    $Method -eq 'Post' -and
                    $Body.grant_type -eq 'client_credentials' -and
                    $Body.client_id -eq '11111111-1111-1111-1111-111111111111' -and
                    $Body.client_secret -eq 'super-secret'
                }
            }
        }

        It 'Should target the US Government endpoints when the tenant reports the USGov region' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Invoke-WebRequest -MockWith {
                    return @{ Content = '{ "tenant_region_sub_scope": "USGov", "token_endpoint": "https://login.microsoftonline.us/t/oauth2/v2.0/token" }' }
                }
                Mock -CommandName Invoke-RestMethod -MockWith {
                    return @{ token_type = 'Bearer'; access_token = 'e2e-token' }
                }

                Connect-M365Tenant -Workload 'Tasks' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -ApplicationSecret 'super-secret'

                $connection = Get-MSCloudLoginConnectionProfile -Workload 'Tasks'

                $connection.EnvironmentName | Should -Be 'AzureUSGovernment'
                $connection.HostUrl | Should -Be 'https://tasks.office365.us'
                $connection.Scope | Should -Be 'https://tasks.office365.us/.default'
                Should -Invoke Invoke-RestMethod -ParameterFilter {
                    $Uri -eq 'https://login.microsoftonline.us/contoso.onmicrosoft.com/oauth2/v2.0/token'
                }
            }
        }

        It 'Should target the DoD endpoints when the tenant reports the DOD region' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Invoke-WebRequest -MockWith {
                    return @{ Content = '{ "tenant_region_sub_scope": "DOD", "token_endpoint": "https://login.microsoftonline.us/t/oauth2/v2.0/token" }' }
                }
                Mock -CommandName Invoke-RestMethod -MockWith {
                    return @{ token_type = 'Bearer'; access_token = 'e2e-token' }
                }

                Connect-M365Tenant -Workload 'Tasks' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -ApplicationSecret 'super-secret'

                $connection = Get-MSCloudLoginConnectionProfile -Workload 'Tasks'
                $connection.EnvironmentName | Should -Be 'AzureDOD'
                $connection.HostUrl | Should -Be 'https://tasks.osi.apps.mil'
            }
        }
    }

    Context 'When connecting with a certificate thumbprint' {
        It 'Should sign a client assertion with the certificate and store the returned bearer token' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Invoke-WebRequest -MockWith {
                    return @{ Content = '{ "token_endpoint": "https://login.microsoftonline.com/tenant/oauth2/v2.0/token" }' }
                }
                Mock -CommandName Invoke-RestMethod -MockWith {
                    return @{ token_type = 'Bearer'; access_token = 'e2e-token' }
                }
                Mock -CommandName Get-MSCloudLoginCertificate -MockWith {
                    $rsa = [System.Security.Cryptography.RSA]::Create(2048)
                    $request = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new(
                        [System.Security.Cryptography.X509Certificates.X500DistinguishedName]::new('CN=MSCloudLoginAssistantE2E'),
                        $rsa,
                        [System.Security.Cryptography.HashAlgorithmName]::SHA256,
                        [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
                    return $request.CreateSelfSigned([System.DateTimeOffset]::Now.AddDays(-1), [System.DateTimeOffset]::Now.AddDays(1))
                }

                Connect-M365Tenant -Workload 'Fabric' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -CertificateThumbprint 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD'

                $connection = Get-MSCloudLoginConnectionProfile -Workload 'Fabric'

                $connection.Connected | Should -BeTrue
                $connection.AuthenticationType | Should -Be 'ServicePrincipalWithThumbprint'
                $connection.AccessToken | Should -Be 'Bearer e2e-token'

                Should -Invoke Invoke-RestMethod -Exactly 1 -ParameterFilter {
                    $Body.client_assertion_type -eq 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer' -and
                    $Body.client_assertion -match '^[\w-]+\.[\w-]+\.[\w-]+$' -and
                    $Headers.Authorization -eq "Bearer $($Body.client_assertion)"
                }
            }
        }
    }

    Context 'When connecting with a pre-acquired access token' {
        It 'Should reuse the supplied token without contacting the token endpoint' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Invoke-WebRequest -MockWith {
                    return @{ Content = '{ "token_endpoint": "https://login.microsoftonline.com/tenant/oauth2/v2.0/token" }' }
                }
                Mock -CommandName Invoke-RestMethod -MockWith { throw 'the token endpoint must not be contacted' }

                Connect-M365Tenant -Workload 'Licensing' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -AccessTokens @('Bearer caller-supplied-token')

                $connection = Get-MSCloudLoginConnectionProfile -Workload 'Licensing'

                $connection.Connected | Should -BeTrue
                $connection.AuthenticationType | Should -Be 'AccessTokens'
                $connection.AccessToken | Should -Be 'Bearer caller-supplied-token'
                Should -Invoke Invoke-RestMethod -Exactly 0
            }
        }

        It 'Should add the Bearer prefix to a token that was supplied without one' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Invoke-WebRequest -MockWith {
                    return @{ Content = '{ "token_endpoint": "https://login.microsoftonline.com/tenant/oauth2/v2.0/token" }' }
                }

                Connect-M365Tenant -Workload 'Licensing' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -AccessTokens @('caller-supplied-token')

                (Get-MSCloudLoginConnectionProfile -Workload 'Licensing').AccessToken | Should -Be 'Bearer caller-supplied-token'
            }
        }
    }

    Context 'When connecting with a managed identity' {
        It 'Should request the instance metadata token for the resource behind the scope' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Invoke-WebRequest -MockWith {
                    return @{ Content = '{ "token_endpoint": "https://login.microsoftonline.com/tenant/oauth2/v2.0/token" }' }
                }
                Mock -CommandName Invoke-RestMethod -MockWith {
                    return @{ access_token = 'managed-identity-token' }
                }

                $savedIdentityEndpoint = $env:IDENTITY_ENDPOINT
                $savedImdsEndpoint = $env:IMDS_ENDPOINT
                $savedAzurePsHost = $env:AZUREPS_HOST_ENVIRONMENT
                $env:IDENTITY_ENDPOINT = ''
                $env:IMDS_ENDPOINT = ''
                $env:AZUREPS_HOST_ENVIRONMENT = ''
                try
                {
                    Connect-M365Tenant -Workload 'Fabric' -Identity -TenantId 'contoso.onmicrosoft.com'
                }
                finally
                {
                    $env:IDENTITY_ENDPOINT = $savedIdentityEndpoint
                    $env:IMDS_ENDPOINT = $savedImdsEndpoint
                    $env:AZUREPS_HOST_ENVIRONMENT = $savedAzurePsHost
                }

                $connection = Get-MSCloudLoginConnectionProfile -Workload 'Fabric'

                $connection.Connected | Should -BeTrue
                $connection.AuthenticationType | Should -Be 'Identity'
                $connection.AccessToken | Should -Be 'Bearer managed-identity-token'
                Should -Invoke Invoke-RestMethod -ParameterFilter {
                    $Uri -like 'http://169.254.169.254/metadata/identity/oauth2/token*resource=https://api.fabric.microsoft.com'
                }
            }
        }
    }

    Context 'When the workload does not support the requested authentication type' {
        It 'Should throw and leave the O365Portal profile disconnected' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Invoke-WebRequest -MockWith {
                    return @{ Content = '{ "token_endpoint": "https://login.microsoftonline.com/tenant/oauth2/v2.0/token" }' }
                }

                { Connect-M365Tenant -Workload 'O365Portal' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -ApplicationSecret 'super-secret' } |
                    Should -Throw "*'ServicePrincipalWithSecret' is not supported for workload 'O365Portal'*"

                (Get-MSCloudLoginConnectionProfile -Workload 'O365Portal').Connected | Should -BeFalse
            }
        }
    }

    Context 'When the same workload is connected more than once' {
        It 'Should reuse the existing session for identical parameters' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Invoke-WebRequest -MockWith {
                    return @{ Content = '{ "token_endpoint": "https://login.microsoftonline.com/tenant/oauth2/v2.0/token" }' }
                }
                Mock -CommandName Invoke-RestMethod -MockWith {
                    return @{ token_type = 'Bearer'; access_token = 'e2e-token' }
                }

                $parameters = @{
                    Workload          = 'AdminAPI'
                    ApplicationId     = '11111111-1111-1111-1111-111111111111'
                    TenantId          = 'contoso.onmicrosoft.com'
                    ApplicationSecret = 'super-secret'
                }
                Connect-M365Tenant @parameters
                Connect-M365Tenant @parameters

                Should -Invoke Invoke-RestMethod -Exactly 1
            }
        }

        It 'Should acquire a new token when the application secret was rotated' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Invoke-WebRequest -MockWith {
                    return @{ Content = '{ "token_endpoint": "https://login.microsoftonline.com/tenant/oauth2/v2.0/token" }' }
                }
                Mock -CommandName Invoke-RestMethod -MockWith {
                    return @{ token_type = 'Bearer'; access_token = 'e2e-token' }
                }

                Connect-M365Tenant -Workload 'AdminAPI' -ApplicationId '11111111-1111-1111-1111-111111111111' -TenantId 'contoso.onmicrosoft.com' -ApplicationSecret 'old-secret'
                Connect-M365Tenant -Workload 'AdminAPI' -ApplicationId '11111111-1111-1111-1111-111111111111' -TenantId 'contoso.onmicrosoft.com' -ApplicationSecret 'new-secret'

                Should -Invoke Invoke-RestMethod -Exactly 2
                Should -Invoke Invoke-RestMethod -Exactly 1 -ParameterFilter { $Body.client_secret -eq 'new-secret' }
            }
        }

        It 'Should reconnect when switching from an application secret to a managed identity' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Invoke-WebRequest -MockWith {
                    return @{ Content = '{ "token_endpoint": "https://login.microsoftonline.com/tenant/oauth2/v2.0/token" }' }
                }
                Mock -CommandName Invoke-RestMethod -MockWith {
                    return @{ token_type = 'Bearer'; access_token = 'e2e-token' }
                }

                Connect-M365Tenant -Workload 'AdminAPI' -ApplicationId '11111111-1111-1111-1111-111111111111' -TenantId 'contoso.onmicrosoft.com' -ApplicationSecret 'super-secret'
                (Get-MSCloudLoginConnectionProfile -Workload 'AdminAPI').AuthenticationType | Should -Be 'ServicePrincipalWithSecret'

                $savedIdentityEndpoint = $env:IDENTITY_ENDPOINT
                $savedImdsEndpoint = $env:IMDS_ENDPOINT
                $savedAzurePsHost = $env:AZUREPS_HOST_ENVIRONMENT
                $env:IDENTITY_ENDPOINT = ''
                $env:IMDS_ENDPOINT = ''
                $env:AZUREPS_HOST_ENVIRONMENT = ''
                try
                {
                    Connect-M365Tenant -Workload 'AdminAPI' -Identity -TenantId 'contoso.onmicrosoft.com'
                }
                finally
                {
                    $env:IDENTITY_ENDPOINT = $savedIdentityEndpoint
                    $env:IMDS_ENDPOINT = $savedImdsEndpoint
                    $env:AZUREPS_HOST_ENVIRONMENT = $savedAzurePsHost
                }

                (Get-MSCloudLoginConnectionProfile -Workload 'AdminAPI').AuthenticationType | Should -Be 'Identity'
                Should -Invoke Invoke-RestMethod -Exactly 2
            }
        }
    }

    Context 'When the SharePoint Online REST workload derives its URLs from the tenant name' {
        It 'Should build the admin URL and scope the token to the admin site' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Invoke-WebRequest -MockWith {
                    return @{ Content = '{ "token_endpoint": "https://login.microsoftonline.com/tenant/oauth2/v2.0/token" }' }
                }
                Mock -CommandName Invoke-RestMethod -MockWith {
                    return @{ token_type = 'Bearer'; access_token = 'e2e-token' }
                }

                Connect-M365Tenant -Workload 'SharePointOnlineREST' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -ApplicationSecret 'super-secret'

                $connection = Get-MSCloudLoginConnectionProfile -Workload 'SharePointOnlineREST'

                $connection.Connected | Should -BeTrue
                $connection.AdminUrl | Should -Be 'https://contoso-admin.sharepoint.com'
                $connection.HostUrl | Should -Be 'https://contoso-admin.sharepoint.com'
                $connection.Scope | Should -Be 'https://contoso-admin.sharepoint.com/.default'
            }
        }
    }

    Context 'When a REST workload is reset' {
        It 'Should clear the token and reconnect on the next call' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Invoke-WebRequest -MockWith {
                    return @{ Content = '{ "token_endpoint": "https://login.microsoftonline.com/tenant/oauth2/v2.0/token" }' }
                }
                Mock -CommandName Invoke-RestMethod -MockWith {
                    return @{ token_type = 'Bearer'; access_token = 'e2e-token' }
                }

                Connect-M365Tenant -Workload 'Fabric' -ApplicationId '11111111-1111-1111-1111-111111111111' -TenantId 'contoso.onmicrosoft.com' -ApplicationSecret 'super-secret'
                Reset-MSCloudLoginConnectionProfileContext -Workload 'Fabric'

                $connection = Get-MSCloudLoginConnectionProfile -Workload 'Fabric'
                $connection.Connected | Should -BeFalse
                $connection.AccessToken | Should -BeNullOrEmpty

                Connect-M365Tenant -Workload 'Fabric' -ApplicationId '11111111-1111-1111-1111-111111111111' -TenantId 'contoso.onmicrosoft.com' -ApplicationSecret 'super-secret'

                (Get-MSCloudLoginConnectionProfile -Workload 'Fabric').Connected | Should -BeTrue
                Should -Invoke Invoke-RestMethod -Exactly 2
            }
        }
    }
}

Describe 'Cloud environment detection during Workload.Setup()' {

    BeforeEach {
        InModuleScope 'MSCloudLoginAssistant' {
            $Script:MSCloudLoginConnectionProfile = $null
            $Script:MSCloudLoginTriedGetEnvironment = $false
            $Script:CloudEnvironmentInfo = $null
        }
    }

    Context 'When the openid configuration identifies the cloud' {
        It 'Should resolve <Description> to the <ExpectedEnvironment> environment' -TestCases @(
            @{ Description = 'a DODCON sub scope'; OpenIdConfiguration = '{ "tenant_region_sub_scope": "DODCON", "token_endpoint": "https://login.microsoftonline.us/t/oauth2/v2.0/token" }'; ExpectedEnvironment = 'AzureUSGovernment' }
            @{ Description = 'a GCC tenant without a sub scope'; OpenIdConfiguration = '{ "tenant_region_scope": "USGov", "token_endpoint": "https://login.microsoftonline.us/t/oauth2/v2.0/token" }'; ExpectedEnvironment = 'AzureUSGovernment' }
            @{ Description = 'the FG region scope'; OpenIdConfiguration = '{ "tenant_region_scope": "FG", "token_endpoint": "https://login.sovcloud-identity.fr/t/oauth2/v2.0/token" }'; ExpectedEnvironment = 'AzureFranceCloud' }
            @{ Description = 'the GG2 region scope'; OpenIdConfiguration = '{ "tenant_region_scope": "GG2", "token_endpoint": "https://login.sovcloud-identity.de/t/oauth2/v2.0/token" }'; ExpectedEnvironment = 'AzureGermanyCloud' }
            @{ Description = 'a worldwide tenant'; OpenIdConfiguration = '{ "tenant_region_scope": "EU", "token_endpoint": "https://login.microsoftonline.com/t/oauth2/v2.0/token" }'; ExpectedEnvironment = 'AzureCloud' }
        ) {
            param ($Description, $OpenIdConfiguration, $ExpectedEnvironment)
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{
                OpenIdConfiguration = $OpenIdConfiguration
                ExpectedEnvironment = $ExpectedEnvironment
            } {
                param ($OpenIdConfiguration, $ExpectedEnvironment)

                Mock -CommandName Invoke-WebRequest -MockWith { return @{ Content = $OpenIdConfiguration } }
                Mock -CommandName Invoke-RestMethod -MockWith {
                    return @{ token_type = 'Bearer'; access_token = 'e2e-token' }
                }

                Connect-M365Tenant -Workload 'AdminAPI' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -ApplicationSecret 'super-secret'

                (Get-MSCloudLoginConnectionProfile -Workload 'AdminAPI').EnvironmentName | Should -Be $ExpectedEnvironment
            }
        }

        It 'Should resolve a China token endpoint to AzureChinaCloud and capture the tenant GUID' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Invoke-WebRequest -MockWith {
                    return @{ Content = '{ "token_endpoint": "https://login.partner.microsoftonline.cn/22222222-2222-2222-2222-222222222222/oauth2/v2.0/token" }' }
                }
                Mock -CommandName Invoke-RestMethod -MockWith {
                    return @{ token_type = 'Bearer'; access_token = 'e2e-token' }
                }

                Connect-M365Tenant -Workload 'AdminAPI' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.partner.onmschina.cn' `
                    -ApplicationSecret 'super-secret'

                $connection = Get-MSCloudLoginConnectionProfile -Workload 'AdminAPI'
                $connection.EnvironmentName | Should -Be 'AzureChinaCloud'
                $connection.TenantGUID | Should -Be '22222222-2222-2222-2222-222222222222'
            }
        }

        It 'Should query the <Endpoint> discovery endpoint for a <TenantId> tenant' -TestCases @(
            @{ TenantId = 'contoso.onsovcloud.de'; Endpoint = 'login.sovcloud-identity.de' }
            @{ TenantId = 'contoso.onsovcloud.fr'; Endpoint = 'login.sovcloud-identity.fr' }
            @{ TenantId = 'contoso.onmicrosoft.com'; Endpoint = 'login.microsoftonline.com' }
        ) {
            param ($TenantId, $Endpoint)
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{ TenantId = $TenantId; Endpoint = $Endpoint } {
                param ($TenantId, $Endpoint)

                Mock -CommandName Invoke-WebRequest -MockWith {
                    return @{ Content = '{ "token_endpoint": "https://login.microsoftonline.com/t/oauth2/v2.0/token" }' }
                }
                Mock -CommandName Invoke-RestMethod -MockWith {
                    return @{ token_type = 'Bearer'; access_token = 'e2e-token' }
                }

                Connect-M365Tenant -Workload 'AdminAPI' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId $TenantId `
                    -ApplicationSecret 'super-secret'

                Should -Invoke Invoke-WebRequest -ParameterFilter {
                    $Uri -eq "https://$Endpoint/$TenantId/v2.0/.well-known/openid-configuration"
                }
            }
        }
    }

    Context 'When the cloud could not be determined' {
        It 'Should fall back to AzureCloud and skip a second discovery attempt' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginTriedGetEnvironment = $true
                Mock -CommandName Invoke-WebRequest -MockWith { throw 'discovery must not be attempted again' }
                Mock -CommandName Invoke-RestMethod -MockWith {
                    return @{ token_type = 'Bearer'; access_token = 'e2e-token' }
                }

                Connect-M365Tenant -Workload 'AdminAPI' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -ApplicationSecret 'super-secret'

                (Get-MSCloudLoginConnectionProfile -Workload 'AdminAPI').EnvironmentName | Should -Be 'AzureCloud'
                Should -Invoke Invoke-WebRequest -Exactly 0
            }
        }

        It 'Should fall back to AzureChinaCloud for a tenant id that ends in .cn' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginTriedGetEnvironment = $true
                Mock -CommandName Invoke-WebRequest -MockWith { throw 'discovery must not be attempted again' }
                Mock -CommandName Invoke-RestMethod -MockWith {
                    return @{ token_type = 'Bearer'; access_token = 'e2e-token' }
                }

                Connect-M365Tenant -Workload 'AdminAPI' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.partner.onmschina.cn' `
                    -ApplicationSecret 'super-secret'

                (Get-MSCloudLoginConnectionProfile -Workload 'AdminAPI').EnvironmentName | Should -Be 'AzureChinaCloud'
            }
        }
    }
}

Describe 'Connect-M365Tenant end-to-end with a custom environment' {

    BeforeAll {
        $script:customEnvironmentFile = 'CustomEnvironment.Tests.psd1'
        $script:customEnvironmentPath = Join-Path $script:moduleRoot $script:customEnvironmentFile
        @'
@{
    CustomEnvironment = $true
    CustomAdminApiScope = "https://admin.contoso.local/.default"
    CustomAdminApiAuthorizationUrl = "https://login.contoso.local"
    CustomTasksHostUrl = "https://tasks.contoso.local"
    CustomTasksScope = "https://tasks.contoso.local/.default"
    CustomTasksAuthorizationUrl = "https://login.contoso.local"
    CustomTasksResourceUrl = "https://tasks.contoso.local"
}
'@ | Set-Content -Path $script:customEnvironmentPath -Encoding utf8
    }

    AfterAll {
        Remove-Item -Path $script:customEnvironmentPath -Force -ErrorAction SilentlyContinue
        InModuleScope 'MSCloudLoginAssistant' -Parameters @{ ModuleRoot = $script:moduleRoot } {
            param ($ModuleRoot)
            $Script:CustomEnvConfig = Import-PowerShellDataFile -Path (Join-Path $ModuleRoot 'CustomEnvironment.psd1')
            $Script:LoadedCustomEnvFileName = 'CustomEnvironment.psd1'
        }
    }

    BeforeEach {
        InModuleScope 'MSCloudLoginAssistant' {
            $Script:MSCloudLoginConnectionProfile = $null
            $Script:MSCloudLoginTriedGetEnvironment = $false
            $Script:CloudEnvironmentInfo = $null
        }
    }

    It 'Should resolve the endpoints of <Workload> from the custom environment file' -TestCases @(
        @{ Workload = 'AdminAPI'; ExpectedScope = 'https://admin.contoso.local/.default' }
        @{ Workload = 'Tasks'; ExpectedScope = 'https://tasks.contoso.local/.default' }
    ) {
        param ($Workload, $ExpectedScope)
        InModuleScope 'MSCloudLoginAssistant' -Parameters @{
            Workload              = $Workload
            ExpectedScope         = $ExpectedScope
            CustomEnvironmentFile = $script:customEnvironmentFile
        } {
            param ($Workload, $ExpectedScope, $CustomEnvironmentFile)

            Mock -CommandName Invoke-WebRequest -MockWith {
                return @{ Content = '{ "token_endpoint": "https://login.contoso.local/t/oauth2/v2.0/token" }' }
            }
            Mock -CommandName Invoke-RestMethod -MockWith {
                return @{ token_type = 'Bearer'; access_token = 'custom-token' }
            }

            Connect-M365Tenant -Workload $Workload `
                -ApplicationId '11111111-1111-1111-1111-111111111111' `
                -TenantId 'contoso.onmicrosoft.com' `
                -ApplicationSecret 'super-secret' `
                -CustomEnvironmentFileName $CustomEnvironmentFile

            $connection = Get-MSCloudLoginConnectionProfile -Workload $Workload

            $connection.EnvironmentName | Should -Be 'Custom'
            $connection.AuthorizationUrl | Should -Be 'https://login.contoso.local'
            $connection.Scope | Should -Be $ExpectedScope
            Should -Invoke Invoke-RestMethod -ParameterFilter {
                $Uri -eq 'https://login.contoso.local/contoso.onmicrosoft.com/oauth2/v2.0/token'
            }
        }
    }

    It 'Should throw when the requested custom environment file does not exist' {
        InModuleScope 'MSCloudLoginAssistant' {
            { Connect-M365Tenant -Workload 'AdminAPI' `
                -ApplicationId '11111111-1111-1111-1111-111111111111' `
                -TenantId 'contoso.onmicrosoft.com' `
                -ApplicationSecret 'super-secret' `
                -CustomEnvironmentFileName 'DoesNotExist.psd1' } | Should -Throw
        }
    }
}
