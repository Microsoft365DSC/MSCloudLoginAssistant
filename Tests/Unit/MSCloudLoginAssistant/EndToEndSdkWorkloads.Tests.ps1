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

Describe 'Connect-M365Tenant end-to-end for PnP' {

    BeforeEach {
        InModuleScope 'MSCloudLoginAssistant' {
            $Script:MSCloudLoginConnectionProfile = $null
            $Script:MSCloudLoginTriedGetEnvironment = $true
            $Script:CloudEnvironmentInfo = $null
        }
    }

    Context 'When connecting with a service principal' {
        It 'Should hand the certificate thumbprint and the production Azure environment to Connect-PnPOnline' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Get-PnPContext -MockWith { return @{ Url = 'https://contoso-admin.sharepoint.com' } }

                Connect-M365Tenant -Workload 'PnP' `
                    -Url 'https://contoso-admin.sharepoint.com' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -CertificateThumbprint 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD'

                $connection = Get-MSCloudLoginConnectionProfile -Workload 'PnP'

                $connection.Connected | Should -BeTrue
                $connection.ConnectionUrl | Should -Be 'https://contoso-admin.sharepoint.com'
                $connection.AdminUrl | Should -Be 'https://contoso-admin.sharepoint.com'
                $connection.PnPAzureEnvironment | Should -Be 'Production'

                Should -Invoke Connect-PnPOnline -Exactly 1 -ParameterFilter {
                    $Url -eq 'https://contoso-admin.sharepoint.com' -and
                    $ClientId -eq '11111111-1111-1111-1111-111111111111' -and
                    $Tenant -eq 'contoso.onmicrosoft.com' -and
                    $Thumbprint -eq 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD' -and
                    $AzureEnvironment -eq 'Production'
                }
            }
        }

        It 'Should hand the application secret to Connect-PnPOnline' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Get-PnPContext -MockWith { return @{ Url = 'https://contoso-admin.sharepoint.com' } }

                Connect-M365Tenant -Workload 'PnP' `
                    -Url 'https://contoso-admin.sharepoint.com' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -ApplicationSecret 'super-secret'

                (Get-MSCloudLoginConnectionProfile -Workload 'PnP').Connected | Should -BeTrue
                Should -Invoke Connect-PnPOnline -Exactly 1 -ParameterFilter {
                    $ClientSecret -eq 'super-secret' -and $AzureEnvironment -eq 'Production'
                }
            }
        }

        It 'Should hand the certificate path and password to Connect-PnPOnline' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Get-PnPContext -MockWith { return @{ Url = 'https://contoso-admin.sharepoint.com' } }

                Connect-M365Tenant -Workload 'PnP' `
                    -Url 'https://contoso-admin.sharepoint.com' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -CertificatePath 'C:\certificates\contoso.pfx' `
                    -CertificatePassword (ConvertTo-SecureString 'certificate-password' -AsPlainText -Force)

                (Get-MSCloudLoginConnectionProfile -Workload 'PnP').AuthenticationType | Should -Be 'ServicePrincipalWithPath'
                Should -Invoke Connect-PnPOnline -Exactly 1 -ParameterFilter {
                    $CertificatePath -eq 'C:\certificates\contoso.pfx'
                }
            }
        }
    }

    Context 'When the SharePoint URLs have to be derived' {
        It 'Should derive the admin and connection URL from the tenant name when no URL is passed' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Get-PnPContext -MockWith { throw 'no context yet' }

                Connect-M365Tenant -Workload 'PnP' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -CertificateThumbprint 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD'

                $connection = Get-MSCloudLoginConnectionProfile -Workload 'PnP'
                $connection.AdminUrl | Should -Be 'https://contoso-admin.sharepoint.com'
                $connection.ConnectionUrl | Should -Be 'https://contoso.sharepoint.com'
            }
        }

        It 'Should ask Microsoft Graph for the admin URL when connecting with credentials only' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Get-PnPContext -MockWith { throw 'no context yet' }
                Mock -CommandName Invoke-MgGraphRequest -MockWith { return @{ webUrl = 'https://contoso.sharepoint.com' } }

                Connect-M365Tenant -Workload 'PnP' `
                    -Credential (New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force)))

                $connection = Get-MSCloudLoginConnectionProfile -Workload 'PnP'
                $connection.AuthenticationType | Should -Be 'Credentials'
                $connection.AdminUrl | Should -Be 'https://contoso-admin.sharepoint.com'
                Should -Invoke Connect-PnPOnline -Exactly 1 -ParameterFilter {
                    $Url -eq 'https://contoso-admin.sharepoint.com' -and
                    $ClientId -eq '9bc3ab49-b65d-410a-85ad-de819febfddc'
                }
            }
        }
    }

    Context 'When connecting with delegated authentication' {
        It 'Should pass the credential together with the application id to Connect-PnPOnline' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Get-PnPContext -MockWith { throw 'no context yet' }

                Connect-M365Tenant -Workload 'PnP' `
                    -Url 'https://contoso-admin.sharepoint.com' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -Credential (New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force)))

                (Get-MSCloudLoginConnectionProfile -Workload 'PnP').AuthenticationType | Should -Be 'CredentialsWithApplicationId'
                Should -Invoke Connect-PnPOnline -Exactly 1 -ParameterFilter {
                    $ClientId -eq '11111111-1111-1111-1111-111111111111' -and $null -ne $Credentials
                }
            }
        }

        It 'Should reject a credential that is combined with a tenant id' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Get-PnPContext -MockWith { throw 'no context yet' }

                { Connect-M365Tenant -Workload 'PnP' `
                    -Url 'https://contoso-admin.sharepoint.com' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -Credential (New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))) } |
                    Should -Throw '*cannot specify TenantId with Credentials*'

                (Get-MSCloudLoginConnectionProfile -Workload 'PnP').Connected | Should -BeFalse
            }
        }
    }

    Context 'When connecting with a token' {
        It 'Should pass a supplied access token to Connect-PnPOnline' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Get-PnPContext -MockWith { throw 'no context yet' }

                Connect-M365Tenant -Workload 'PnP' `
                    -Url 'https://contoso-admin.sharepoint.com' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -AccessTokens @('caller-supplied-token')

                (Get-MSCloudLoginConnectionProfile -Workload 'PnP').Connected | Should -BeTrue
                Should -Invoke Connect-PnPOnline -Exactly 1 -ParameterFilter {
                    $AccessToken -eq 'caller-supplied-token'
                }
            }
        }

        It 'Should request a managed identity token for the connection URL' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Get-PnPContext -MockWith { throw 'no context yet' }
                Mock -CommandName Get-AuthToken -MockWith { return 'managed-identity-token' }

                Connect-M365Tenant -Workload 'PnP' `
                    -Url 'https://contoso-admin.sharepoint.com' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -Identity

                (Get-MSCloudLoginConnectionProfile -Workload 'PnP').Connected | Should -BeTrue
                Should -Invoke Get-AuthToken -Exactly 1 -ParameterFilter {
                    $Resource -eq 'https://contoso-admin.sharepoint.com' -and $Identity.IsPresent
                }
                Should -Invoke Connect-PnPOnline -Exactly 1 -ParameterFilter {
                    $AccessToken -eq 'managed-identity-token'
                }
            }
        }
    }

    Context 'When the tenant lives in a sovereign cloud' {
        It 'Should translate <EnvironmentName> into the <ExpectedPnPEnvironment> PnP environment' -TestCases @(
            @{ EnvironmentName = 'AzureUSGovernment'; ExpectedPnPEnvironment = 'USGovernmentHigh'; Discovery = '{ "tenant_region_sub_scope": "USGov", "token_endpoint": "https://login.microsoftonline.us/t/oauth2/v2.0/token" }' }
            @{ EnvironmentName = 'AzureDOD'; ExpectedPnPEnvironment = 'USGovernmentDoD'; Discovery = '{ "tenant_region_sub_scope": "DOD", "token_endpoint": "https://login.microsoftonline.us/t/oauth2/v2.0/token" }' }
            @{ EnvironmentName = 'AzureChinaCloud'; ExpectedPnPEnvironment = 'China'; Discovery = '{ "token_endpoint": "https://login.partner.microsoftonline.cn/22222222-2222-2222-2222-222222222222/oauth2/v2.0/token" }' }
        ) {
            param ($EnvironmentName, $ExpectedPnPEnvironment, $Discovery)
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{
                EnvironmentName        = $EnvironmentName
                ExpectedPnPEnvironment = $ExpectedPnPEnvironment
                Discovery              = $Discovery
            } {
                param ($EnvironmentName, $ExpectedPnPEnvironment, $Discovery)

                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Get-PnPContext -MockWith { throw 'no context yet' }

                $Script:CloudEnvironmentInfo = ConvertFrom-Json $Discovery

                Connect-M365Tenant -Workload 'PnP' `
                    -Url 'https://contoso-admin.sharepoint.com' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -CertificateThumbprint 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD'

                $connection = Get-MSCloudLoginConnectionProfile -Workload 'PnP'
                $connection.EnvironmentName | Should -Be $EnvironmentName
                $connection.PnPAzureEnvironment | Should -Be $ExpectedPnPEnvironment
                Should -Invoke Connect-PnPOnline -Exactly 1 -ParameterFilter {
                    $AzureEnvironment -eq $ExpectedPnPEnvironment
                }
            }
        }
    }

    Context 'When the connection fails' {
        It 'Should surface the underlying error and leave the profile disconnected' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Get-PnPContext -MockWith { throw 'no context yet' }
                Mock -CommandName Connect-PnPOnline -MockWith { throw 'The remote server returned an error: (403) Forbidden.' }

                { Connect-M365Tenant -Workload 'PnP' `
                    -Url 'https://contoso-admin.sharepoint.com' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -CertificateThumbprint 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD' } |
                    Should -Throw '*403*'

                (Get-MSCloudLoginConnectionProfile -Workload 'PnP').Connected | Should -BeFalse
            }
        }

        It 'Should ask for management shell consent when the application was never consented to' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Get-PnPContext -MockWith { throw 'no context yet' }
                Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $false }
                Mock -CommandName Register-PnPManagementShellAccess -MockWith { throw 'consent was declined' }
                Mock -CommandName Connect-PnPOnline -MockWith {
                    throw 'AADSTS65001: The user or administrator has not consented to use the application with ID 11111111-1111-1111-1111-111111111111'
                }

                { Connect-M365Tenant -Workload 'PnP' `
                    -Url 'https://contoso-admin.sharepoint.com' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -CertificateThumbprint 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD' } |
                    Should -Throw "*Register-PnPManagementShellAccess*"
            }
        }

        It 'Should retry interactively when the sign-in requires multi-factor authentication' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Get-PnPContext -MockWith { throw 'no context yet' }
                Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $false }
                Mock -CommandName Connect-PnPOnline -MockWith {
                    if (-not $Interactive.IsPresent)
                    {
                        throw 'AADSTS50076: Due to a configuration change made by your administrator you must use multi-factor authentication.'
                    }
                }

                Connect-M365Tenant -Workload 'PnP' `
                    -Url 'https://contoso-admin.sharepoint.com' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -Credential (New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force)))

                $connection = Get-MSCloudLoginConnectionProfile -Workload 'PnP'
                $connection.Connected | Should -BeTrue
                $connection.MultiFactorAuthentication | Should -BeTrue
                Should -Invoke Connect-PnPOnline -Exactly 1 -ParameterFilter { $Interactive.IsPresent }
            }
        }
    }

    Context 'When PnP is reset' {
        It 'Should disconnect the PnP session and clear the connection state' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Disconnect-PnPOnline -MockWith { }
                Mock -CommandName Get-PnPContext -MockWith { throw 'no context yet' }

                Connect-M365Tenant -Workload 'PnP' `
                    -Url 'https://contoso-admin.sharepoint.com' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -CertificateThumbprint 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD'

                Reset-MSCloudLoginConnectionProfileContext -Workload 'PnP'

                Should -Invoke Disconnect-PnPOnline -Exactly 1
                (Get-MSCloudLoginConnectionProfile -Workload 'PnP').Connected | Should -BeFalse
            }
        }
    }
}

Describe 'Connect-M365Tenant end-to-end for Microsoft Graph' {

    BeforeEach {
        InModuleScope 'MSCloudLoginAssistant' {
            $Script:MSCloudLoginConnectionProfile = $null
            $Script:MSCloudLoginTriedGetEnvironment = $true
            $Script:CloudEnvironmentInfo = $null
        }
    }

    Context 'When connecting with a service principal' {
        It 'Should pass the certificate to Connect-MgGraph and expose the Global environment' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MgGraph -MockWith { }
                Mock -CommandName Get-MgContext -MockWith { return $null }
                Mock -CommandName Get-MSCloudLoginCertificate -MockWith {
                    return New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
                }

                Connect-M365Tenant -Workload 'MicrosoftGraph' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -CertificateThumbprint 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD'

                $connection = Get-MSCloudLoginConnectionProfile -Workload 'MicrosoftGraph'

                $connection.Connected | Should -BeTrue
                $connection.GraphEnvironment | Should -Be 'Global'
                $connection.ResourceUrl | Should -Be 'https://graph.microsoft.com/'
                $connection.TokenUrl | Should -Be 'https://login.microsoftonline.com/contoso.onmicrosoft.com/oauth2/v2.0/token'

                Should -Invoke Connect-MgGraph -Exactly 1 -ParameterFilter {
                    $ClientId -eq '11111111-1111-1111-1111-111111111111' -and
                    $TenantId -eq 'contoso.onmicrosoft.com' -and
                    $Environment -eq 'Global' -and
                    $null -ne $Certificate
                }
            }
        }

        It 'Should build a client secret credential for Connect-MgGraph' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MgGraph -MockWith { }
                Mock -CommandName Get-MgContext -MockWith { return $null }

                Connect-M365Tenant -Workload 'MicrosoftGraph' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -ApplicationSecret 'super-secret'

                (Get-MSCloudLoginConnectionProfile -Workload 'MicrosoftGraph').Connected | Should -BeTrue
                Should -Invoke Connect-MgGraph -Exactly 1 -ParameterFilter {
                    $ClientSecretCredential.UserName -eq '11111111-1111-1111-1111-111111111111' -and
                    $ClientSecretCredential.GetNetworkCredential().Password -eq 'super-secret'
                }
            }
        }

        It 'Should load the certificate from disk for a certificate path connection' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MgGraph -MockWith { }
                Mock -CommandName Get-MgContext -MockWith { return $null }
                Mock -CommandName Get-MSCloudLoginCertificate -MockWith {
                    return New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
                }

                Connect-M365Tenant -Workload 'MicrosoftGraph' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -CertificatePath 'C:\certificates\contoso.pfx' `
                    -CertificatePassword (ConvertTo-SecureString 'certificate-password' -AsPlainText -Force)

                (Get-MSCloudLoginConnectionProfile -Workload 'MicrosoftGraph').Connected | Should -BeTrue
                Should -Invoke Get-MSCloudLoginCertificate -Exactly 1 -ParameterFilter {
                    $CertificatePath -eq 'C:\certificates\contoso.pfx'
                }
            }
        }
    }

    Context 'When the tenant lives in a sovereign cloud' {
        It 'Should map <EnvironmentName> to the <ExpectedGraphEnvironment> Graph environment and resource URL <ExpectedResourceUrl>' -TestCases @(
            @{ EnvironmentName = 'AzureUSGovernment'; ExpectedGraphEnvironment = 'USGov'; ExpectedResourceUrl = 'https://graph.microsoft.us/'; Discovery = '{ "tenant_region_sub_scope": "USGov", "token_endpoint": "https://login.microsoftonline.us/t/oauth2/v2.0/token" }' }
            @{ EnvironmentName = 'AzureDOD'; ExpectedGraphEnvironment = 'USGovDoD'; ExpectedResourceUrl = 'https://dod-graph.microsoft.us/'; Discovery = '{ "tenant_region_sub_scope": "DOD", "token_endpoint": "https://login.microsoftonline.us/t/oauth2/v2.0/token" }' }
            @{ EnvironmentName = 'AzureChinaCloud'; ExpectedGraphEnvironment = 'China'; ExpectedResourceUrl = 'https://microsoftgraph.chinacloudapi.cn/'; Discovery = '{ "token_endpoint": "https://login.partner.microsoftonline.cn/22222222-2222-2222-2222-222222222222/oauth2/v2.0/token" }' }
            @{ EnvironmentName = 'AzureFranceCloud'; ExpectedGraphEnvironment = 'BleuCloud'; ExpectedResourceUrl = 'https://graph.svc.sovcloud.fr/'; Discovery = '{ "tenant_region_scope": "FG", "token_endpoint": "https://login.sovcloud-identity.fr/t/oauth2/v2.0/token" }' }
            @{ EnvironmentName = 'AzureGermanyCloud'; ExpectedGraphEnvironment = 'DelosCloud'; ExpectedResourceUrl = 'https://graph.svc.sovcloud.de/'; Discovery = '{ "tenant_region_scope": "GG2", "token_endpoint": "https://login.sovcloud-identity.de/t/oauth2/v2.0/token" }' }
        ) {
            param ($EnvironmentName, $ExpectedGraphEnvironment, $ExpectedResourceUrl, $Discovery)
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{
                EnvironmentName          = $EnvironmentName
                ExpectedGraphEnvironment = $ExpectedGraphEnvironment
                ExpectedResourceUrl      = $ExpectedResourceUrl
                Discovery                = $Discovery
            } {
                param ($EnvironmentName, $ExpectedGraphEnvironment, $ExpectedResourceUrl, $Discovery)

                Mock -CommandName Connect-MgGraph -MockWith { }
                Mock -CommandName Get-MgContext -MockWith { return $null }

                $Script:CloudEnvironmentInfo = ConvertFrom-Json $Discovery

                Connect-M365Tenant -Workload 'MicrosoftGraph' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -ApplicationSecret 'super-secret'

                $connection = Get-MSCloudLoginConnectionProfile -Workload 'MicrosoftGraph'
                $connection.EnvironmentName | Should -Be $EnvironmentName
                $connection.GraphEnvironment | Should -Be $ExpectedGraphEnvironment
                $connection.ResourceUrl | Should -Be $ExpectedResourceUrl
                $connection.Scope | Should -Be "$ExpectedResourceUrl.default"
            }
        }
    }

    Context 'When connecting with a token' {
        It 'Should pass a supplied access token to Connect-MgGraph as a secure string' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MgGraph -MockWith { }
                Mock -CommandName Get-MgContext -MockWith { return $null }

                Connect-M365Tenant -Workload 'MicrosoftGraph' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -AccessTokens @('caller-supplied-token')

                (Get-MSCloudLoginConnectionProfile -Workload 'MicrosoftGraph').Connected | Should -BeTrue
                Should -Invoke Connect-MgGraph -Exactly 1 -ParameterFilter {
                    $AccessToken -is [System.Security.SecureString]
                }
            }
        }

        It 'Should exchange a managed identity token and adopt the tenant id reported by the context' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MgGraph -MockWith { }
                Mock -CommandName Get-AuthToken -MockWith { return 'managed-identity-token' }
                Mock -CommandName Get-MgContext -MockWith { return @{ TenantId = '33333333-3333-3333-3333-333333333333' } }

                Connect-M365Tenant -Workload 'MicrosoftGraph' -Identity -TenantId 'contoso.onmicrosoft.com'

                $connection = Get-MSCloudLoginConnectionProfile -Workload 'MicrosoftGraph'
                $connection.Connected | Should -BeTrue
                $connection.TenantId | Should -Be '33333333-3333-3333-3333-333333333333'
                Should -Invoke Get-AuthToken -Exactly 1 -ParameterFilter {
                    $Resource -eq 'https://graph.microsoft.com' -and $Identity.IsPresent
                }
            }
        }
    }

    Context 'When connecting with user credentials' {
        It 'Should acquire a delegated token and connect with the default Graph PowerShell application id' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MgGraph -MockWith { }
                Mock -CommandName Disconnect-MgGraph -MockWith { }
                Mock -CommandName Get-MgContext -MockWith { return $null }
                Mock -CommandName Get-AuthToken -MockWith { return @{ access_token = 'delegated-token' } }

                Connect-M365Tenant -Workload 'MicrosoftGraph' `
                    -Credential (New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force)))

                $connection = Get-MSCloudLoginConnectionProfile -Workload 'MicrosoftGraph'
                $connection.Connected | Should -BeTrue
                $connection.ApplicationId | Should -Be '14d82eec-204b-4c2f-b7e8-296a70dab67e'
                $connection.AccessTokens | Should -Be @('delegated-token')
                Should -Invoke Get-AuthToken -Exactly 1 -ParameterFilter {
                    $TenantId -eq 'contoso.onmicrosoft.com'
                }
            }
        }

        It 'Should switch to the device code flow when the account requires multi-factor authentication' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MgGraph -MockWith { }
                Mock -CommandName Disconnect-MgGraph -MockWith { }
                Mock -CommandName Get-MgContext -MockWith { return $null }
                Mock -CommandName Get-AuthToken -MockWith {
                    if (-not $DeviceCode.IsPresent)
                    {
                        throw 'AADSTS50076: Due to a configuration change made by your administrator you must use multi-factor authentication.'
                    }
                    return @{ access_token = 'device-code-token' }
                }

                Connect-M365Tenant -Workload 'MicrosoftGraph' `
                    -Credential (New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force)))

                $connection = Get-MSCloudLoginConnectionProfile -Workload 'MicrosoftGraph'
                $connection.Connected | Should -BeTrue
                $connection.MultiFactorAuthentication | Should -BeTrue
                $connection.AccessTokens | Should -Be @('device-code-token')
            }
        }
    }

    Context 'When an existing Graph context is still valid' {
        It 'Should reuse the session instead of connecting again' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MgGraph -MockWith { }
                Mock -CommandName Get-MgContext -MockWith { return @{ TenantId = 'contoso.onmicrosoft.com' } }

                $parameters = @{
                    Workload          = 'MicrosoftGraph'
                    ApplicationId     = '11111111-1111-1111-1111-111111111111'
                    TenantId          = 'contoso.onmicrosoft.com'
                    ApplicationSecret = 'super-secret'
                }
                Connect-M365Tenant @parameters
                Connect-M365Tenant @parameters

                Should -Invoke Connect-MgGraph -Exactly 1
            }
        }
    }

    Context 'When Microsoft Graph is reset' {
        It 'Should disconnect the Graph session and clear the connection state' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MgGraph -MockWith { }
                Mock -CommandName Disconnect-MgGraph -MockWith { }
                Mock -CommandName Get-MgContext -MockWith { return $null }

                Connect-M365Tenant -Workload 'MicrosoftGraph' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -ApplicationSecret 'super-secret'

                Reset-MSCloudLoginConnectionProfileContext -Workload 'MicrosoftGraph'

                Should -Invoke Disconnect-MgGraph -Exactly 1
                (Get-MSCloudLoginConnectionProfile -Workload 'MicrosoftGraph').Connected | Should -BeFalse
            }
        }
    }
}
