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

    $script:customEnvironmentFile = 'CustomEnvironment.Sovereign.Tests.psd1'
    $script:customEnvironmentPath = Join-Path $script:moduleRoot $script:customEnvironmentFile
    @'
@{
    CustomEnvironment = $true

    CustomGraphAuthorizationUrl = "https://login.contoso.local"
    CustomGraphResourceUrl = "https://graph.contoso.local/"
    CustomGraphScope = "https://graph.contoso.local/.default"
    CustomGraphTokenUrl = "https://login.contoso.local"

    CustomPnPScope = "https://sharepoint.contoso.local/.default"
    CustomPnPTokenUrl = "https://login.contoso.local"

    CustomEXOConnectionUri = "https://outlook.contoso.local/powershell-liveid/"
    CustomEXOAzureADAuthorizationEndpointUri = "https://login.contoso.local/common"

    CustomSCCConnectionUrl = "https://compliance.contoso.local/powershell-liveid/"
    CustomSCCAuthorizationUrl = "https://login.contoso.local/organizations"
    CustomSCCAzureADAuthorizationEndpointUri = "https://login.contoso.local/common"

    CustomPowerPlatformRESTScope = "https://powerapps.contoso.local/.default"
    CustomPowerPlatformRESTAuthorizationUrl = "https://login.contoso.local"
    CustomPowerPlatformRESTAudience = "https://powerapps.contoso.local/"
    CustomPowerPlatformRESTClientId = "99999999-9999-9999-9999-999999999999"
    CustomPowerPlatformRESTBapEndpoint = "api.bap.contoso.local"

    CustomTeamsTokenUrl = "https://login.contoso.local"
    CustomTeamsScope = "https://teams.contoso.local/.default"
    CustomTeamsEndpoints = @{
        ActiveDirectory = "https://login.contoso.local"
        MsGraphEndpointResourceId = "https://graph.contoso.local/"
        TeamsConfigApiEndpoint = "https://config.teams.contoso.local"
    }
}
'@ | Set-Content -Path $script:customEnvironmentPath -Encoding utf8
}

AfterAll {
    Remove-Item -Path $script:customEnvironmentPath -Force -ErrorAction SilentlyContinue
    if ($script:tempModuleBase -and (Test-Path $script:tempModuleBase))
    {
        Remove-Item -Path $script:tempModuleBase -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Describe 'Connect-M365Tenant in a custom environment' {

    BeforeEach {
        InModuleScope 'MSCloudLoginAssistant' {
            $Script:MSCloudLoginConnectionProfile = $null
            $Script:MSCloudLoginTriedGetEnvironment = $true
            $Script:CloudEnvironmentInfo = $null
        }
    }

    AfterEach {
        InModuleScope 'MSCloudLoginAssistant' -Parameters @{ ModuleRoot = $script:moduleRoot } {
            param ($ModuleRoot)
            $Script:CustomEnvConfig = Import-PowerShellDataFile -Path (Join-Path $ModuleRoot 'CustomEnvironment.psd1')
            $Script:LoadedCustomEnvFileName = 'CustomEnvironment.psd1'
        }
    }

    Context 'Microsoft Graph' {
        It 'Should register the custom Graph environment and connect with a locally issued token' {
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{ CustomEnvironmentFile = $script:customEnvironmentFile } {
                param ($CustomEnvironmentFile)

                Mock -CommandName Get-MgContext -MockWith { return $null }
                Mock -CommandName Get-MgEnvironment -MockWith { return $null }
                Mock -CommandName Add-MgEnvironment -MockWith { }
                Mock -CommandName Connect-MgGraph -MockWith { }
                Mock -CommandName Get-AuthToken -MockWith { return @{ access_token = 'custom-graph-token' } }

                Connect-M365Tenant -Workload 'MicrosoftGraph' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.local' `
                    -CertificateThumbprint 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD' `
                    -CustomEnvironmentFileName $CustomEnvironmentFile

                $connection = Get-MSCloudLoginConnectionProfile -Workload 'MicrosoftGraph'
                $connection.Connected | Should -BeTrue
                $connection.EnvironmentName | Should -Be 'Custom'
                $connection.GraphEnvironment | Should -Be 'Custom'
                $connection.ResourceUrl | Should -Be 'https://graph.contoso.local/'
                $connection.TokenUrl | Should -Be 'https://login.contoso.local/contoso.local/oauth2/v2.0/token'

                Should -Invoke Add-MgEnvironment -Exactly 1 -ParameterFilter {
                    $Name -eq 'Custom' -and $GraphEndpoint -eq 'https://graph.contoso.local/'
                }
                Should -Invoke Connect-MgGraph -Exactly 1 -ParameterFilter {
                    $AccessToken -is [System.Security.SecureString] -and $Environment -eq 'Custom'
                }
            }
        }

        It 'Should not register the custom environment twice' {
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{ CustomEnvironmentFile = $script:customEnvironmentFile } {
                param ($CustomEnvironmentFile)

                Mock -CommandName Get-MgContext -MockWith { return $null }
                Mock -CommandName Get-MgEnvironment -MockWith { return @([PSCustomObject]@{ Name = 'Custom' }) }
                Mock -CommandName Add-MgEnvironment -MockWith { }
                Mock -CommandName Connect-MgGraph -MockWith { }
                Mock -CommandName Get-AuthToken -MockWith { return @{ access_token = 'custom-graph-token' } }

                Connect-M365Tenant -Workload 'MicrosoftGraph' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.local' `
                    -CertificateThumbprint 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD' `
                    -CustomEnvironmentFileName $CustomEnvironmentFile

                Should -Invoke Add-MgEnvironment -Exactly 0
            }
        }
    }

    Context 'PnP' {
        It 'Should acquire a token from the custom token endpoint and pass it to Connect-PnPOnline' {
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{ CustomEnvironmentFile = $script:customEnvironmentFile } {
                param ($CustomEnvironmentFile)

                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Get-PnPContext -MockWith { throw 'no context yet' }
                Mock -CommandName Invoke-RestMethod -MockWith { return @{ access_token = 'custom-pnp-token' } }
                Mock -CommandName Get-MSCloudLoginCertificate -MockWith {
                    $rsa = [System.Security.Cryptography.RSA]::Create(2048)
                    $request = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new(
                        [System.Security.Cryptography.X509Certificates.X500DistinguishedName]::new('CN=MSCloudLoginAssistantCustomPnP'),
                        $rsa,
                        [System.Security.Cryptography.HashAlgorithmName]::SHA256,
                        [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
                    return $request.CreateSelfSigned([System.DateTimeOffset]::Now.AddDays(-1), [System.DateTimeOffset]::Now.AddDays(1))
                }

                Connect-M365Tenant -Workload 'PnP' `
                    -Url 'https://contoso-admin.sharepoint.contoso.local' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.local' `
                    -CertificateThumbprint 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD' `
                    -CustomEnvironmentFileName $CustomEnvironmentFile

                $connection = Get-MSCloudLoginConnectionProfile -Workload 'PnP'
                $connection.Connected | Should -BeTrue
                $connection.PnPAzureEnvironment | Should -Be 'Custom'
                $connection.AuthorizationUrl | Should -Be 'https://login.contoso.local'
                $connection.Scope | Should -Be 'https://sharepoint.contoso.local/.default'
                $connection.TokenUrl | Should -Be 'https://login.contoso.local/contoso.local/oauth2/v2.0/token'

                Should -Invoke Invoke-RestMethod -Exactly 1 -ParameterFilter {
                    $Uri -eq 'https://login.contoso.local/contoso.local/oauth2/v2.0/token'
                }
                Should -Invoke Connect-PnPOnline -Exactly 1 -ParameterFilter {
                    $AccessToken -eq 'custom-pnp-token'
                }
            }
        }
    }

    Context 'Exchange Online' {
        It 'Should connect through the custom endpoint URIs' {
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{ CustomEnvironmentFile = $script:customEnvironmentFile } {
                param ($CustomEnvironmentFile)

                Mock -CommandName Remove-MSCloudLoginProxyModule -MockWith { }
                Mock -CommandName Get-ConnectionInformation -MockWith { return @() }
                Mock -CommandName Disconnect-ExchangeOnline -MockWith { }
                Mock -CommandName Connect-ExchangeOnline -MockWith { }

                Connect-M365Tenant -Workload 'ExchangeOnline' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.local' `
                    -CertificateThumbprint 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD' `
                    -CustomEnvironmentFileName $CustomEnvironmentFile

                $connection = Get-MSCloudLoginConnectionProfile -Workload 'ExchangeOnline'
                $connection.ExchangeEnvironmentName | Should -Be 'Custom'
                $connection.ConnectionUri | Should -Be 'https://outlook.contoso.local/powershell-liveid/'

                Should -Invoke Connect-ExchangeOnline -Exactly 1 -ParameterFilter {
                    $ConnectionUri -eq 'https://outlook.contoso.local/powershell-liveid/' -and
                    $AzureADAuthorizationEndpointUri -eq 'https://login.contoso.local/common'
                }
            }
        }
    }

    Context 'Power Platform REST' {
        It 'Should take the client id from the custom environment configuration' {
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{ CustomEnvironmentFile = $script:customEnvironmentFile } {
                param ($CustomEnvironmentFile)

                Mock -CommandName Invoke-RestMethod -MockWith { return @{ token_type = 'Bearer'; access_token = 'token' } }

                Connect-M365Tenant -Workload 'PowerPlatformREST' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.local' `
                    -ApplicationSecret 'super-secret' `
                    -CustomEnvironmentFileName $CustomEnvironmentFile

                $connection = Get-MSCloudLoginConnectionProfile -Workload 'PowerPlatformREST'
                $connection.ClientId | Should -Be '99999999-9999-9999-9999-999999999999'
                $connection.BapEndpoint | Should -Be 'api.bap.contoso.local'
                $connection.Audience | Should -Be 'https://powerapps.contoso.local/'
            }
        }
    }

    Context 'Microsoft Teams' {
        It 'Should refuse a custom environment connection outside of Windows PowerShell 5' {
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{ CustomEnvironmentFile = $script:customEnvironmentFile } {
                param ($CustomEnvironmentFile)

                Mock -CommandName Get-CsTeamsCallingPolicy -MockWith { throw 'no session' }
                Mock -CommandName Connect-MicrosoftTeams -MockWith { }
                Mock -CommandName Set-TeamsEnvironmentConfig -MockWith { }

                { Connect-M365Tenant -Workload 'MicrosoftTeams' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.local' `
                    -CertificateThumbprint 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD' `
                    -CustomEnvironmentFileName $CustomEnvironmentFile } |
                    Should -Throw '*only supported in PowerShell 5*'

                $connection = Get-MSCloudLoginConnectionProfile -Workload 'MicrosoftTeams'
                $connection.EnvironmentName | Should -Be 'Custom'
                $connection.TokenUrl | Should -Be 'https://login.contoso.local/contoso.local/oauth2/v2.0/token'
                $connection.TeamsScope | Should -Be 'https://teams.contoso.local/.default'
            }
        }
    }
}

Describe 'Connect-M365Tenant in the French sovereign cloud' {

    BeforeEach {
        InModuleScope 'MSCloudLoginAssistant' {
            $Script:MSCloudLoginConnectionProfile = $null
            $Script:MSCloudLoginTriedGetEnvironment = $true
            $Script:CloudEnvironmentInfo = ConvertFrom-Json '{ "tenant_region_scope": "FG", "token_endpoint": "https://login.sovcloud-identity.fr/t/oauth2/v2.0/token" }'
        }
    }

    AfterEach {
        InModuleScope 'MSCloudLoginAssistant' -Parameters @{ ModuleRoot = $script:moduleRoot } {
            param ($ModuleRoot)
            $Script:CustomEnvConfig = Import-PowerShellDataFile -Path (Join-Path $ModuleRoot 'CustomEnvironment.psd1')
            $Script:LoadedCustomEnvFileName = 'CustomEnvironment.psd1'
        }
    }

    It 'Should connect Exchange Online through the French sovereign endpoints' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Remove-MSCloudLoginProxyModule -MockWith { }
            Mock -CommandName Get-ConnectionInformation -MockWith { return @() }
            Mock -CommandName Disconnect-ExchangeOnline -MockWith { }
            Mock -CommandName Connect-ExchangeOnline -MockWith { }

            Connect-M365Tenant -Workload 'ExchangeOnline' `
                -ApplicationId '11111111-1111-1111-1111-111111111111' `
                -TenantId 'contoso.onsovcloud.fr' `
                -CertificateThumbprint 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD'

            $connection = Get-MSCloudLoginConnectionProfile -Workload 'ExchangeOnline'
            $connection.EnvironmentName | Should -Be 'AzureFranceCloud'
            $connection.ExchangeEnvironmentName | Should -Be 'Custom'
            $connection.ConnectionUri | Should -Be 'https://outlook.sovcloud.fr/PowerShell-LiveID'
        }
    }

    It 'Should register the French Teams endpoints and refuse the connection outside of Windows PowerShell 5' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Get-CsTeamsCallingPolicy -MockWith { throw 'no session' }
            Mock -CommandName Connect-MicrosoftTeams -MockWith { }
            Mock -CommandName Set-TeamsEnvironmentConfig -MockWith { }

            { Connect-M365Tenant -Workload 'MicrosoftTeams' `
                -ApplicationId '11111111-1111-1111-1111-111111111111' `
                -TenantId 'contoso.onsovcloud.fr' `
                -CertificateThumbprint 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD' } |
                Should -Throw '*only supported in PowerShell 5*'

            $connection = Get-MSCloudLoginConnectionProfile -Workload 'MicrosoftTeams'
            $connection.EnvironmentName | Should -Be 'AzureFranceCloud'
            $connection.AuthorizationUrl | Should -Be 'https://login.sovcloud-identity.fr/'
            $Script:CustomEnvConfig.CustomTeamsEndpoints.TeamsConfigApiEndPoint | Should -Be 'https://config.teams.sovcloud.fr'
        }
    }
}

Describe 'SharePoint Online REST admin URL discovery' {

    BeforeEach {
        InModuleScope 'MSCloudLoginAssistant' {
            $Script:MSCloudLoginConnectionProfile = $null
            $Script:MSCloudLoginTriedGetEnvironment = $true
            $Script:CloudEnvironmentInfo = $null
        }
    }

    It 'Should ask Microsoft Graph for the admin URL when connecting with credentials only' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Invoke-RestMethod -MockWith { return @{ token_type = 'Bearer'; access_token = 'token' } }
            Mock -CommandName Get-SPOAdminUrl -MockWith { return 'https://contoso-admin.sharepoint.com' }

            Connect-M365Tenant -Workload 'SharePointOnlineREST' `
                -Credential (New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force)))

            $connection = Get-MSCloudLoginConnectionProfile -Workload 'SharePointOnlineREST'
            $connection.AdminUrl | Should -Be 'https://contoso-admin.sharepoint.com'
            $connection.ConnectionUrl | Should -Be 'https://contoso-admin.sharepoint.com'
        }
    }

    It 'Should fail with a clear message when the admin URL cannot be resolved' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Invoke-RestMethod -MockWith { return @{ token_type = 'Bearer'; access_token = 'token' } }
            Mock -CommandName Get-SPOAdminUrl -MockWith { return '' }

            { Connect-M365Tenant -Workload 'SharePointOnlineREST' `
                -Credential (New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))) } |
                Should -Throw '*Unable to retrieve SharePoint Admin Url*'
        }
    }
}

Describe 'Security and Compliance Center connection URL reuse' {

    BeforeEach {
        InModuleScope 'MSCloudLoginAssistant' {
            $Script:MSCloudLoginConnectionProfile = $null
            $Script:MSCloudLoginTriedGetEnvironment = $true
            $Script:CloudEnvironmentInfo = $null
        }
    }

    It 'Should adopt the connection URI of an existing protection session' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Remove-MSCloudLoginProxyModule -MockWith { }
            Mock -CommandName Get-PSSession -MockWith { return @() }
            Mock -CommandName Connect-IPPSSession -MockWith { }
            Mock -CommandName Get-ConnectionInformation -MockWith {
                return @([PSCustomObject]@{
                    Name          = 'ExchangeOnlineProtection_1'
                    ConnectionUri = 'https://eur01b.ps.compliance.protection.contoso.com/powershell-liveid/'
                })
            }

            Connect-M365Tenant -Workload 'SecurityComplianceCenter' `
                -ApplicationId '11111111-1111-1111-1111-111111111111' `
                -TenantId 'contoso.onmicrosoft.com' `
                -CertificateThumbprint 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD'

            (Get-MSCloudLoginConnectionProfile -Workload 'SecurityComplianceCenter').ConnectionUrl |
                Should -Be 'https://eur01b.ps.compliance.protection.contoso.com/powershell-liveid/'
        }
    }
}
