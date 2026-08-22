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

Describe 'Connect-M365Tenant end-to-end for Exchange Online' {

    BeforeEach {
        InModuleScope 'MSCloudLoginAssistant' {
            $Script:MSCloudLoginConnectionProfile = $null
            $Script:MSCloudLoginTriedGetEnvironment = $true
            $Script:CloudEnvironmentInfo = $null
            $Script:MSCloudLoginCurrentLoadedModule = $null
        }
    }

    Context 'When connecting with a service principal' {
        It 'Should connect by environment name and report the tenant as the organization' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Remove-MSCloudLoginProxyModule -MockWith { }
                Mock -CommandName Get-ConnectionInformation -MockWith { return @() }
                Mock -CommandName Disconnect-ExchangeOnline -MockWith { }
                Mock -CommandName Connect-ExchangeOnline -MockWith { }

                Connect-M365Tenant -Workload 'ExchangeOnline' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -CertificateThumbprint 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD'

                $connection = Get-MSCloudLoginConnectionProfile -Workload 'ExchangeOnline'
                $connection.Connected | Should -BeTrue
                $connection.ExchangeEnvironmentName | Should -Be 'O365Default'
                $connection.LoadedAllCmdlets | Should -BeTrue

                Should -Invoke Connect-ExchangeOnline -Exactly 1 -ParameterFilter {
                    $AppId -eq '11111111-1111-1111-1111-111111111111' -and
                    $Organization -eq 'contoso.onmicrosoft.com' -and
                    $CertificateThumbprint -eq 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD' -and
                    $ExchangeEnvironmentName -eq 'O365Default'
                }
            }
        }

        It 'Should reject a certificate path connection when the tenant is not the onmicrosoft.com domain' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Remove-MSCloudLoginProxyModule -MockWith { }
                Mock -CommandName Get-ConnectionInformation -MockWith { return @() }
                Mock -CommandName Disconnect-ExchangeOnline -MockWith { }
                Mock -CommandName Connect-ExchangeOnline -MockWith { }

                { Connect-M365Tenant -Workload 'ExchangeOnline' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.com' `
                    -CertificatePath 'C:\certificates\contoso.pfx' `
                    -CertificatePassword (ConvertTo-SecureString 'certificate-password' -AsPlainText -Force) } |
                    Should -Throw '*primary domain in the format*'

                (Get-MSCloudLoginConnectionProfile -Workload 'ExchangeOnline').Connected | Should -BeFalse
            }
        }
    }

    Context 'When the tenant lives in a sovereign cloud' {
        It 'Should select the <ExpectedExchangeEnvironment> Exchange environment for <EnvironmentName>' -TestCases @(
            @{ EnvironmentName = 'AzureUSGovernment'; ExpectedExchangeEnvironment = 'O365USGovGCCHigh'; Discovery = '{ "tenant_region_sub_scope": "USGov", "token_endpoint": "https://login.microsoftonline.us/t/oauth2/v2.0/token" }' }
            @{ EnvironmentName = 'AzureDOD'; ExpectedExchangeEnvironment = 'O365USGovDoD'; Discovery = '{ "tenant_region_sub_scope": "DOD", "token_endpoint": "https://login.microsoftonline.us/t/oauth2/v2.0/token" }' }
            @{ EnvironmentName = 'AzureChinaCloud'; ExpectedExchangeEnvironment = 'O365China'; Discovery = '{ "token_endpoint": "https://login.partner.microsoftonline.cn/22222222-2222-2222-2222-222222222222/oauth2/v2.0/token" }' }
        ) {
            param ($EnvironmentName, $ExpectedExchangeEnvironment, $Discovery)
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{
                EnvironmentName             = $EnvironmentName
                ExpectedExchangeEnvironment = $ExpectedExchangeEnvironment
                Discovery                   = $Discovery
            } {
                param ($EnvironmentName, $ExpectedExchangeEnvironment, $Discovery)

                Mock -CommandName Remove-MSCloudLoginProxyModule -MockWith { }
                Mock -CommandName Get-ConnectionInformation -MockWith { return @() }
                Mock -CommandName Disconnect-ExchangeOnline -MockWith { }
                Mock -CommandName Connect-ExchangeOnline -MockWith { }

                $Script:CloudEnvironmentInfo = ConvertFrom-Json $Discovery

                Connect-M365Tenant -Workload 'ExchangeOnline' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -CertificateThumbprint 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD'

                $connection = Get-MSCloudLoginConnectionProfile -Workload 'ExchangeOnline'
                $connection.EnvironmentName | Should -Be $EnvironmentName
                $connection.ExchangeEnvironmentName | Should -Be $ExpectedExchangeEnvironment
            }
        }

        It 'Should connect the German sovereign cloud through its dedicated endpoint URIs' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Remove-MSCloudLoginProxyModule -MockWith { }
                Mock -CommandName Get-ConnectionInformation -MockWith { return @() }
                Mock -CommandName Disconnect-ExchangeOnline -MockWith { }
                Mock -CommandName Connect-ExchangeOnline -MockWith { }

                $Script:CloudEnvironmentInfo = ConvertFrom-Json '{ "tenant_region_scope": "GG2", "token_endpoint": "https://login.sovcloud-identity.de/t/oauth2/v2.0/token" }'

                Connect-M365Tenant -Workload 'ExchangeOnline' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onsovcloud.de' `
                    -CertificateThumbprint 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD'

                $connection = Get-MSCloudLoginConnectionProfile -Workload 'ExchangeOnline'
                $connection.EnvironmentName | Should -Be 'AzureGermanyCloud'
                $connection.ExchangeEnvironmentName | Should -Be 'Custom'
                $connection.ConnectionUri | Should -Be 'https://outlook.sovcloud.de/PowerShell-LiveID'

                Should -Invoke Connect-ExchangeOnline -Exactly 1 -ParameterFilter {
                    $ConnectionUri -eq 'https://outlook.sovcloud.de/PowerShell-LiveID'
                }
            }
        }
    }

    Context 'When connecting with a token or a managed identity' {
        It 'Should pass the access token and the organization to Connect-ExchangeOnline' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Remove-MSCloudLoginProxyModule -MockWith { }
                Mock -CommandName Get-ConnectionInformation -MockWith { return @() }
                Mock -CommandName Disconnect-ExchangeOnline -MockWith { }
                Mock -CommandName Connect-ExchangeOnline -MockWith { }

                Connect-M365Tenant -Workload 'ExchangeOnline' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -AccessTokens @('caller-supplied-token')

                (Get-MSCloudLoginConnectionProfile -Workload 'ExchangeOnline').Connected | Should -BeTrue
                Should -Invoke Connect-ExchangeOnline -Exactly 1 -ParameterFilter {
                    $AccessToken -eq 'caller-supplied-token' -and $Organization -eq 'contoso.onmicrosoft.com'
                }
            }
        }

        It 'Should flag the managed identity connection as multi-factor authenticated' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Remove-MSCloudLoginProxyModule -MockWith { }
                Mock -CommandName Get-ConnectionInformation -MockWith { return @() }
                Mock -CommandName Disconnect-ExchangeOnline -MockWith { }
                Mock -CommandName Connect-ExchangeOnline -MockWith { }

                Connect-M365Tenant -Workload 'ExchangeOnline' -Identity -TenantId 'contoso.onmicrosoft.com'

                $connection = Get-MSCloudLoginConnectionProfile -Workload 'ExchangeOnline'
                $connection.Connected | Should -BeTrue
                $connection.MultiFactorAuthentication | Should -BeTrue
                Should -Invoke Connect-ExchangeOnline -Exactly 1 -ParameterFilter { $ManagedIdentity.IsPresent }
            }
        }
    }

    Context 'When connecting with user credentials' {
        It 'Should connect without a delegated organization' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Remove-MSCloudLoginProxyModule -MockWith { }
                Mock -CommandName Get-ConnectionInformation -MockWith { return @() }
                Mock -CommandName Disconnect-ExchangeOnline -MockWith { }
                Mock -CommandName Connect-ExchangeOnline -MockWith { }

                Connect-M365Tenant -Workload 'ExchangeOnline' `
                    -Credential (New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force)))

                (Get-MSCloudLoginConnectionProfile -Workload 'ExchangeOnline').Connected | Should -BeTrue
                Should -Invoke Connect-ExchangeOnline -Exactly 1 -ParameterFilter {
                    $null -ne $Credential -and [System.String]::IsNullOrEmpty($DelegatedOrganization)
                }
            }
        }

        It 'Should pass the tenant as delegated organization when a tenant id is supplied' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Remove-MSCloudLoginProxyModule -MockWith { }
                Mock -CommandName Get-ConnectionInformation -MockWith { return @() }
                Mock -CommandName Disconnect-ExchangeOnline -MockWith { }
                Mock -CommandName Connect-ExchangeOnline -MockWith { }

                Connect-M365Tenant -Workload 'ExchangeOnline' `
                    -TenantId 'fabrikam.onmicrosoft.com' `
                    -Credential (New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force)))

                Should -Invoke Connect-ExchangeOnline -Exactly 1 -ParameterFilter {
                    $DelegatedOrganization -eq 'fabrikam.onmicrosoft.com'
                }
            }
        }

        It 'Should fall back to the interactive MFA sign-in when the password grant is refused' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Remove-MSCloudLoginProxyModule -MockWith { }
                Mock -CommandName Get-ConnectionInformation -MockWith { return @() }
                Mock -CommandName Disconnect-ExchangeOnline -MockWith { }
                Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $false }
                Mock -CommandName Connect-ExchangeOnline -MockWith {
                    if ($null -ne $Credential)
                    {
                        throw 'AADSTS50076: multi-factor authentication is required.'
                    }
                }

                Connect-M365Tenant -Workload 'ExchangeOnline' `
                    -Credential (New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force)))

                $connection = Get-MSCloudLoginConnectionProfile -Workload 'ExchangeOnline'
                $connection.Connected | Should -BeTrue
                $connection.MultiFactorAuthentication | Should -BeTrue
                Should -Invoke Connect-ExchangeOnline -Exactly 1 -ParameterFilter {
                    $UserPrincipalName -eq 'admin@contoso.onmicrosoft.com'
                }
            }
        }
    }

    Context 'When only a subset of the cmdlets is requested' {
        It 'Should always add Get-AcceptedDomain to the requested cmdlets' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Remove-MSCloudLoginProxyModule -MockWith { }
                Mock -CommandName Get-ConnectionInformation -MockWith { return @() }
                Mock -CommandName Disconnect-ExchangeOnline -MockWith { }
                Mock -CommandName Connect-ExchangeOnline -MockWith { }

                Connect-M365Tenant -Workload 'ExchangeOnline' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -CertificateThumbprint 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD' `
                    -ExchangeOnlineCmdlets @('Get-Mailbox')

                (Get-MSCloudLoginConnectionProfile -Workload 'ExchangeOnline').LoadedAllCmdlets | Should -BeFalse
                Should -Invoke Connect-ExchangeOnline -Exactly 1 -ParameterFilter {
                    $CommandName -contains 'Get-Mailbox' -and $CommandName -contains 'Get-AcceptedDomain'
                }
            }
        }
    }

    Context 'When an Exchange Online session already exists' {
        It 'Should reuse the loaded proxy module instead of connecting again' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Remove-MSCloudLoginProxyModule -MockWith { }
                Mock -CommandName Connect-ExchangeOnline -MockWith { }
                Mock -CommandName Get-ConnectionInformation -MockWith { return @() }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.LoadedAllCmdlets = $true
                $Script:MSCloudLoginCurrentLoadedModule = 'EXO'

                Connect-M365Tenant -Workload 'ExchangeOnline' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -CertificateThumbprint 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD'

                (Get-MSCloudLoginConnectionProfile -Workload 'ExchangeOnline').Connected | Should -BeTrue
                Should -Invoke Connect-ExchangeOnline -Exactly 0
            }
        }

        It 'Should adopt an existing session that matches the application and the tenant' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Remove-MSCloudLoginProxyModule -MockWith { }
                Mock -CommandName Connect-ExchangeOnline -MockWith { }
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Get-ConnectionInformation -MockWith {
                    return @([PSCustomObject]@{
                        Name         = 'ExchangeOnline_1'
                        AppId        = '11111111-1111-1111-1111-111111111111'
                        Organization = 'contoso.onmicrosoft.com'
                        ModuleName   = 'tmpEXO_abcdefgh'
                    })
                }

                Connect-M365Tenant -Workload 'ExchangeOnline' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -CertificateThumbprint 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD'

                $connection = Get-MSCloudLoginConnectionProfile -Workload 'ExchangeOnline'
                $connection.Connected | Should -BeTrue
                $connection.LoadedAllCmdlets | Should -BeTrue
                Should -Invoke Import-Module -Exactly 1 -ParameterFilter { $Name -eq 'tmpEXO_abcdefgh' }
                Should -Invoke Connect-ExchangeOnline -Exactly 0
            }
        }
    }

    Context 'When Exchange Online is reset' {
        It 'Should disconnect the session and drop the loaded cmdlets' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Remove-MSCloudLoginProxyModule -MockWith { }
                Mock -CommandName Get-ConnectionInformation -MockWith { return @() }
                Mock -CommandName Disconnect-ExchangeOnline -MockWith { }
                Mock -CommandName Connect-ExchangeOnline -MockWith { }

                Connect-M365Tenant -Workload 'ExchangeOnline' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -CertificateThumbprint 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD' `
                    -ExchangeOnlineCmdlets @('Get-Mailbox')

                Reset-MSCloudLoginConnectionProfileContext -Workload 'ExchangeOnline'

                $connection = Get-MSCloudLoginConnectionProfile -Workload 'ExchangeOnline'
                $connection.Connected | Should -BeFalse
                $connection.LoadedAllCmdlets | Should -BeFalse
                $connection.LoadedCmdlets | Should -BeNullOrEmpty
                $connection.CmdletsToLoad | Should -BeNullOrEmpty
            }
        }
    }
}

Describe 'Connect-M365Tenant end-to-end for the Security and Compliance Center' {

    BeforeEach {
        InModuleScope 'MSCloudLoginAssistant' {
            $Script:MSCloudLoginConnectionProfile = $null
            $Script:MSCloudLoginTriedGetEnvironment = $true
            $Script:CloudEnvironmentInfo = $null
        }
    }

    Context 'When connecting with a service principal' {
        It 'Should target the commercial compliance endpoint with the certificate thumbprint' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Remove-MSCloudLoginProxyModule -MockWith { }
                Mock -CommandName Get-ConnectionInformation -MockWith { return $null }
                Mock -CommandName Get-PSSession -MockWith { return @() }
                Mock -CommandName Connect-IPPSSession -MockWith { }

                Connect-M365Tenant -Workload 'SecurityComplianceCenter' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -CertificateThumbprint 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD'

                $connection = Get-MSCloudLoginConnectionProfile -Workload 'SecurityComplianceCenter'
                $connection.Connected | Should -BeTrue
                $connection.ConnectionUrl | Should -Be 'https://ps.compliance.protection.outlook.com/powershell-liveid/'
                $connection.AzureADAuthorizationEndpointUri | Should -Be 'https://login.microsoftonline.com/organizations'

                Should -Invoke Connect-IPPSSession -Exactly 1 -ParameterFilter {
                    $AppId -eq '11111111-1111-1111-1111-111111111111' -and
                    $Organization -eq 'contoso.onmicrosoft.com' -and
                    $ConnectionUri -eq 'https://ps.compliance.protection.outlook.com/powershell-liveid/'
                }
            }
        }

        It 'Should pass the certificate file path and password' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Remove-MSCloudLoginProxyModule -MockWith { }
                Mock -CommandName Get-ConnectionInformation -MockWith { return $null }
                Mock -CommandName Get-PSSession -MockWith { return @() }
                Mock -CommandName Connect-IPPSSession -MockWith { }

                Connect-M365Tenant -Workload 'SecurityComplianceCenter' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -CertificatePath 'C:\certificates\contoso.pfx' `
                    -CertificatePassword (ConvertTo-SecureString 'certificate-password' -AsPlainText -Force)

                Should -Invoke Connect-IPPSSession -Exactly 1 -ParameterFilter {
                    $CertificateFilePath -eq 'C:\certificates\contoso.pfx'
                }
            }
        }
    }

    Context 'When a search only session is requested' {
        It 'Should forward the search only switch and keep it on the connection profile' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Remove-MSCloudLoginProxyModule -MockWith { }
                Mock -CommandName Get-ConnectionInformation -MockWith { return $null }
                Mock -CommandName Get-PSSession -MockWith { return @() }
                Mock -CommandName Connect-IPPSSession -MockWith { }

                Connect-M365Tenant -Workload 'SecurityComplianceCenter' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -CertificateThumbprint 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD' `
                    -EnableSearchOnlySession

                (Get-MSCloudLoginConnectionProfile -Workload 'SecurityComplianceCenter').EnableSearchOnlySession | Should -BeTrue
                Should -Invoke Connect-IPPSSession -Exactly 1 -ParameterFilter { $EnableSearchOnlySession.IsPresent }
            }
        }
    }

    Context 'When connecting with user credentials and a tenant id' {
        It 'Should replace the organizations segment of the authorization endpoint with the tenant' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Remove-MSCloudLoginProxyModule -MockWith { }
                Mock -CommandName Get-ConnectionInformation -MockWith { return $null }
                Mock -CommandName Get-PSSession -MockWith { return @() }
                Mock -CommandName Connect-IPPSSession -MockWith { }

                Connect-M365Tenant -Workload 'SecurityComplianceCenter' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -Credential (New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force)))

                Should -Invoke Connect-IPPSSession -Exactly 1 -ParameterFilter {
                    $AzureADAuthorizationEndpointUri -eq 'https://login.microsoftonline.com/contoso.onmicrosoft.com' -and
                    $DelegatedOrganization -eq 'contoso.onmicrosoft.com'
                }
            }
        }

        It 'Should fall back to the interactive MFA sign-in when the password grant is refused' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Remove-MSCloudLoginProxyModule -MockWith { }
                Mock -CommandName Get-ConnectionInformation -MockWith { return $null }
                Mock -CommandName Get-PSSession -MockWith { return @() }
                Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $false }
                Mock -CommandName Connect-IPPSSession -MockWith {
                    if ($null -ne $Credential)
                    {
                        throw 'AADSTS50076: multi-factor authentication is required.'
                    }
                }

                Connect-M365Tenant -Workload 'SecurityComplianceCenter' `
                    -Credential (New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force)))

                $connection = Get-MSCloudLoginConnectionProfile -Workload 'SecurityComplianceCenter'
                $connection.Connected | Should -BeTrue
                $connection.MultiFactorAuthentication | Should -BeTrue
                Should -Invoke Connect-IPPSSession -Exactly 1 -ParameterFilter {
                    $UserPrincipalName -eq 'admin@contoso.onmicrosoft.com'
                }
            }
        }
    }

    Context 'When connecting with a managed identity' {
        It 'Should connect the compliance session as a managed identity' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Remove-MSCloudLoginProxyModule -MockWith { }
                Mock -CommandName Get-ConnectionInformation -MockWith { return $null }
                Mock -CommandName Get-PSSession -MockWith { return @() }
                Mock -CommandName Connect-IPPSSession -MockWith { }

                Connect-M365Tenant -Workload 'SecurityComplianceCenter' -Identity -TenantId 'contoso.onmicrosoft.com'

                (Get-MSCloudLoginConnectionProfile -Workload 'SecurityComplianceCenter').Connected | Should -BeTrue
                Should -Invoke Connect-IPPSSession -Exactly 1 -ParameterFilter { $ManagedIdentity.IsPresent }
            }
        }
    }

    Context 'When an existing compliance session is still open' {
        It 'Should re-import the session instead of connecting again' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Remove-MSCloudLoginProxyModule -MockWith { }
                Mock -CommandName Get-ConnectionInformation -MockWith { return $null }
                Mock -CommandName Connect-IPPSSession -MockWith { }
                Mock -CommandName Import-PSSession -MockWith { return 'tmpSCC_abcdefgh' }
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Get-PSSession -MockWith {
                    return @([PSCustomObject]@{ ComputerName = 'ps.compliance.protection.outlook.com'; State = 'Opened' })
                }

                Connect-M365Tenant -Workload 'SecurityComplianceCenter' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -CertificateThumbprint 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD'

                (Get-MSCloudLoginConnectionProfile -Workload 'SecurityComplianceCenter').Connected | Should -BeTrue
                Should -Invoke Import-PSSession -Exactly 1
                Should -Invoke Connect-IPPSSession -Exactly 0
            }
        }
    }

    Context 'When the Security and Compliance Center is reset' {
        It 'Should disconnect the underlying Exchange session' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Remove-MSCloudLoginProxyModule -MockWith { }
                Mock -CommandName Get-ConnectionInformation -MockWith { return $null }
                Mock -CommandName Get-PSSession -MockWith { return @() }
                Mock -CommandName Connect-IPPSSession -MockWith { }
                Mock -CommandName Disconnect-ExchangeOnline -MockWith { }

                Connect-M365Tenant -Workload 'SecurityComplianceCenter' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -CertificateThumbprint 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD'

                Reset-MSCloudLoginConnectionProfileContext -Workload 'SecurityComplianceCenter'

                Should -Invoke Disconnect-ExchangeOnline -Exactly 1
                (Get-MSCloudLoginConnectionProfile -Workload 'SecurityComplianceCenter').Connected | Should -BeFalse
            }
        }
    }
}

Describe 'Connect-M365Tenant end-to-end for Microsoft Teams' {

    BeforeEach {
        InModuleScope 'MSCloudLoginAssistant' {
            $Script:MSCloudLoginConnectionProfile = $null
            $Script:MSCloudLoginTriedGetEnvironment = $true
            $Script:CloudEnvironmentInfo = $null
        }
    }

    Context 'When no Teams session exists yet' {
        It 'Should connect with the certificate thumbprint of the application' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Get-CsTeamsCallingPolicy -MockWith { throw 'no session' }
                Mock -CommandName Connect-MicrosoftTeams -MockWith { }

                Connect-M365Tenant -Workload 'MicrosoftTeams' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -CertificateThumbprint 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD'

                (Get-MSCloudLoginConnectionProfile -Workload 'MicrosoftTeams').Connected | Should -BeTrue
                Should -Invoke Connect-MicrosoftTeams -Exactly 1 -ParameterFilter {
                    $ApplicationId -eq '11111111-1111-1111-1111-111111111111' -and
                    $TenantId -eq 'contoso.onmicrosoft.com' -and
                    $CertificateThumbprint -eq 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD' -and
                    [System.String]::IsNullOrEmpty($TeamsEnvironmentName)
                }
            }
        }

        It 'Should load the certificate from disk for a certificate path connection' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Get-CsTeamsCallingPolicy -MockWith { throw 'no session' }
                Mock -CommandName Connect-MicrosoftTeams -MockWith { }
                Mock -CommandName Get-MSCloudLoginCertificate -MockWith {
                    return New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
                }

                Connect-M365Tenant -Workload 'MicrosoftTeams' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -CertificatePath 'C:\certificates\contoso.pfx' `
                    -CertificatePassword (ConvertTo-SecureString 'certificate-password' -AsPlainText -Force)

                (Get-MSCloudLoginConnectionProfile -Workload 'MicrosoftTeams').Connected | Should -BeTrue
                Should -Invoke Connect-MicrosoftTeams -Exactly 1 -ParameterFilter { $null -ne $Certificate }
            }
        }

        It 'Should select the <ExpectedTeamsEnvironment> Teams environment for <EnvironmentName>' -TestCases @(
            @{ EnvironmentName = 'AzureUSGovernment'; ExpectedTeamsEnvironment = 'TeamsGCCH'; Discovery = '{ "tenant_region_sub_scope": "USGov", "token_endpoint": "https://login.microsoftonline.us/t/oauth2/v2.0/token" }' }
            @{ EnvironmentName = 'AzureDOD'; ExpectedTeamsEnvironment = 'TeamsDOD'; Discovery = '{ "tenant_region_sub_scope": "DOD", "token_endpoint": "https://login.microsoftonline.us/t/oauth2/v2.0/token" }' }
            @{ EnvironmentName = 'AzureChinaCloud'; ExpectedTeamsEnvironment = 'TeamsChina'; Discovery = '{ "token_endpoint": "https://login.partner.microsoftonline.cn/22222222-2222-2222-2222-222222222222/oauth2/v2.0/token" }' }
        ) {
            param ($EnvironmentName, $ExpectedTeamsEnvironment, $Discovery)
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{
                EnvironmentName          = $EnvironmentName
                ExpectedTeamsEnvironment = $ExpectedTeamsEnvironment
                Discovery                = $Discovery
            } {
                param ($EnvironmentName, $ExpectedTeamsEnvironment, $Discovery)

                Mock -CommandName Get-CsTeamsCallingPolicy -MockWith { throw 'no session' }
                Mock -CommandName Connect-MicrosoftTeams -MockWith { }

                $Script:CloudEnvironmentInfo = ConvertFrom-Json $Discovery

                Connect-M365Tenant -Workload 'MicrosoftTeams' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -CertificateThumbprint 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD'

                (Get-MSCloudLoginConnectionProfile -Workload 'MicrosoftTeams').EnvironmentName | Should -Be $EnvironmentName
                Should -Invoke Connect-MicrosoftTeams -Exactly 1 -ParameterFilter {
                    $TeamsEnvironmentName -eq $ExpectedTeamsEnvironment
                }
            }
        }

        It 'Should connect with a managed identity' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Get-CsTeamsCallingPolicy -MockWith { throw 'no session' }
                Mock -CommandName Connect-MicrosoftTeams -MockWith { }

                Connect-M365Tenant -Workload 'MicrosoftTeams' -Identity -TenantId 'contoso.onmicrosoft.com'

                (Get-MSCloudLoginConnectionProfile -Workload 'MicrosoftTeams').Connected | Should -BeTrue
                Should -Invoke Connect-MicrosoftTeams -Exactly 1 -ParameterFilter { $Identity.IsPresent }
            }
        }

        It 'Should forward every supplied access token' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Get-CsTeamsCallingPolicy -MockWith { throw 'no session' }
                Mock -CommandName Connect-MicrosoftTeams -MockWith { }

                Connect-M365Tenant -Workload 'MicrosoftTeams' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -AccessTokens @('graph-token', 'teams-token')

                (Get-MSCloudLoginConnectionProfile -Workload 'MicrosoftTeams').Connected | Should -BeTrue
                Should -Invoke Connect-MicrosoftTeams -Exactly 1 -ParameterFilter {
                    $AccessTokens.Count -eq 2 -and $AccessTokens[0] -eq 'graph-token'
                }
            }
        }

        It 'Should retry through the MFA sign-in when the credential sign-in is refused' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Get-CsTeamsCallingPolicy -MockWith { throw 'no session' }
                Mock -CommandName Disconnect-MicrosoftTeams -MockWith { }
                Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $false }
                Mock -CommandName Connect-MicrosoftTeams -MockWith {
                    if ($null -ne $Credential)
                    {
                        throw 'AADSTS50076: multi-factor authentication is required.'
                    }
                }

                Connect-M365Tenant -Workload 'MicrosoftTeams' `
                    -Credential (New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force)))

                $connection = Get-MSCloudLoginConnectionProfile -Workload 'MicrosoftTeams'
                $connection.Connected | Should -BeTrue
                $connection.MultiFactorAuthentication | Should -BeTrue
                Should -Invoke Disconnect-MicrosoftTeams -Exactly 1
            }
        }

        It 'Should surface a credential sign-in failure that is unrelated to MFA' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Get-CsTeamsCallingPolicy -MockWith { throw 'no session' }
                Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $true }
                Mock -CommandName Connect-MicrosoftTeams -MockWith { throw 'AADSTS50126: Invalid username or password.' }

                { Connect-M365Tenant -Workload 'MicrosoftTeams' `
                    -Credential (New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))) } |
                    Should -Throw '*AADSTS50126*'

                (Get-MSCloudLoginConnectionProfile -Workload 'MicrosoftTeams').Connected | Should -BeFalse
            }
        }
    }

    Context 'When a Teams session is still usable' {
        It 'Should reuse the session without calling Connect-MicrosoftTeams' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MicrosoftTeams -MockWith { }
                Mock -CommandName Get-MSCloudLoginAccessToken -MockWith { return 'access-token' }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }

                Connect-M365Tenant -Workload 'MicrosoftTeams' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -CertificateThumbprint 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD'

                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $true }
                Mock -CommandName Connect-MicrosoftTeams -MockWith { }

                Connect-M365Tenant -Workload 'MicrosoftTeams' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -CertificateThumbprint 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD'

                (Get-MSCloudLoginConnectionProfile -Workload 'MicrosoftTeams').Connected | Should -BeTrue
                Should -Invoke Connect-MicrosoftTeams -Exactly 1
            }
        }
    }

    Context 'When Microsoft Teams is reset' {
        It 'Should disconnect the Teams session' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Get-CsTeamsCallingPolicy -MockWith { throw 'no session' }
                Mock -CommandName Connect-MicrosoftTeams -MockWith { }
                Mock -CommandName Disconnect-MicrosoftTeams -MockWith { }

                Connect-M365Tenant -Workload 'MicrosoftTeams' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -CertificateThumbprint 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD'

                Reset-MSCloudLoginConnectionProfileContext -Workload 'MicrosoftTeams'

                Should -Invoke Disconnect-MicrosoftTeams -Exactly 1
                (Get-MSCloudLoginConnectionProfile -Workload 'MicrosoftTeams').Connected | Should -BeFalse
            }
        }
    }
}

Describe 'Connect-M365Tenant end-to-end for the Power Platform' {

    BeforeEach {
        InModuleScope 'MSCloudLoginAssistant' {
            $Script:MSCloudLoginConnectionProfile = $null
            $Script:MSCloudLoginTriedGetEnvironment = $true
            $Script:CloudEnvironmentInfo = $null
        }
    }

    Context 'When connecting with a service principal' {
        It 'Should read the bearer token out of the PowerApps session for a thumbprint connection' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Add-PowerAppsAccount -MockWith {
                    $Global:currentSession = @{
                        resourceTokens = @{
                            'https://service.powerapps.com/' = @{ accessToken = 'powerapps-token' }
                        }
                    }
                }

                Connect-M365Tenant -Workload 'PowerPlatforms' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -CertificateThumbprint 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD'

                $connection = Get-MSCloudLoginConnectionProfile -Workload 'PowerPlatforms'
                $connection.Connected | Should -BeTrue
                $connection.Endpoint | Should -Be 'prod'
                $connection.AccessTokens | Should -Be @('Bearer powerapps-token')

                Should -Invoke Add-PowerAppsAccount -Exactly 1 -ParameterFilter {
                    $ApplicationId -eq '11111111-1111-1111-1111-111111111111' -and
                    $CertificateThumbprint -eq 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD' -and
                    $Endpoint -eq 'prod'
                }
            }
        }

        It 'Should pass the application secret through to Add-PowerAppsAccount' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Add-PowerAppsAccount -MockWith { }

                Connect-M365Tenant -Workload 'PowerPlatforms' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -ApplicationSecret 'super-secret'

                (Get-MSCloudLoginConnectionProfile -Workload 'PowerPlatforms').Connected | Should -BeTrue
                Should -Invoke Add-PowerAppsAccount -Exactly 1 -ParameterFilter { $ClientSecret -eq 'super-secret' }
            }
        }
    }

    Context 'When the tenant lives in a government cloud' {
        It 'Should select the <ExpectedEndpoint> PowerApps endpoint for the <SubScope> sub scope' -TestCases @(
            @{ SubScope = 'DODCON'; ExpectedEndpoint = 'usgovhigh' }
            @{ SubScope = 'DOD'; ExpectedEndpoint = 'dod' }
            @{ SubScope = 'GCC'; ExpectedEndpoint = 'usgov' }
        ) {
            param ($SubScope, $ExpectedEndpoint)
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{ SubScope = $SubScope; ExpectedEndpoint = $ExpectedEndpoint } {
                param ($SubScope, $ExpectedEndpoint)

                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Add-PowerAppsAccount -MockWith { }

                $Script:CloudEnvironmentInfo = [PSCustomObject]@{
                    tenant_region_sub_scope = $SubScope
                    token_endpoint          = 'https://login.microsoftonline.us/t/oauth2/v2.0/token'
                }

                Connect-M365Tenant -Workload 'PowerPlatforms' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -ApplicationSecret 'super-secret'

                (Get-MSCloudLoginConnectionProfile -Workload 'PowerPlatforms').Endpoint | Should -Be $ExpectedEndpoint
                Should -Invoke Add-PowerAppsAccount -Exactly 1 -ParameterFilter { $Endpoint -eq $ExpectedEndpoint }
            }
        }
    }

    Context 'When connecting with user credentials' {
        It 'Should pass the user name and password to Add-PowerAppsAccount' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Add-PowerAppsAccount -MockWith { }

                Connect-M365Tenant -Workload 'PowerPlatforms' `
                    -Credential (New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force)))

                (Get-MSCloudLoginConnectionProfile -Workload 'PowerPlatforms').Connected | Should -BeTrue
                Should -Invoke Add-PowerAppsAccount -Exactly 1 -ParameterFilter {
                    $Username -eq 'admin@contoso.onmicrosoft.com'
                }
            }
        }

        It 'Should reject a credential that is combined with a tenant id' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Add-PowerAppsAccount -MockWith { }

                { Connect-M365Tenant -Workload 'PowerPlatforms' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -Credential (New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))) } |
                    Should -Throw '*cannot specify TenantId with Credentials*'
            }
        }

        It 'Should retry against the preview endpoint when the user type is unknown' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Add-PowerAppsAccount -MockWith {
                    if ($Endpoint -ne 'preview')
                    {
                        throw 'unknown_user_type: Unknown User Type'
                    }
                }

                Connect-M365Tenant -Workload 'PowerPlatforms' `
                    -Credential (New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force)))

                (Get-MSCloudLoginConnectionProfile -Workload 'PowerPlatforms').Connected | Should -BeTrue
                Should -Invoke Add-PowerAppsAccount -Exactly 1 -ParameterFilter { $Endpoint -eq 'preview' }
            }
        }

        It 'Should fall back to the interactive MFA sign-in when the password grant is refused' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $false }
                Mock -CommandName Add-PowerAppsAccount -MockWith {
                    if (-not [System.String]::IsNullOrEmpty($Username))
                    {
                        throw 'AADSTS50076: multi-factor authentication is required.'
                    }
                }

                Connect-M365Tenant -Workload 'PowerPlatforms' `
                    -Credential (New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force)))

                $connection = Get-MSCloudLoginConnectionProfile -Workload 'PowerPlatforms'
                $connection.Connected | Should -BeTrue
                $connection.MultiFactorAuthentication | Should -BeTrue
            }
        }
    }

    Context 'When the Power Platform is reset' {
        It 'Should clear the cached PowerApps session' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Add-PowerAppsAccount -MockWith { }

                Connect-M365Tenant -Workload 'PowerPlatforms' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -ApplicationSecret 'super-secret'
                $Global:currentSession = @{ resourceTokens = @{} }

                Reset-MSCloudLoginConnectionProfileContext -Workload 'PowerPlatform'

                $Global:currentSession | Should -BeNullOrEmpty
                (Get-MSCloudLoginConnectionProfile -Workload 'PowerPlatforms').Connected | Should -BeFalse
            }
        }
    }
}

Describe 'Connect-M365Tenant end-to-end for Azure' {

    BeforeEach {
        InModuleScope 'MSCloudLoginAssistant' {
            $Script:MSCloudLoginConnectionProfile = $null
            $Script:MSCloudLoginTriedGetEnvironment = $true
            $Script:CloudEnvironmentInfo = $null
        }
    }

    Context 'When connecting with a service principal' {
        It 'Should sign in as a service principal and record the management URL' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Invoke-WebRequest -MockWith {
                    return @{ Content = '{ "token_endpoint": "https://login.microsoftonline.com/t/oauth2/v2.0/token" }' }
                }
                Mock -CommandName Connect-AzAccount -MockWith { }
                Mock -CommandName Get-AzContext -MockWith {
                    return @{ Environment = @{ ResourceManagerUrl = 'https://management.azure.com/' } }
                }

                Connect-M365Tenant -Workload 'Azure' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -CertificateThumbprint 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD'

                $connection = Get-MSCloudLoginConnectionProfile -Workload 'Azure'
                $connection.Connected | Should -BeTrue
                $connection.ManagementUrl | Should -Be 'https://management.azure.com/'

                Should -Invoke Connect-AzAccount -Exactly 1 -ParameterFilter {
                    $ServicePrincipal.IsPresent -and
                    $CertificateThumbprint -eq 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD' -and
                    $Environment -eq 'AzureCloud'
                }
            }
        }

        It 'Should scope the sign-in to the requested subscription' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Invoke-WebRequest -MockWith {
                    return @{ Content = '{ "token_endpoint": "https://login.microsoftonline.com/t/oauth2/v2.0/token" }' }
                }
                Mock -CommandName Connect-AzAccount -MockWith { }
                Mock -CommandName Get-AzContext -MockWith {
                    return @{ Environment = @{ ResourceManagerUrl = 'https://management.azure.com/' } }
                }

                Connect-M365Tenant -Workload 'Azure' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -ApplicationSecret 'super-secret' `
                    -SubscriptionId '44444444-4444-4444-4444-444444444444'

                (Get-MSCloudLoginConnectionProfile -Workload 'Azure').SubscriptionId | Should -Be '44444444-4444-4444-4444-444444444444'
                Should -Invoke Connect-AzAccount -Exactly 1 -ParameterFilter {
                    $Subscription -eq '44444444-4444-4444-4444-444444444444' -and
                    $Credential.UserName -eq '11111111-1111-1111-1111-111111111111'
                }
            }
        }
    }

    Context 'When connecting with a token or a managed identity' {
        It 'Should sign in with the supplied access token under a synthetic account id' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Invoke-WebRequest -MockWith {
                    return @{ Content = '{ "token_endpoint": "https://login.microsoftonline.com/t/oauth2/v2.0/token" }' }
                }
                Mock -CommandName Connect-AzAccount -MockWith { }
                Mock -CommandName Get-AzContext -MockWith {
                    return @{ Environment = @{ ResourceManagerUrl = 'https://management.azure.com/' } }
                }

                Connect-M365Tenant -Workload 'Azure' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -AccessTokens @('caller-supplied-token')

                (Get-MSCloudLoginConnectionProfile -Workload 'Azure').Connected | Should -BeTrue
                Should -Invoke Connect-AzAccount -Exactly 1 -ParameterFilter {
                    $AccessToken -eq 'caller-supplied-token' -and $AccountId -eq 'MSCloudLoginAssistant'
                }
            }
        }

        It 'Should sign in with a managed identity' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Invoke-WebRequest -MockWith {
                    return @{ Content = '{ "token_endpoint": "https://login.microsoftonline.com/t/oauth2/v2.0/token" }' }
                }
                Mock -CommandName Connect-AzAccount -MockWith { }
                Mock -CommandName Get-AzContext -MockWith {
                    return @{ Environment = @{ ResourceManagerUrl = 'https://management.azure.com/' } }
                }

                Connect-M365Tenant -Workload 'Azure' -Identity -TenantId 'contoso.onmicrosoft.com'

                (Get-MSCloudLoginConnectionProfile -Workload 'Azure').Connected | Should -BeTrue
                Should -Invoke Connect-AzAccount -Exactly 1 -ParameterFilter { $Identity.IsPresent }
            }
        }
    }

    Context 'When connecting with user credentials' {
        It 'Should derive the tenant from the credential when none was supplied' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Invoke-WebRequest -MockWith {
                    return @{ Content = '{ "token_endpoint": "https://login.microsoftonline.com/t/oauth2/v2.0/token" }' }
                }
                Mock -CommandName Connect-AzAccount -MockWith { }
                Mock -CommandName Get-AzContext -MockWith {
                    return @{ Environment = @{ ResourceManagerUrl = 'https://management.azure.com/' } }
                }

                Connect-M365Tenant -Workload 'Azure' `
                    -Credential (New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force)))

                (Get-MSCloudLoginConnectionProfile -Workload 'Azure').TenantId | Should -Be 'contoso.onmicrosoft.com'
                Should -Invoke Connect-AzAccount -Exactly 1 -ParameterFilter {
                    $TenantId -eq 'contoso.onmicrosoft.com'
                }
            }
        }

        It 'Should fall back to the interactive sign-in when the account requires MFA' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Invoke-WebRequest -MockWith {
                    return @{ Content = '{ "token_endpoint": "https://login.microsoftonline.com/t/oauth2/v2.0/token" }' }
                }
                Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $false }
                Mock -CommandName Get-AzContext -MockWith {
                    return @{ Environment = @{ ResourceManagerUrl = 'https://management.azure.com/' } }
                }
                Mock -CommandName Connect-AzAccount -MockWith {
                    if ($null -ne $Credential)
                    {
                        throw 'AADSTS50076: multi-factor authentication is required.'
                    }
                }

                Connect-M365Tenant -Workload 'Azure' `
                    -Credential (New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force)))

                $connection = Get-MSCloudLoginConnectionProfile -Workload 'Azure'
                $connection.Connected | Should -BeTrue
                $connection.MultiFactorAuthentication | Should -BeTrue
            }
        }
    }

    Context 'When Azure is reset' {
        It 'Should sign out of the Azure account' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Invoke-WebRequest -MockWith {
                    return @{ Content = '{ "token_endpoint": "https://login.microsoftonline.com/t/oauth2/v2.0/token" }' }
                }
                Mock -CommandName Connect-AzAccount -MockWith { }
                Mock -CommandName Disconnect-AzAccount -MockWith { }
                Mock -CommandName Get-AzContext -MockWith {
                    return @{ Environment = @{ ResourceManagerUrl = 'https://management.azure.com/' } }
                }

                Connect-M365Tenant -Workload 'Azure' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -ApplicationSecret 'super-secret'

                Reset-MSCloudLoginConnectionProfileContext -Workload 'Azure'

                Should -Invoke Disconnect-AzAccount -Exactly 1
                (Get-MSCloudLoginConnectionProfile -Workload 'Azure').Connected | Should -BeFalse
            }
        }
    }
}
