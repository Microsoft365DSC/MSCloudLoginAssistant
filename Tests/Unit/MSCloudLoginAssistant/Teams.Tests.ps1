#Requires -Modules Pester

Describe 'Connect-MSCloudLoginTeams' {
    BeforeAll {
        # Plain function stubs keep command resolution off the real SDK modules,
        # which would otherwise be discovered and imported on first use.
        Import-Module ./Tests/Unit/Stubs/Stubs.psm1 -Force -Global -WarningAction SilentlyContinue
        Import-Module ./Modules/MSCloudLoginAssistant/MSCloudLoginAssistant.psd1 -Force

        # Compile and instantiate the workload classes once here so that the cost
        # does not show up inside the first test of this file.
        InModuleScope 'MSCloudLoginAssistant' {
            $null = New-Object MSCloudLoginConnectionProfile
        }
    }

    Context 'Get-MSCloudLoginTeamsEnvironmentParameters' {
        It 'Should return TeamsGCCH for AzureUSGovernment' {
            InModuleScope 'MSCloudLoginAssistant' {
                $result = Get-MSCloudLoginTeamsEnvironmentParameters -EnvironmentName 'AzureUSGovernment'
                $result.TeamsEnvironmentName | Should -Be 'TeamsGCCH'
            }
        }

        It 'Should return TeamsDOD for USGovernmentDoD' {
            InModuleScope 'MSCloudLoginAssistant' {
                $result = Get-MSCloudLoginTeamsEnvironmentParameters -EnvironmentName 'USGovernmentDoD'
                $result.TeamsEnvironmentName | Should -Be 'TeamsDOD'
            }
        }

        It 'Should return TeamsDOD for AzureDOD' {
            InModuleScope 'MSCloudLoginAssistant' {
                $result = Get-MSCloudLoginTeamsEnvironmentParameters -EnvironmentName 'AzureDOD'
                $result.TeamsEnvironmentName | Should -Be 'TeamsDOD'
            }
        }

        It 'Should return TeamsChina for AzureChinaCloud' {
            InModuleScope 'MSCloudLoginAssistant' {
                $result = Get-MSCloudLoginTeamsEnvironmentParameters -EnvironmentName 'AzureChinaCloud'
                $result.TeamsEnvironmentName | Should -Be 'TeamsChina'
            }
        }
    }

    Context 'When connecting with ServicePrincipalWithThumbprint' {
        It 'Should call Connect-MicrosoftTeams with AccessTokens when GraphScope and TeamsScope are set' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MicrosoftTeams -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Get-MSCloudLoginAccessToken -MockWith { return 'access-token' }
                Mock -CommandName Get-CsTeamsCallingPolicy -MockWith { throw 'No session' }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-MSCloudLoginCertificate -MockWith { return New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Teams.AuthenticationType = 'ServicePrincipalWithThumbprint'
                $Script:MSCloudLoginConnectionProfile.Teams.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.Teams.TenantId = 'tenant-id'
                $Script:MSCloudLoginConnectionProfile.Teams.CertificateThumbprint = 'thumbprint'
                $Script:MSCloudLoginConnectionProfile.Teams.GraphScope = 'https://graph.microsoft.com/.default'
                $Script:MSCloudLoginConnectionProfile.Teams.TeamsScope = 'https://teams.microsoft.com/.default'
                $Script:MSCloudLoginConnectionProfile.Teams.AuthorizationUrl = 'https://login.microsoftonline.com'
                $Script:MSCloudLoginConnectionProfile.Teams.TokenUrl = 'https://login.microsoftonline.com/organizations/oauth2/v2.0/token'
                $Script:MSCloudLoginConnectionProfile.Teams.Connected = $false
                $Script:MSCloudLoginConnectionProfile.Teams.EnvironmentName = 'AzureCloud'
                $Script:CustomEnvConfig.CustomEnvironment = $false
                $Script:CustomEnvConfig.CustomTeamsEndpoints = $null

                Connect-MSCloudLoginTeams

                Should -Invoke Connect-MicrosoftTeams -ParameterFilter {
                    $AccessTokens -like '*access-token*'
                }
            }
        }

        It 'Should keep only the tokens of the last connection' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MicrosoftTeams -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Get-MSCloudLoginAccessToken -MockWith { return 'access-token' }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Teams.AuthenticationType = 'ServicePrincipalWithThumbprint'
                $Script:MSCloudLoginConnectionProfile.Teams.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.Teams.TenantId = 'tenant-id'
                $Script:MSCloudLoginConnectionProfile.Teams.CertificateThumbprint = 'thumbprint'
                $Script:MSCloudLoginConnectionProfile.Teams.GraphScope = 'https://graph.microsoft.com/.default'
                $Script:MSCloudLoginConnectionProfile.Teams.TeamsScope = 'https://teams.microsoft.com/.default'
                $Script:MSCloudLoginConnectionProfile.Teams.AuthorizationUrl = 'https://login.microsoftonline.com'
                $Script:MSCloudLoginConnectionProfile.Teams.TokenUrl = 'https://login.microsoftonline.com/organizations/oauth2/v2.0/token'
                $Script:CustomEnvConfig.CustomEnvironment = $false
                $Script:CustomEnvConfig.CustomTeamsEndpoints = $null

                Connect-MSCloudLoginTeams
                Connect-MSCloudLoginTeams

                $Script:MSCloudLoginConnectionProfile.Teams.AccessTokens.Count | Should -Be 2
            }
        }

        It 'Should call Connect-MicrosoftTeams with CertificateThumbprint when GraphScope is not set and not custom env' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MicrosoftTeams -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Get-CsTeamsCallingPolicy -MockWith { throw 'No session' }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Teams.AuthenticationType = 'ServicePrincipalWithThumbprint'
                $Script:MSCloudLoginConnectionProfile.Teams.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.Teams.TenantId = 'tenant-id'
                $Script:MSCloudLoginConnectionProfile.Teams.CertificateThumbprint = 'thumbprint'
                # Do NOT set GraphScope, TeamsScope, TokenUrl - they default to $null which makes the condition false
                $Script:MSCloudLoginConnectionProfile.Teams.Connected = $false
                $Script:MSCloudLoginConnectionProfile.Teams.EnvironmentName = 'AzureCloud'
                $Script:CustomEnvConfig.CustomEnvironment = $false
                $Script:CustomEnvConfig.CustomTeamsEndpoints = $null

                Connect-MSCloudLoginTeams

                Should -Invoke Connect-MicrosoftTeams -ParameterFilter {
                    $ApplicationId -eq 'app-id' -and
                    $TenantId -eq 'tenant-id' -and
                    $CertificateThumbprint -eq 'thumbprint'
                }
            }
        }

        It 'Should handle connection failure for ServicePrincipalWithThumbprint (direct Connect-MicrosoftTeams path)' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MicrosoftTeams -MockWith { throw 'Connection failed' }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Get-CsTeamsCallingPolicy -MockWith { throw 'No session' }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Teams.AuthenticationType = 'ServicePrincipalWithThumbprint'
                $Script:MSCloudLoginConnectionProfile.Teams.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.Teams.TenantId = 'tenant-id'
                $Script:MSCloudLoginConnectionProfile.Teams.CertificateThumbprint = 'thumbprint'
                # Do NOT set GraphScope, TeamsScope, TokenUrl - they default to $null
                $Script:MSCloudLoginConnectionProfile.Teams.Connected = $false
                $Script:MSCloudLoginConnectionProfile.Teams.EnvironmentName = 'AzureCloud'
                $Script:CustomEnvConfig.CustomEnvironment = $false
                $Script:CustomEnvConfig.CustomTeamsEndpoints = $null

                { Connect-MSCloudLoginTeams } | Should -Throw 'Connection failed'
                $Script:MSCloudLoginConnectionProfile.Teams.Connected | Should -BeFalse
            }
        }
    }

    Context 'When connecting with a custom environment' {
        It 'Should reject custom environment connections outside Windows PowerShell 5' -Skip:($PSVersionTable.PSVersion.Major -eq 5) {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Get-CsTeamsCallingPolicy -MockWith { throw 'No session' }
                Mock -CommandName Connect-MicrosoftTeams -MockWith { }
                Mock -CommandName Set-TeamsEnvironmentConfig -MockWith { }

                $originalCustomEnvironment = $Script:CustomEnvConfig.CustomEnvironment
                $originalCustomTeamsEndpoints = $Script:CustomEnvConfig.CustomTeamsEndpoints
                try
                {
                    $Script:CustomEnvConfig.CustomEnvironment = $true
                    $Script:CustomEnvConfig.CustomTeamsEndpoints = @{ Teams = 'https://teams.example.test' }

                    $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                    $Script:MSCloudLoginConnectionProfile.Teams.AuthenticationType = 'ServicePrincipalWithThumbprint'
                    $Script:MSCloudLoginConnectionProfile.Teams.ApplicationId = 'app-id'
                    $Script:MSCloudLoginConnectionProfile.Teams.TenantId = 'contoso.onmicrosoft.com'
                    $Script:MSCloudLoginConnectionProfile.Teams.CertificateThumbprint = 'thumbprint'
                    $Script:MSCloudLoginConnectionProfile.Teams.Connected = $false

                    { Connect-MSCloudLoginTeams } |
                        Should -Throw '*only supported in PowerShell 5*'

                    Should -Invoke Set-TeamsEnvironmentConfig -Exactly 0
                    Should -Invoke Connect-MicrosoftTeams -Exactly 0
                }
                finally
                {
                    $Script:CustomEnvConfig.CustomEnvironment = $originalCustomEnvironment
                    $Script:CustomEnvConfig.CustomTeamsEndpoints = $originalCustomTeamsEndpoints
                }
            }
        }

        It 'Should configure and connect through the custom endpoints in Windows PowerShell 5' -Skip:($PSVersionTable.PSVersion.Major -gt 5) {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Get-CsTeamsCallingPolicy -MockWith { throw 'No session' }
                Mock -CommandName Connect-MicrosoftTeams -MockWith { }
                Mock -CommandName Set-TeamsEnvironmentConfig -MockWith { }

                $originalCustomEnvironment = $Script:CustomEnvConfig.CustomEnvironment
                $originalCustomTeamsEndpoints = $Script:CustomEnvConfig.CustomTeamsEndpoints
                try
                {
                    $Script:CustomEnvConfig.CustomEnvironment = $true
                    $Script:CustomEnvConfig.CustomTeamsEndpoints = @{ Teams = 'https://teams.example.test' }

                    $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                    $Script:MSCloudLoginConnectionProfile.Teams.AuthenticationType = 'ServicePrincipalWithThumbprint'
                    $Script:MSCloudLoginConnectionProfile.Teams.ApplicationId = 'app-id'
                    $Script:MSCloudLoginConnectionProfile.Teams.TenantId = 'contoso.onmicrosoft.com'
                    $Script:MSCloudLoginConnectionProfile.Teams.CertificateThumbprint = 'thumbprint'
                    $Script:MSCloudLoginConnectionProfile.Teams.Connected = $false

                    Connect-MSCloudLoginTeams

                    $Script:MSCloudLoginConnectionProfile.Teams.Connected | Should -BeTrue
                    Should -Invoke Set-TeamsEnvironmentConfig -Exactly 1 -ParameterFilter {
                        $EndpointUris.Teams -eq 'https://teams.example.test'
                    }
                    Should -Invoke Connect-MicrosoftTeams -Exactly 1 -ParameterFilter {
                        $ApplicationId -eq 'app-id' -and
                        $TenantId -eq 'contoso.onmicrosoft.com' -and
                        $CertificateThumbprint -eq 'thumbprint'
                    }
                }
                finally
                {
                    $Script:CustomEnvConfig.CustomEnvironment = $originalCustomEnvironment
                    $Script:CustomEnvConfig.CustomTeamsEndpoints = $originalCustomTeamsEndpoints
                }
            }
        }

        It 'Should call Connect-MicrosoftTeams with CertificateThumbprint in custom environment' -Skip:($PSVersionTable.PSVersion.Major -gt 5) {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Get-CsTeamsCallingPolicy -MockWith { throw 'No session' }
                Mock -CommandName Connect-MicrosoftTeams -MockWith { }
                Mock -CommandName Set-TeamsEnvironmentConfig -MockWith { }

                $originalCustomEnvironment = $Script:CustomEnvConfig.CustomEnvironment
                $originalCustomTeamsEndpoints = $Script:CustomEnvConfig.CustomTeamsEndpoints
                try
                {
                    $Script:CustomEnvConfig.CustomEnvironment = $true
                    $Script:CustomEnvConfig.CustomTeamsEndpoints = @{ Teams = 'https://teams.example.test' }

                    $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                    $Script:MSCloudLoginConnectionProfile.Teams.AuthenticationType = 'ServicePrincipalWithThumbprint'
                    $Script:MSCloudLoginConnectionProfile.Teams.ApplicationId = 'app-id'
                    $Script:MSCloudLoginConnectionProfile.Teams.TenantId = 'contoso.onmicrosoft.com'
                    $Script:MSCloudLoginConnectionProfile.Teams.CertificateThumbprint = 'thumbprint'
                    $Script:MSCloudLoginConnectionProfile.Teams.Connected = $false

                    Connect-MSCloudLoginTeams

                    Should -Invoke Connect-MicrosoftTeams -ParameterFilter {
                        $ApplicationId -eq 'app-id' -and
                        $TenantId -eq 'contoso.onmicrosoft.com' -and
                        $CertificateThumbprint -eq 'thumbprint'
                    }
                }
                finally
                {
                    $Script:CustomEnvConfig.CustomEnvironment = $originalCustomEnvironment
                    $Script:CustomEnvConfig.CustomTeamsEndpoints = $originalCustomTeamsEndpoints
                }
            }
        }
    }

    Context 'When connecting with ServicePrincipalWithPath' {
        It 'Should call Connect-MicrosoftTeams with CertificatePath' {
            InModuleScope 'MSCloudLoginAssistant' {
                $testCert = [Security.Cryptography.X509Certificates.X509Certificate2]::new()
                Mock -CommandName Connect-MicrosoftTeams -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Get-CsTeamsCallingPolicy -MockWith { throw 'No session' }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-MSCloudLoginCertificate -MockWith {
                    param ($CertificatePath, $CertificatePassword)
                    return $testCert
                }

                $secPwd = ConvertTo-SecureString 'password' -AsPlainText -Force

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Teams.AuthenticationType = 'ServicePrincipalWithPath'
                $Script:MSCloudLoginConnectionProfile.Teams.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.Teams.TenantId = 'tenant-id'
                $Script:MSCloudLoginConnectionProfile.Teams.CertificatePath = 'C:\cert.pfx'
                $Script:MSCloudLoginConnectionProfile.Teams.CertificatePassword = $secPwd
                $Script:MSCloudLoginConnectionProfile.Teams.Connected = $false
                $Script:MSCloudLoginConnectionProfile.Teams.EnvironmentName = 'AzureCloud'
                $Script:CustomEnvConfig.CustomEnvironment = $false
                $Script:CustomEnvConfig.CustomTeamsEndpoints = $null

                Connect-MSCloudLoginTeams

                Should -Invoke Connect-MicrosoftTeams -ParameterFilter {
                    $ApplicationId -eq 'app-id' -and
                    $TenantId -eq 'tenant-id' -and
                    $Certificate -eq $testCert
                }
            }
        }

        It 'Should stay disconnected when Connect-MicrosoftTeams writes an error' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MicrosoftTeams -MockWith { Write-Error -Message 'Connection failed' }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-MSCloudLoginCertificate -MockWith { return [Security.Cryptography.X509Certificates.X509Certificate2]::new() }
                Mock -CommandName Set-MSCloudLoginProcessConnectionIdentity -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Teams.AuthenticationType = 'ServicePrincipalWithPath'
                $Script:MSCloudLoginConnectionProfile.Teams.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.Teams.TenantId = 'tenant-id'
                $Script:MSCloudLoginConnectionProfile.Teams.CertificatePath = 'C:\cert.pfx'
                $Script:CustomEnvConfig.CustomEnvironment = $false
                $Script:CustomEnvConfig.CustomTeamsEndpoints = $null

                { Connect-MSCloudLoginTeams } | Should -Throw -ExpectedMessage '*Connection failed*'
                $Script:MSCloudLoginConnectionProfile.Teams.Connected | Should -BeFalse
                Should -Invoke Set-MSCloudLoginProcessConnectionIdentity -Exactly 0
            }
        }
    }

    Context 'When connecting with Credentials' {
        It 'Should call Connect-MicrosoftTeams with Credentials' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MicrosoftTeams -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Get-CsTeamsCallingPolicy -MockWith { throw 'No session' }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }

                $secPwd = ConvertTo-SecureString 'password' -AsPlainText -Force
                $cred = New-Object PSCredential ('user@contoso.com', $secPwd)

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Teams.AuthenticationType = 'Credentials'
                $Script:MSCloudLoginConnectionProfile.Teams.Credentials = $cred
                $Script:MSCloudLoginConnectionProfile.Teams.Connected = $false
                $Script:MSCloudLoginConnectionProfile.Teams.EnvironmentName = 'AzureCloud'
                $Script:CustomEnvConfig.CustomEnvironment = $false

                Connect-MSCloudLoginTeams

                Should -Invoke Connect-MicrosoftTeams -ParameterFilter {
                    $Credential -eq $cred
                }
            }
        }

        It 'Should handle connection failure for Credentials' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MicrosoftTeams -MockWith { throw 'Invalid credentials' }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Get-CsTeamsCallingPolicy -MockWith { throw 'No session' }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }

                $secPwd = ConvertTo-SecureString 'password' -AsPlainText -Force
                $cred = New-Object PSCredential ('user@contoso.com', $secPwd)

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Teams.AuthenticationType = 'Credentials'
                $Script:MSCloudLoginConnectionProfile.Teams.Credentials = $cred
                $Script:MSCloudLoginConnectionProfile.Teams.Connected = $false
                $Script:MSCloudLoginConnectionProfile.Teams.EnvironmentName = 'AzureCloud'
                $Script:CustomEnvConfig.CustomEnvironment = $false

                { Connect-MSCloudLoginTeams } | Should -Throw 'Invalid credentials'
                $Script:MSCloudLoginConnectionProfile.Teams.Connected | Should -BeFalse
            }
        }
    }

    Context 'When MFA is required with Credentials' {
        It 'Should call Connect-MSCloudLoginTeamsMFA' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MicrosoftTeams -MockWith { throw 'MFA required' }
                Mock -CommandName Connect-MSCloudLoginTeamsMFA -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Get-CsTeamsCallingPolicy -MockWith { throw 'No session' }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Test-MSCloudLoginMFARequiredError -MockWith { return $true }
                Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $false }

                $secPwd = ConvertTo-SecureString 'password' -AsPlainText -Force
                $cred = New-Object PSCredential ('user@contoso.com', $secPwd)

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Teams.AuthenticationType = 'Credentials'
                $Script:MSCloudLoginConnectionProfile.Teams.Credentials = $cred
                $Script:MSCloudLoginConnectionProfile.Teams.Connected = $false

                Connect-MSCloudLoginTeams

                Should -Invoke Connect-MSCloudLoginTeamsMFA
            }
        }
    }

    Context 'Connect-MSCloudLoginTeamsMFA' {
        It 'Should disconnect existing session and connect with MFA' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Get-MSCloudLoginTeamsEnvironmentParameters -MockWith { return @{} }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Disconnect-MicrosoftTeams -MockWith { }
                Mock -CommandName Connect-MicrosoftTeams -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Teams.EnvironmentName = 'AzureCloud'
                $Script:MSCloudLoginConnectionProfile.Teams.TenantId = 'tenant-id'
                $Script:MSCloudLoginConnectionProfile.Teams.Connected = $false

                Connect-MSCloudLoginTeamsMFA

                Should -Invoke Disconnect-MicrosoftTeams -Exactly 1
                Should -Invoke Connect-MicrosoftTeams -Exactly 1
                $Script:MSCloudLoginConnectionProfile.Teams.Connected | Should -BeTrue
                $Script:MSCloudLoginConnectionProfile.Teams.MultiFactorAuthentication | Should -BeTrue
            }
        }

        It 'Should handle error in MFA path' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Get-MSCloudLoginTeamsEnvironmentParameters -MockWith { return @{} }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Disconnect-MicrosoftTeams -MockWith { }
                Mock -CommandName Connect-MicrosoftTeams -MockWith { throw 'MFA failed' }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Teams.EnvironmentName = 'AzureCloud'
                $Script:MSCloudLoginConnectionProfile.Teams.TenantId = 'tenant-id'
                $Script:MSCloudLoginConnectionProfile.Teams.Connected = $false

                { Connect-MSCloudLoginTeamsMFA } | Should -Throw 'MFA failed'
                $Script:MSCloudLoginConnectionProfile.Teams.Connected | Should -BeFalse
            }
        }
    }

    Context 'When connecting with Identity' {
        It 'Should read the tenant from the managed identity token when no TenantId is set' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MicrosoftTeams -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Get-CsTeamsCallingPolicy -MockWith { throw 'No session' }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-AuthToken -MockWith {
                    $encode = { param ($Text) [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($Text)).TrimEnd('=').Replace('+', '-').Replace('/', '_') }
                    return '{0}.{1}.signature' -f (& $encode '{"alg":"none"}'), (& $encode '{"tid":"33333333-3333-3333-3333-333333333333"}')
                }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Teams.AuthenticationType = 'Identity'
                $Script:MSCloudLoginConnectionProfile.Teams.EnvironmentName = 'AzureUSGovernment'
                $Script:MSCloudLoginConnectionProfile.Teams.Connected = $false

                Connect-MSCloudLoginTeams

                Should -Invoke Get-AuthToken -Exactly 1 -ParameterFilter {
                    $Identity -and $Resource -eq 'https://graph.microsoft.us'
                }
                Should -Invoke Connect-MicrosoftTeams -ParameterFilter {
                    $Identity -eq $true -and $TenantId -eq '33333333-3333-3333-3333-333333333333'
                }
            }
        }

        It 'Should connect without TenantId when the managed identity token cannot be acquired' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MicrosoftTeams -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Get-CsTeamsCallingPolicy -MockWith { throw 'No session' }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-AuthToken -MockWith { throw 'No managed identity endpoint' }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Teams.AuthenticationType = 'Identity'
                $Script:MSCloudLoginConnectionProfile.Teams.Connected = $false

                Connect-MSCloudLoginTeams

                Should -Invoke Add-MSCloudLoginAssistantEvent -ParameterFilter { $EntryType -eq 'Warning' }
                Should -Invoke Connect-MicrosoftTeams -ParameterFilter {
                    $Identity -eq $true -and [System.String]::IsNullOrEmpty($TenantId)
                }
            }
        }

        It 'Should pass the resolved tenant GUID when a TenantId is set' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MicrosoftTeams -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Get-CsTeamsCallingPolicy -MockWith { throw 'No session' }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-MSCloudLoginTenantGuid -MockWith { return '22222222-2222-2222-2222-222222222222' }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Teams.AuthenticationType = 'Identity'
                $Script:MSCloudLoginConnectionProfile.Teams.TenantId = 'contoso.onmicrosoft.com'
                $Script:MSCloudLoginConnectionProfile.Teams.Connected = $false

                Connect-MSCloudLoginTeams

                Should -Invoke Get-MSCloudLoginTenantGuid -Exactly 1 -ParameterFilter {
                    $TenantId -eq 'contoso.onmicrosoft.com'
                }
                Should -Invoke Connect-MicrosoftTeams -ParameterFilter {
                    $Identity -eq $true -and $TenantId -eq '22222222-2222-2222-2222-222222222222'
                }
            }
        }

        It 'Should fall back to the TenantId when the tenant GUID cannot be resolved' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MicrosoftTeams -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Get-CsTeamsCallingPolicy -MockWith { throw 'No session' }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-MSCloudLoginTenantGuid -MockWith { return $null }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Teams.AuthenticationType = 'Identity'
                $Script:MSCloudLoginConnectionProfile.Teams.TenantId = 'contoso.onmicrosoft.com'
                $Script:MSCloudLoginConnectionProfile.Teams.Connected = $false

                Connect-MSCloudLoginTeams

                Should -Invoke Connect-MicrosoftTeams -ParameterFilter {
                    $Identity -eq $true -and $TenantId -eq 'contoso.onmicrosoft.com'
                }
            }
        }
    }

    Context 'When connecting with AccessTokens' {
        It 'Should call Connect-MicrosoftTeams with AccessTokens' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MicrosoftTeams -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Get-MSCloudLoginAccessTokenValue -MockWith { return 'token-value' }
                Mock -CommandName Get-CsTeamsCallingPolicy -MockWith { throw 'No session' }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Teams.AuthenticationType = 'AccessTokens'
                $Script:MSCloudLoginConnectionProfile.Teams.AccessTokens = @(@{ access_token = 'token123' })
                $Script:MSCloudLoginConnectionProfile.Teams.Connected = $false

                Connect-MSCloudLoginTeams

                Should -Invoke Connect-MicrosoftTeams -ParameterFilter {
                    $AccessTokens -like '*token-value*'
                }
            }
        }
    }

    Context 'When unsupported authentication type is provided' {
        It 'Should throw an error' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Get-CsTeamsCallingPolicy -MockWith { throw 'No session' }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Teams.AuthenticationType = 'Interactive'
                $Script:MSCloudLoginConnectionProfile.Teams.Connected = $false

                { Connect-MSCloudLoginTeams } | Should -Throw '*not supported for workload*'
            }
        }
    }

    Context 'When session is already usable' {
        It 'Should reuse existing connection without reconnecting' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $true }
                Mock -CommandName Connect-MicrosoftTeams -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Teams.CompleteConnection()

                Connect-MSCloudLoginTeams

                $Script:MSCloudLoginConnectionProfile.Teams.Connected | Should -BeTrue
                Should -Invoke Connect-MicrosoftTeams -Exactly 0
            }
        }
    }
}

Describe 'Disconnect-MSCloudLoginTeams' {
    Context 'When Teams is connected' {
        It 'Should call Disconnect-MicrosoftTeams' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Disconnect-MicrosoftTeams -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Teams.Connected = $true

                Disconnect-MSCloudLoginTeams

                Should -Invoke Disconnect-MicrosoftTeams
            }
        }
    }

    Context 'When Teams is not connected' {
        It 'Should not throw and log message' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Teams.Connected = $false

                { Disconnect-MSCloudLoginTeams } | Should -Not -Throw
            }
        }
    }
}

AfterAll {
    Remove-Module MSCloudLoginAssistant
}
