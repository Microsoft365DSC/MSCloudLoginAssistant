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

    Context 'When connecting with ServicePrincipalWithThumbprint' {
        It 'Should connect with the Graph and Teams access tokens and keep only the tokens of the last connection' {
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
                Connect-MSCloudLoginTeams

                $Script:MSCloudLoginConnectionProfile.Teams.AccessTokens.Count | Should -Be 2
                Should -Invoke Connect-MicrosoftTeams -Exactly 2 -ParameterFilter {
                    $AccessTokens -like '*access-token*'
                }
            }
        }
    }

    Context 'When the connection fails' {
        It 'Should rethrow the <AuthenticationType> failure and stay disconnected without a connection identity' -TestCases @(
            @{
                AuthenticationType = 'ServicePrincipalWithThumbprint'
                Settings           = @{ ApplicationId = 'app-id'; TenantId = 'tenant-id'; CertificateThumbprint = 'thumbprint' }
                Terminating        = $true
                ExpectedError      = '*Connection failed*'
            }
            @{
                AuthenticationType = 'ServicePrincipalWithPath'
                Settings           = @{ ApplicationId = 'app-id'; TenantId = 'tenant-id'; CertificatePath = 'C:\cert.pfx' }
                Terminating        = $false
                ExpectedError      = '*Connection failed*'
            }
            @{
                AuthenticationType = 'Interactive'
                Settings           = @{}
                Terminating        = $true
                ExpectedError      = "*Authentication type 'Interactive' is not supported for workload 'MicrosoftTeams'*"
            }
        ) {
            InModuleScope 'MSCloudLoginAssistant' -Parameters $_ {
                param ($AuthenticationType, $Settings, $Terminating, $ExpectedError)

                Mock -CommandName Connect-MicrosoftTeams -MockWith {
                    if ($Terminating)
                    {
                        throw 'Connection failed'
                    }
                    Write-Error -Message 'Connection failed'
                }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-MSCloudLoginCertificate -MockWith { return [Security.Cryptography.X509Certificates.X509Certificate2]::new() }
                Mock -CommandName Set-MSCloudLoginProcessConnectionIdentity -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Teams.AuthenticationType = $AuthenticationType
                $Script:MSCloudLoginConnectionProfile.Teams.EnvironmentName = 'AzureCloud'
                foreach ($setting in $Settings.GetEnumerator())
                {
                    $Script:MSCloudLoginConnectionProfile.Teams.($setting.Key) = $setting.Value
                }
                $Script:CustomEnvConfig.CustomEnvironment = $false
                $Script:CustomEnvConfig.CustomTeamsEndpoints = $null

                { Connect-MSCloudLoginTeams } | Should -Throw -ExpectedMessage $ExpectedError
                $Script:MSCloudLoginConnectionProfile.Teams.Connected | Should -BeFalse
                Should -Invoke Set-MSCloudLoginProcessConnectionIdentity -Exactly 0
            }
        }
    }

    Context 'When connecting with a custom environment' {
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
    }

    Context 'Connect-MSCloudLoginTeamsMFA' {
        It 'Should disconnect the existing session and pass the government environment to the MFA sign-in' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Disconnect-MicrosoftTeams -MockWith { }
                Mock -CommandName Connect-MicrosoftTeams -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Teams.EnvironmentName = 'AzureUSGovernment'
                $Script:MSCloudLoginConnectionProfile.Teams.TenantId = 'contoso.onmicrosoft.com'
                $Script:MSCloudLoginConnectionProfile.Teams.Connected = $false

                Connect-MSCloudLoginTeamsMFA

                $Script:MSCloudLoginConnectionProfile.Teams.Connected | Should -BeTrue
                $Script:MSCloudLoginConnectionProfile.Teams.MultiFactorAuthentication | Should -BeTrue
                Should -Invoke Disconnect-MicrosoftTeams -Exactly 1
                Should -Invoke Connect-MicrosoftTeams -Exactly 1 -ParameterFilter {
                    $TeamsEnvironmentName -eq 'TeamsGCCH' -and $TenantId -eq 'contoso.onmicrosoft.com'
                }
            }
        }

        It 'Should rethrow and stay disconnected when the MFA sign-in fails' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Disconnect-MicrosoftTeams -MockWith { }
                Mock -CommandName Connect-MicrosoftTeams -MockWith { throw 'MFA failed' }

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

        It 'Should pass <ExpectedTenantId> as TenantId when the tenant GUID resolves to <ResolvedGuid>' -TestCases @(
            @{ ResolvedGuid = '22222222-2222-2222-2222-222222222222'; ExpectedTenantId = '22222222-2222-2222-2222-222222222222' }
            @{ ResolvedGuid = $null; ExpectedTenantId = 'contoso.onmicrosoft.com' }
        ) {
            param ($ResolvedGuid, $ExpectedTenantId)
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{ ResolvedGuid = $ResolvedGuid; ExpectedTenantId = $ExpectedTenantId } {
                param ($ResolvedGuid, $ExpectedTenantId)
                Mock -CommandName Connect-MicrosoftTeams -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Get-CsTeamsCallingPolicy -MockWith { throw 'No session' }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-MSCloudLoginTenantGuid -MockWith { return $ResolvedGuid }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Teams.AuthenticationType = 'Identity'
                $Script:MSCloudLoginConnectionProfile.Teams.TenantId = 'contoso.onmicrosoft.com'
                $Script:MSCloudLoginConnectionProfile.Teams.Connected = $false

                Connect-MSCloudLoginTeams

                Should -Invoke Get-MSCloudLoginTenantGuid -Exactly 1 -ParameterFilter {
                    $TenantId -eq 'contoso.onmicrosoft.com'
                }
                Should -Invoke Connect-MicrosoftTeams -ParameterFilter {
                    $Identity -eq $true -and $TenantId -eq $ExpectedTenantId
                }
            }
        }
    }
}

AfterAll {
    Remove-Module MSCloudLoginAssistant
}
