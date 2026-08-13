#Requires -Modules Pester

Describe 'Connect-MSCloudLoginTeams' {
    BeforeAll {
        Import-Module ./Modules/MSCloudLoginAssistant/MSCloudLoginAssistant.psd1 -Force
    }

    Context 'When connecting with ServicePrincipalWithThumbprint' {
        It 'Should call Connect-MicrosoftTeams with AccessTokens' {
            InModuleScope 'MSCloudLoginAssistant' {
                Import-Module ./Tests/Unit/Stubs/Stubs.psm1 -Force
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
                $Script:MSCloudLoginConnectionProfile.Teams.EnvironmentName = 'Custom'
                $Script:CustomEnvConfig.CustomEnvironment = $false
                $Script:CustomEnvConfig.CustomTeamsEndpoints = $null

                Connect-MSCloudLoginTeams

                Should -Invoke Connect-MicrosoftTeams -ParameterFilter {
                    $AccessTokens -like '*access-token*'
                }
            }
        }
    }

    Context 'When connecting with ServicePrincipalWithPath' {
        It 'Should call Connect-MicrosoftTeams with CertificatePath' {
            InModuleScope 'MSCloudLoginAssistant' {
                Import-Module ./Tests/Unit/Stubs/Stubs.psm1 -Force
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
    }

    Context 'When connecting with Credentials' {
        It 'Should call Connect-MicrosoftTeams with Credentials' {
            InModuleScope 'MSCloudLoginAssistant' {
                Import-Module ./Tests/Unit/Stubs/Stubs.psm1 -Force
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
    }

    Context 'When MFA is required with Credentials' {
        It 'Should call Connect-MSCloudLoginTeamsMFA' {
            InModuleScope 'MSCloudLoginAssistant' {
                Import-Module ./Tests/Unit/Stubs/Stubs.psm1 -Force
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

    Context 'When connecting with Identity' {
        It 'Should call Connect-MicrosoftTeams with Identity' {
            InModuleScope 'MSCloudLoginAssistant' {
                Import-Module ./Tests/Unit/Stubs/Stubs.psm1 -Force
                Mock -CommandName Connect-MicrosoftTeams -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Get-CsTeamsCallingPolicy -MockWith { throw 'No session' }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Teams.AuthenticationType = 'Identity'
                $Script:MSCloudLoginConnectionProfile.Teams.Connected = $false

                Connect-MSCloudLoginTeams

                Should -Invoke Connect-MicrosoftTeams -ParameterFilter {
                    $Identity -eq $true
                }
            }
        }
    }

    Context 'When connecting with AccessTokens' {
        It 'Should call Connect-MicrosoftTeams with AccessTokens' {
            InModuleScope 'MSCloudLoginAssistant' {
                Import-Module ./Tests/Unit/Stubs/Stubs.psm1 -Force
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
                Import-Module ./Tests/Unit/Stubs/Stubs.psm1 -Force
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
}

Describe 'Disconnect-MSCloudLoginTeams' {
    Context 'When Teams is connected' {
        It 'Should call Disconnect-MicrosoftTeams' {
            InModuleScope 'MSCloudLoginAssistant' {
                Import-Module ./Tests/Unit/Stubs/Stubs.psm1 -Force
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
                Import-Module ./Tests/Unit/Stubs/Stubs.psm1 -Force
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
