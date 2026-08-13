#Requires -Modules Pester

Describe 'Connect-MSCloudLoginMicrosoftGraph' {
    BeforeAll {
        Import-Module ./Modules/MSCloudLoginAssistant/MSCloudLoginAssistant.psd1 -Force
    }

    Context 'When connecting with Credentials' {
        It 'Should call Connect-MSCloudLoginMSGraphWithUser' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MSCloudLoginMSGraphWithUser -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AuthenticationType = 'Credentials'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected = $false

                Connect-MSCloudLoginMicrosoftGraph

                Should -Invoke Connect-MSCloudLoginMSGraphWithUser
            }
        }
    }

    Context 'When connecting with CredentialsWithTenantId' {
        It 'Should call Connect-MSCloudLoginMSGraphWithUser with TenantId' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MSCloudLoginMSGraphWithUser -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AuthenticationType = 'CredentialsWithTenantId'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.TenantId = 'tenant-id'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected = $false

                Connect-MSCloudLoginMicrosoftGraph

                Should -Invoke Connect-MSCloudLoginMSGraphWithUser -ParameterFilter {
                    $TenantId -eq 'tenant-id'
                }
            }
        }
    }

    Context 'When connecting with Identity' {
        It 'Should call Connect-MgGraph with AccessToken' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MgGraph -MockWith { }
                Mock -CommandName Get-AuthToken -MockWith { return 'access-token' }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AuthenticationType = 'Identity'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ResourceUrl = 'https://graph.microsoft.com'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.GraphEnvironment = 'Global'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected = $false

                Connect-MSCloudLoginMicrosoftGraph

                Should -Invoke Connect-MgGraph -ParameterFilter {
                    $AccessToken -ne $null -and
                    $Environment -eq 'Global'
                }
            }
        }
    }

    Context 'When connecting with ServicePrincipalWithThumbprint' {
        It 'Should call Connect-MgGraph with Certificate' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MgGraph -MockWith { }
                Mock -CommandName Get-MSCloudLoginCertificate -MockWith { return New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AuthenticationType = 'ServicePrincipalWithThumbprint'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.TenantId = 'tenant-id'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.CertificateThumbprint = 'thumbprint'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.GraphEnvironment = 'Global'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected = $false

                Connect-MSCloudLoginMicrosoftGraph

                Should -Invoke Connect-MgGraph -ParameterFilter {
                    $ClientId -eq 'app-id' -and
                    $TenantId -eq 'tenant-id' -and
                    $Environment -eq 'Global'
                }
            }
        }
    }

    Context 'When connecting with ServicePrincipalWithSecret' {
        It 'Should call Connect-MgGraph with ClientSecret' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MgGraph -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AuthenticationType = 'ServicePrincipalWithSecret'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.TenantId = 'tenant-id'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ApplicationSecret = 'secret'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.GraphEnvironment = 'Global'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected = $false

                Connect-MSCloudLoginMicrosoftGraph

                Should -Invoke Connect-MgGraph -ParameterFilter {
                    $ClientSecretCredential -ne $null -and
                    $Environment -eq 'Global'
                }
            }
        }
    }

    Context 'When connecting with AccessTokens' {
        It 'Should call Connect-MgGraph with AccessToken' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MgGraph -MockWith { }
                Mock -CommandName Get-MSCloudLoginAccessTokenValue -MockWith { return 'token-value' }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AuthenticationType = 'AccessTokens'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AccessTokens = @(@{ access_token = 'token123' })
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.GraphEnvironment = 'Global'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected = $false

                Connect-MSCloudLoginMicrosoftGraph

                Should -Invoke Connect-MgGraph -ParameterFilter {
                    $AccessToken -ne $null -and
                    $Environment -eq 'Global'
                }
            }
        }
    }
}

Describe 'Disconnect-MSCloudLoginMicrosoftGraph' {
    Context 'When MicrosoftGraph is connected' {
        It 'Should call Disconnect-MgGraph' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Disconnect-MgGraph -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected = $true

                Disconnect-MSCloudLoginMicrosoftGraph

                Should -Invoke Disconnect-MgGraph
            }
        }
    }

    Context 'When MicrosoftGraph is not connected' {
        It 'Should not throw and log message' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected = $false

                { Disconnect-MSCloudLoginMicrosoftGraph } | Should -Not -Throw
            }
        }
    }
}

AfterAll {
    Remove-Module MSCloudLoginAssistant
}
