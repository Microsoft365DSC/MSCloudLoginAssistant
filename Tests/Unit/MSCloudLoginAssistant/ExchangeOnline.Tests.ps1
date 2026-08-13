#Requires -Modules Pester

Describe 'Connect-MSCloudLoginExchangeOnline' {
    BeforeAll {
        Import-Module ./Modules/MSCloudLoginAssistant/MSCloudLoginAssistant.psd1 -Force
    }

    Context 'When connecting with ServicePrincipalWithThumbprint' {
        It 'Should call Connect-ExchangeOnline with correct parameters' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-ExchangeOnline -MockWith { }
                Mock -CommandName Get-ConnectionInformation -MockWith { return @() }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-MSCloudLoginSPOUrlFromTenantId -MockWith { return @{ ConnectionUrl = 'https://outlook.office365.com/powershell-liveid/'; AdminUrl = 'https://outlook.office365.com/powershell-liveid/' } }
                Mock -CommandName Get-MSCloudLoginCertificate -MockWith { return New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.AuthenticationType = 'ServicePrincipalWithThumbprint'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.TenantId = 'tenant-id.onmicrosoft.com'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CertificateThumbprint = 'thumbprint'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ExchangeEnvironmentName = 'O365Default'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected = $false

                Connect-MSCloudLoginExchangeOnline

                Should -Invoke Connect-ExchangeOnline -ParameterFilter {
                    $AppId -eq 'app-id' -and
                    $Organization -eq 'tenant-id.onmicrosoft.com' -and
                    $CertificateThumbprint -eq 'thumbprint'
                }
            }
        }
    }

    Context 'When connecting with ServicePrincipalWithSecret' {
        It 'Should throw error for unsupported type' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.AuthenticationType = 'ServicePrincipalWithSecret'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.TenantId = 'tenant-id.onmicrosoft.com'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ApplicationSecret = 'secret'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ExchangeEnvironmentName = 'O365Default'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected = $false

                { Connect-MSCloudLoginExchangeOnline } | Should -Throw '*No valid authentication type found*'
            }
        }
    }

    Context 'When connecting with Credentials' {
        It 'Should call Connect-ExchangeOnline with correct parameters' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-ExchangeOnline -MockWith { }
                Mock -CommandName Get-ConnectionInformation -MockWith { return @() }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-MSCloudLoginSPOUrlFromTenantId -MockWith { return @{ ConnectionUrl = 'https://outlook.office365.com/powershell-liveid/'; AdminUrl = 'https://outlook.office365.com/powershell-liveid/' } }

                $secPwd = ConvertTo-SecureString 'password' -AsPlainText -Force
                $cred = New-Object PSCredential ('user@contoso.com', $secPwd)

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.AuthenticationType = 'Credentials'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Credentials = $cred
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ExchangeEnvironmentName = 'O365Default'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected = $false

                Connect-MSCloudLoginExchangeOnline

                Should -Invoke Connect-ExchangeOnline -ParameterFilter {
                    $Credential.UserName -eq 'user@contoso.com'
                }
            }
        }
    }

Context 'When connecting with AccessTokens' {
        It 'Should call Connect-ExchangeOnline with AccessToken' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-ExchangeOnline -MockWith { }
                Mock -CommandName Get-ConnectionInformation -MockWith { return @() }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-MSCloudLoginSPOUrlFromTenantId -MockWith { return @{ ConnectionUrl = 'https://outlook.office365.com/powershell-liveid/'; AdminUrl = 'https://outlook.office365.com/powershell-liveid/' } }
                Mock -CommandName Get-MSCloudLoginAccessTokenValue -MockWith { return 'token123' }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.AuthenticationType = 'AccessTokens'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.AccessTokens = @(@{ access_token = 'token123' })
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.TenantId = 'tenant-id.onmicrosoft.com'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ExchangeEnvironmentName = 'O365Default'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected = $false

                Connect-MSCloudLoginExchangeOnline

                Should -Invoke Connect-ExchangeOnline -ParameterFilter {
                    $AccessToken -eq 'token123' -and
                    $Organization -eq 'tenant-id.onmicrosoft.com'
                }
            }
        }
    }

    Context 'When connecting with Identity' {
        It 'Should call Connect-ExchangeOnline with ManagedIdentity' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-ExchangeOnline -MockWith { }
                Mock -CommandName Get-ConnectionInformation -MockWith { return @() }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-MSCloudLoginSPOUrlFromTenantId -MockWith { return @{ ConnectionUrl = 'https://outlook.office365.com/powershell-liveid/'; AdminUrl = 'https://outlook.office365.com/powershell-liveid/' } }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.AuthenticationType = 'Identity'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.TenantId = 'tenant-id.onmicrosoft.com'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ExchangeEnvironmentName = 'O365Default'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected = $false

                Connect-MSCloudLoginExchangeOnline

                Should -Invoke Connect-ExchangeOnline -ParameterFilter {
                    $ManagedIdentity -eq $true -and
                    $AppId -eq 'app-id' -and
                    $Organization -eq 'tenant-id.onmicrosoft.com'
                }
            }
        }
    }

    Context 'When unsupported authentication type is provided' {
        It 'Should throw an error' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.AuthenticationType = 'Interactive'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected = $false

                { Connect-MSCloudLoginExchangeOnline } | Should -Throw '*No valid authentication type found*'
            }
        }
    }
}

Describe 'Disconnect-MSCloudLoginExchangeOnline' {
    Context 'When ExchangeOnline is connected' {
        It 'Should call Disconnect-ExchangeOnline' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Disconnect-ExchangeOnline -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Get-Module -MockWith { return @() }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected = $true

                Disconnect-MSCloudLoginExchangeOnline

                Should -Invoke Disconnect-ExchangeOnline
            }
        }
    }

    Context 'When ExchangeOnline is not connected' {
        It 'Should not throw and log message' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected = $false

                { Disconnect-MSCloudLoginExchangeOnline } | Should -Not -Throw
            }
        }
    }
}

AfterAll {
    Remove-Module MSCloudLoginAssistant
}
