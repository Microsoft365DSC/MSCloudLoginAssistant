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

    Context 'When connecting with ServicePrincipalWithPath' {
        It 'Should connect when the session is elevated and derive the organization from the tenant id' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-ExchangeOnline -MockWith { }
                Mock -CommandName Get-ConnectionInformation -MockWith { return @() }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Remove-MSCloudLoginProxyModule -MockWith { }
                Mock -CommandName Disconnect-ExchangeOnline -MockWith { }
                Mock -CommandName Get-MSCloudLoginWindowsPrincipal -MockWith {
                    # A generic principal whose role list answers IsInRole without
                    # depending on the privileges of the test runner.
                    [System.Security.Principal.GenericPrincipal]::new(
                        [System.Security.Principal.GenericIdentity]::new('test-user'),
                        [System.String[]]@('Administrator'))
                }

                $certificatePassword = ConvertTo-SecureString 'cert-password' -AsPlainText -Force
                $Script:MSCloudLoginCurrentLoadedModule = ''
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.AuthenticationType = 'ServicePrincipalWithPath'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.TenantId = 'contoso.onmicrosoft.com'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CertificatePath = 'C:\certs\contoso.pfx'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CertificatePassword = $certificatePassword
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ExchangeEnvironmentName = 'O365Default'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected = $false

                Connect-MSCloudLoginExchangeOnline

                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected | Should -BeTrue
                $Script:MSCloudLoginConnectionProfile.OrganizationName | Should -Be 'contoso.onmicrosoft.com'
                Should -Invoke Connect-ExchangeOnline -Exactly 1 -ParameterFilter {
                    $AppId -eq 'app-id' -and
                    $Organization -eq 'contoso.onmicrosoft.com' -and
                    $CertificateFilePath -eq 'C:\certs\contoso.pfx' -and
                    $CertificatePassword -is [System.Security.SecureString]
                }
            }
        }

        It 'Should refuse an unelevated session and leave the workload disconnected' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-ExchangeOnline -MockWith { }
                Mock -CommandName Get-ConnectionInformation -MockWith { return @() }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Remove-MSCloudLoginProxyModule -MockWith { }
                Mock -CommandName Disconnect-ExchangeOnline -MockWith { }
                Mock -CommandName Get-MSCloudLoginWindowsPrincipal -MockWith {
                    [System.Security.Principal.GenericPrincipal]::new(
                        [System.Security.Principal.GenericIdentity]::new('test-user'),
                        [System.String[]]@('Users'))
                }

                $Script:MSCloudLoginCurrentLoadedModule = ''
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.AuthenticationType = 'ServicePrincipalWithPath'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.TenantId = 'contoso.onmicrosoft.com'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CertificatePath = 'C:\certs\contoso.pfx'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CertificatePassword = ConvertTo-SecureString 'cert-password' -AsPlainText -Force
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ExchangeEnvironmentName = 'O365Default'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected = $false

                { Connect-MSCloudLoginExchangeOnline } | Should -Throw '*requires the command to be run as Administrator*'

                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected | Should -BeFalse
                Should -Invoke Add-MSCloudLoginAssistantEvent -ParameterFilter {
                    $EntryType -eq 'Error' -and $Message -like '*Failed to connect to Exchange Online with Certificate Path*'
                }
            }
        }
    }

    Context 'When connecting with ServicePrincipalWithThumbprint on a non-Windows platform' {
        It 'Should refuse certificate thumbprint authentication outside of Windows' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-ExchangeOnline -MockWith { }
                Mock -CommandName Get-ConnectionInformation -MockWith { return @() }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }

                $Script:MSCloudLoginCurrentLoadedModule = ''
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.AuthenticationType = 'ServicePrincipalWithThumbprint'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.TenantId = 'tenant-id.onmicrosoft.com'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CertificateThumbprint = 'thumbprint'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ExchangeEnvironmentName = 'O365Default'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected = $false

                # Simulate a modern PowerShell running on a non-Windows platform.
                $originalPlatform = $PSVersionTable['Platform']
                $originalVersion = $PSVersionTable['PSVersion']
                try
                {
                    $PSVersionTable['Platform'] = 'Unix'
                    $PSVersionTable['PSVersion'] = [System.Version]'7.6.0'

                    { Connect-MSCloudLoginExchangeOnline } |
                        Should -Throw '*Certificate Thumbprint authentication is only supported on the Windows platform*'
                }
                finally
                {
                    if ($null -ne $originalPlatform)
                    {
                        $PSVersionTable['Platform'] = $originalPlatform
                    }
                    if ($null -ne $originalVersion)
                    {
                        $PSVersionTable['PSVersion'] = $originalVersion
                    }
                }

                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected | Should -BeFalse
            }
        }
    }

    Context 'When connecting with CredentialsWithTenantId' {
        It 'Should surface a non MFA sign-in failure and leave the workload disconnected' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-ExchangeOnline -MockWith { throw 'the account was disabled' }
                Mock -CommandName Get-ConnectionInformation -MockWith { return @() }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }

                $credential = New-Object PSCredential ('user@contoso.com', (ConvertTo-SecureString 'password' -AsPlainText -Force))

                $Script:MSCloudLoginCurrentLoadedModule = ''
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.AuthenticationType = 'CredentialsWithTenantId'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Credentials = $credential
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.TenantId = 'partner.contoso.com'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ExchangeEnvironmentName = 'O365Default'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected = $false

                { Connect-MSCloudLoginExchangeOnline } | Should -Throw '*account was disabled*'

                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected | Should -BeFalse
                Should -Invoke Add-MSCloudLoginAssistantEvent -ParameterFilter {
                    $EntryType -eq 'Error' -and $Message -like '*Failed to connect to Exchange Online with Credentials and TenantId*'
                }
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

Describe 'Get-MSCloudLoginWindowsPrincipal' {
    BeforeAll {
        Import-Module ./Modules/MSCloudLoginAssistant/MSCloudLoginAssistant.psd1 -Force
    }

    It 'Should return the principal of the current Windows user' -Skip:(-not $IsWindows) {
        InModuleScope 'MSCloudLoginAssistant' {
            $principal = Get-MSCloudLoginWindowsPrincipal
            $principal | Should -BeOfType [System.Security.Principal.WindowsPrincipal]
            $principal.Identity.Name | Should -Not -BeNullOrEmpty
        }
    }
}

AfterAll {
    Remove-Module MSCloudLoginAssistant -Force
}
