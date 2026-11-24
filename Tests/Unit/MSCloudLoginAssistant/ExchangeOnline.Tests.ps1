#Requires -Modules Pester

Describe 'Connect-MSCloudLoginExchangeOnline' {
    BeforeAll {
        Import-Module ./Modules/MSCloudLoginAssistant/MSCloudLoginAssistant.psd1 -Force
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

    Context 'When the workload connects by endpoint URIs' {
        BeforeEach {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-ExchangeOnline -MockWith { }
                Mock -CommandName Get-ConnectionInformation -MockWith { return @() }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Remove-MSCloudLoginProxyModule -MockWith { }
                Mock -CommandName Disconnect-ExchangeOnline -MockWith { }
                Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $false }
                Mock -CommandName Get-MSCloudLoginWindowsPrincipal -MockWith {
                    [System.Security.Principal.GenericPrincipal]::new(
                        [System.Security.Principal.GenericIdentity]::new('test-user'),
                        [System.String[]]@('Administrator'))
                }

                $Script:MSCloudLoginCurrentLoadedModule = ''
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $workloadProfile = $Script:MSCloudLoginConnectionProfile.ExchangeOnline
                $workloadProfile.ApplicationId = 'app-id'
                $workloadProfile.TenantId = 'contoso.onsovcloud.de'
                $workloadProfile.CertificateThumbprint = 'thumbprint'
                $workloadProfile.CertificatePath = 'C:\certs\contoso.pfx'
                $workloadProfile.CertificatePassword = ConvertTo-SecureString 'cert-password' -AsPlainText -Force
                $workloadProfile.AccessTokens = @('token')
                $workloadProfile.Credentials = New-Object PSCredential ('admin@contoso.onsovcloud.de', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))
                $workloadProfile.ExchangeEnvironmentName = 'Custom'
                $workloadProfile.ConnectionUri = 'https://outlook.sovcloud.de/PowerShell-LiveID'
                $workloadProfile.AzureADAuthorizationEndpointUri = 'https://login.sovcloud-identity.de/contoso.onsovcloud.de'
            }
        }

        It 'Should pass the endpoint URIs and not the Custom environment name for <AuthenticationType>' -TestCases @(
            @{ AuthenticationType = 'ServicePrincipalWithThumbprint' }
            @{ AuthenticationType = 'ServicePrincipalWithPath' }
            @{ AuthenticationType = 'Credentials' }
            @{ AuthenticationType = 'CredentialsWithApplicationId' }
            @{ AuthenticationType = 'CredentialsWithTenantId' }
            @{ AuthenticationType = 'Identity' }
            @{ AuthenticationType = 'AccessTokens' }
            @{ AuthenticationType = 'Interactive' }
        ) {
            param ($AuthenticationType)
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{ AuthenticationType = $AuthenticationType } {
                param ($AuthenticationType)

                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.AuthenticationType = $AuthenticationType

                # Thumbprint authentication is limited to Windows.
                $originalPlatform = $PSVersionTable['Platform']
                try
                {
                    $PSVersionTable['Platform'] = 'Win32NT'
                    Connect-MSCloudLoginExchangeOnline
                }
                finally
                {
                    $PSVersionTable['Platform'] = $originalPlatform
                }

                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected | Should -BeTrue
                Should -Invoke Connect-ExchangeOnline -Exactly 1 -ParameterFilter {
                    $ConnectionUri -eq 'https://outlook.sovcloud.de/PowerShell-LiveID' -and
                    $AzureADAuthorizationEndpointUri -eq 'https://login.sovcloud-identity.de/contoso.onsovcloud.de' -and
                    -not $PesterBoundParameters.ContainsKey('ExchangeEnvironmentName')
                }
            }
        }

        It 'Should pass the onsovcloud tenant as the organization of a certificate path connection' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.AuthenticationType = 'ServicePrincipalWithPath'

                Connect-MSCloudLoginExchangeOnline

                Should -Invoke Connect-ExchangeOnline -Exactly 1 -ParameterFilter {
                    $Organization -eq 'contoso.onsovcloud.de' -and $CertificateFilePath -eq 'C:\certs\contoso.pfx'
                }
            }
        }

        It 'Should throw before connecting when the Custom environment has no ConnectionUri' {
            InModuleScope 'MSCloudLoginAssistant' {
                $workloadProfile = $Script:MSCloudLoginConnectionProfile.ExchangeOnline
                $workloadProfile.AuthenticationType = 'Credentials'
                $workloadProfile.ConnectionUri = $null

                { Connect-MSCloudLoginExchangeOnline } | Should -Throw '*Set CustomEXOConnectionUri*'
                $workloadProfile.Connected | Should -BeFalse
                Should -Invoke Connect-ExchangeOnline -Exactly 0
            }
        }
    }

    Context 'When connecting with CredentialsWithApplicationId' {
        BeforeEach {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Get-ConnectionInformation -MockWith { return @() }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Remove-MSCloudLoginProxyModule -MockWith { }
                Mock -CommandName Disconnect-ExchangeOnline -MockWith { }
                Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $false }

                $Script:MSCloudLoginCurrentLoadedModule = ''
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $workloadProfile = $Script:MSCloudLoginConnectionProfile.ExchangeOnline
                $workloadProfile.AuthenticationType = 'CredentialsWithApplicationId'
                $workloadProfile.ApplicationId = 'app-id'
                $workloadProfile.Credentials = New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))
            }
        }

        It 'Should pass the credential and the application id' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-ExchangeOnline -MockWith { }

                Connect-MSCloudLoginExchangeOnline

                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected | Should -BeTrue
                Should -Invoke Connect-ExchangeOnline -Exactly 1 -ParameterFilter {
                    $Credential.UserName -eq 'admin@contoso.onmicrosoft.com' -and
                    $AppId -eq 'app-id' -and
                    $ExchangeEnvironmentName -eq 'O365Default'
                }
            }
        }

        It 'Should keep the application id when it falls back to the MFA sign-in' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-ExchangeOnline -MockWith {
                    if ($null -ne $Credential)
                    {
                        throw 'AADSTS50076: multi-factor authentication is required.'
                    }
                }

                Connect-MSCloudLoginExchangeOnline

                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.MultiFactorAuthentication | Should -BeTrue
                Should -Invoke Connect-ExchangeOnline -Exactly 1 -ParameterFilter {
                    $UserPrincipalName -eq 'admin@contoso.onmicrosoft.com' -and $AppId -eq 'app-id'
                }
            }
        }
    }

    Context 'When connecting with an interactive sign-in' {
        BeforeEach {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-ExchangeOnline -MockWith { }
                Mock -CommandName Get-ConnectionInformation -MockWith { return @() }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Remove-MSCloudLoginProxyModule -MockWith { }
                Mock -CommandName Disconnect-ExchangeOnline -MockWith { }

                $Script:MSCloudLoginCurrentLoadedModule = ''
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.AuthenticationType = 'Interactive'
            }
        }

        It 'Should connect without any credential and flag the session as multi-factor authenticated' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $false }

                Connect-MSCloudLoginExchangeOnline

                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected | Should -BeTrue
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.MultiFactorAuthentication | Should -BeTrue
                Should -Invoke Connect-ExchangeOnline -Exactly 1 -ParameterFilter {
                    $null -eq $Credential -and
                    [System.String]::IsNullOrEmpty($AppId) -and
                    $ExchangeEnvironmentName -eq 'O365Default'
                }
            }
        }

        It 'Should refuse to sign in from a non-interactive shell' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $true }

                { Connect-MSCloudLoginExchangeOnline } | Should -Throw '*requires an interactive shell*'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected | Should -BeFalse
                Should -Invoke Connect-ExchangeOnline -Exactly 0
            }
        }
    }
}

Describe 'Get-MSCloudLoginExchangeOnlineConnectionParameters' {
    BeforeAll {
        Import-Module ./Modules/MSCloudLoginAssistant/MSCloudLoginAssistant.psd1 -Force
    }

    It 'Should select the environment name and pass the cmdlets to load' {
        InModuleScope 'MSCloudLoginAssistant' {
            $workloadProfile = New-Object ExchangeOnline
            $workloadProfile.ExchangeEnvironmentName = 'O365USGovGCCHigh'

            $result = Get-MSCloudLoginExchangeOnlineConnectionParameters -WorkloadProfile $workloadProfile -CommandName @{ CommandName = @('Get-Mailbox') }

            $result.ExchangeEnvironmentName | Should -Be 'O365USGovGCCHigh'
            $result.ContainsKey('ConnectionUri') | Should -BeFalse
            $result.CommandName | Should -Be @('Get-Mailbox')
            $result.ErrorAction | Should -Be 'Stop'
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
