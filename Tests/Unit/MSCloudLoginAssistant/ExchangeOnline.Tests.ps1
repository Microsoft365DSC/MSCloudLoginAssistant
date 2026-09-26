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
