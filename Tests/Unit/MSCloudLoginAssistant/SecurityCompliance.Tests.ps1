#Requires -Modules Pester

Describe 'Connect-MSCloudLoginSecurityCompliance' {
    BeforeAll {
        Import-Module ./Modules/MSCloudLoginAssistant/MSCloudLoginAssistant.psd1 -Force
    }

    Context 'When connecting with ServicePrincipalWithThumbprint' {
        It 'Should call Connect-IPPSSession with correct parameters' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-IPPSSession -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Remove-MSCloudLoginProxyModule -MockWith { }
                Mock -CommandName Get-Module -MockWith { return @() }
                Mock -CommandName Get-PSSession -MockWith { return @() }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.AuthenticationType = 'ServicePrincipalWithThumbprint'
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.TenantId = 'tenant-id'
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.CertificateThumbprint = 'thumbprint'
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ConnectionUrl = 'https://ps.compliance.protection.outlook.com/powershell-liveid/'
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.AzureADAuthorizationEndpointUri = 'https://login.microsoftonline.com/organizations'
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.EnableSearchOnlySession = $false
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Connected = $false

                Connect-MSCloudLoginSecurityCompliance

                Should -Invoke Connect-IPPSSession -ParameterFilter {
                    $AppId -eq 'app-id' -and
                    $Organization -eq 'tenant-id' -and
                    $CertificateThumbprint -eq 'thumbprint'
                }
            }
        }
    }

    Context 'When connecting with ServicePrincipalWithPath' {
        It 'Should call Connect-IPPSSession with CertificatePath' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-IPPSSession -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Remove-MSCloudLoginProxyModule -MockWith { }
                Mock -CommandName Get-Module -MockWith { return @() }
                Mock -CommandName Get-PSSession -MockWith { return @() }

                $secPwd = ConvertTo-SecureString 'password' -AsPlainText -Force

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.AuthenticationType = 'ServicePrincipalWithPath'
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.TenantId = 'tenant-id'
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.CertificatePath = 'C:\cert.pfx'
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.CertificatePassword = $secPwd
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ConnectionUrl = 'https://ps.compliance.protection.outlook.com/powershell-liveid/'
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.AzureADAuthorizationEndpointUri = 'https://login.microsoftonline.com/organizations'
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.EnableSearchOnlySession = $false
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Connected = $false

                Connect-MSCloudLoginSecurityCompliance

                Should -Invoke Connect-IPPSSession -ParameterFilter {
                    $AppId -eq 'app-id' -and
                    $CertificateFilePath -eq 'C:\cert.pfx'
                }
            }
        }
    }

    Context 'When connecting with Credentials' {
        It 'Should call Connect-IPPSSession with Credentials' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-IPPSSession -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Remove-MSCloudLoginProxyModule -MockWith { }
                Mock -CommandName Get-Module -MockWith { return @() }
                Mock -CommandName Get-PSSession -MockWith { return @() }

                $secPwd = ConvertTo-SecureString 'password' -AsPlainText -Force
                $cred = New-Object PSCredential ('user@contoso.com', $secPwd)

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.AuthenticationType = 'Credentials'
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Credentials = $cred
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ConnectionUrl = 'https://ps.compliance.protection.outlook.com/powershell-liveid/'
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.AzureADAuthorizationEndpointUri = 'https://login.microsoftonline.com/organizations'
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.EnableSearchOnlySession = $false
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Connected = $false

                Connect-MSCloudLoginSecurityCompliance

                Should -Invoke Connect-IPPSSession -ParameterFilter {
                    $Credential -eq $cred
                }
            }
        }
    }

    Context 'When MFA is required with Credentials' {
        It 'Should call Connect-MSCloudLoginSecurityComplianceMFA' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-IPPSSession -MockWith { throw 'MFA required' }
                Mock -CommandName Connect-MSCloudLoginSecurityComplianceMFA -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Remove-MSCloudLoginProxyModule -MockWith { }
                Mock -CommandName Get-Module -MockWith { return @() }
                Mock -CommandName Get-PSSession -MockWith { return @() }
                Mock -CommandName Test-MSCloudLoginMFARequiredError -MockWith { return $true }
                Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $false }

                $secPwd = ConvertTo-SecureString 'password' -AsPlainText -Force
                $cred = New-Object PSCredential ('user@contoso.com', $secPwd)

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.AuthenticationType = 'Credentials'
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Credentials = $cred
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.ConnectionUrl = 'https://ps.compliance.protection.outlook.com/powershell-liveid/'
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.AzureADAuthorizationEndpointUri = 'https://login.microsoftonline.com/organizations'
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.EnableSearchOnlySession = $false
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Connected = $false

                Connect-MSCloudLoginSecurityCompliance

                Should -Invoke Connect-MSCloudLoginSecurityComplianceMFA
            }
        }
    }

    Context 'When unsupported authentication type is provided' {
        It 'Should throw an error' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Remove-MSCloudLoginProxyModule -MockWith { }
                Mock -CommandName Get-Module -MockWith { return @() }
                Mock -CommandName Get-PSSession -MockWith { return @() }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.AuthenticationType = 'Interactive'
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Connected = $false

                { Connect-MSCloudLoginSecurityCompliance } | Should -Throw '*not supported for workload*'
            }
        }
    }
}

Describe 'Disconnect-MSCloudLoginSecurityCompliance' {
    Context 'When SecurityComplianceCenter is connected' {
        It 'Should call Disconnect-ExchangeOnline' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Disconnect-ExchangeOnline -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Connected = $true

                Disconnect-MSCloudLoginSecurityCompliance

                Should -Invoke Disconnect-ExchangeOnline
            }
        }
    }

    Context 'When SecurityComplianceCenter is not connected' {
        It 'Should not throw and log message' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Connected = $false

                { Disconnect-MSCloudLoginSecurityCompliance } | Should -Not -Throw
            }
        }
    }
}

AfterAll {
    Remove-Module MSCloudLoginAssistant
}
