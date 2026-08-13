#Requires -Modules Pester

Describe 'Connect-MSCloudLoginAzure' {
    BeforeAll {
        Import-Module ./Modules/MSCloudLoginAssistant/MSCloudLoginAssistant.psd1 -Force
    }

    Context 'When connecting with ServicePrincipalWithThumbprint' {
        It 'Should call Connect-AzAccount with correct parameters' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-AzAccount -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Azure.AuthenticationType = 'ServicePrincipalWithThumbprint'
                $Script:MSCloudLoginConnectionProfile.Azure.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.Azure.TenantId = 'tenant-id'
                $Script:MSCloudLoginConnectionProfile.Azure.CertificateThumbprint = 'thumbprint'
                $Script:MSCloudLoginConnectionProfile.Azure.EnvironmentName = 'AzureCloud'
                $Script:MSCloudLoginConnectionProfile.Azure.Connected = $false

                Connect-MSCloudLoginAzure

                Should -Invoke Connect-AzAccount -ParameterFilter {
                    $ServicePrincipal -and
                    $ApplicationId -eq 'app-id' -and
                    $TenantId -eq 'tenant-id' -and
                    $CertificateThumbprint -eq 'thumbprint' -and
                    $Environment -eq 'AzureCloud'
                }
            }
        }
    }

    Context 'When connecting with ServicePrincipalWithSecret' {
        It 'Should call Connect-AzAccount with correct parameters' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-AzAccount -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Azure.AuthenticationType = 'ServicePrincipalWithSecret'
                $Script:MSCloudLoginConnectionProfile.Azure.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.Azure.TenantId = 'tenant-id'
                $Script:MSCloudLoginConnectionProfile.Azure.ApplicationSecret = 'secret'
                $Script:MSCloudLoginConnectionProfile.Azure.EnvironmentName = 'AzureCloud'
                $Script:MSCloudLoginConnectionProfile.Azure.Connected = $false

                Connect-MSCloudLoginAzure

                Should -Invoke Connect-AzAccount -ParameterFilter {
                    $ServicePrincipal -and
                    $Credential -ne $null
                }
            }
        }
    }

    Context 'When connecting with ServicePrincipalWithPath' {
        It 'Should call Connect-AzAccount with correct parameters' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-AzAccount -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }

                $secPwd = ConvertTo-SecureString 'password' -AsPlainText -Force

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Azure.AuthenticationType = 'ServicePrincipalWithPath'
                $Script:MSCloudLoginConnectionProfile.Azure.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.Azure.TenantId = 'tenant-id'
                $Script:MSCloudLoginConnectionProfile.Azure.CertificatePath = 'C:\cert.pfx'
                $Script:MSCloudLoginConnectionProfile.Azure.CertificatePassword = $secPwd
                $Script:MSCloudLoginConnectionProfile.Azure.EnvironmentName = 'AzureCloud'
                $Script:MSCloudLoginConnectionProfile.Azure.Connected = $false

                Connect-MSCloudLoginAzure

                Should -Invoke Connect-AzAccount -ParameterFilter {
                    $ServicePrincipal -and
                    $ApplicationId -eq 'app-id' -and
                    $CertificatePath -eq 'C:\cert.pfx'
                }
            }
        }
    }

    Context 'When connecting with Credentials' {
        It 'Should call Connect-AzAccount with correct parameters' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-AzAccount -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-MSCloudLoginTenantDomainFromCredentials -MockWith { return 'tenant.onmicrosoft.com' }

                $secPwd = ConvertTo-SecureString 'password' -AsPlainText -Force
                $cred = New-Object PSCredential ('user@contoso.com', $secPwd)

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Azure.AuthenticationType = 'Credentials'
                $Script:MSCloudLoginConnectionProfile.Azure.Credentials = $cred
                $Script:MSCloudLoginConnectionProfile.Azure.EnvironmentName = 'AzureCloud'
                $Script:MSCloudLoginConnectionProfile.Azure.Connected = $false

                Connect-MSCloudLoginAzure

                Should -Invoke Connect-AzAccount -ParameterFilter {
                    $Credential -eq $cred
                }
            }
        }
    }

    Context 'When MFA is required with Credentials' {
        It 'Should fall back to interactive login' {
            $callCount = 0
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-AzAccount -MockWith {
                    $script:callCount++
                    if ($script:callCount -eq 1) {
                        $errorRecord = [System.Management.Automation.ErrorRecord]::new(
                            [System.Exception]::new('AADSTS50076: MFA required'),
                            'MFARequired',
                            [System.Management.Automation.ErrorCategory]::AuthenticationError,
                            $null
                        )
                        throw $errorRecord
                    }
                }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Test-MSCloudLoginMFARequiredError -MockWith { return $true }
                Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $false }

                $secPwd = ConvertTo-SecureString 'password' -AsPlainText -Force
                $cred = New-Object PSCredential ('user@contoso.com', $secPwd)

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Azure.AuthenticationType = 'Credentials'
                $Script:MSCloudLoginConnectionProfile.Azure.Credentials = $cred
                $Script:MSCloudLoginConnectionProfile.Azure.TenantId = 'tenant-id'
                $Script:MSCloudLoginConnectionProfile.Azure.EnvironmentName = 'AzureCloud'
                $Script:MSCloudLoginConnectionProfile.Azure.Connected = $false

                Connect-MSCloudLoginAzure

                Should -Invoke Connect-AzAccount -ParameterFilter {
                    $TenantId -eq 'tenant-id' -and
                    $Environment -eq 'AzureCloud'
                } -Times 2
            }
        }
    }

    Context 'When connecting with AccessTokens' {
        It 'Should call Connect-AzAccount with AccessToken' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-AzAccount -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Azure.AuthenticationType = 'AccessTokens'
                $Script:MSCloudLoginConnectionProfile.Azure.AccessTokens = @('token123')
                $Script:MSCloudLoginConnectionProfile.Azure.TenantId = 'tenant-id'
                $Script:MSCloudLoginConnectionProfile.Azure.EnvironmentName = 'AzureCloud'
                $Script:MSCloudLoginConnectionProfile.Azure.Connected = $false

                Connect-MSCloudLoginAzure

                Should -Invoke Connect-AzAccount -ParameterFilter {
                    $AccessToken -eq 'token123'
                }
            }
        }
    }

    Context 'When connecting with Identity' {
        It 'Should call Connect-AzAccount with Identity' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-AzAccount -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Azure.AuthenticationType = 'Identity'
                $Script:MSCloudLoginConnectionProfile.Azure.EnvironmentName = 'AzureCloud'
                $Script:MSCloudLoginConnectionProfile.Azure.Connected = $false

                Connect-MSCloudLoginAzure

                Should -Invoke Connect-AzAccount -ParameterFilter {
                    $Identity -and
                    $Environment -eq 'AzureCloud'
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
                $Script:MSCloudLoginConnectionProfile.Azure.AuthenticationType = 'Interactive'
                $Script:MSCloudLoginConnectionProfile.Azure.Connected = $false

                { Connect-MSCloudLoginAzure } | Should -Throw '*Specified authentication method is not supported*'
            }
        }
    }
}

Describe 'Disconnect-MSCloudLoginAzure' {
    Context 'When Azure is connected' {
        It 'Should call Disconnect-AzAccount' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Disconnect-AzAccount -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Azure.Connected = $true

                Disconnect-MSCloudLoginAzure

                Should -Invoke Disconnect-AzAccount
            }
        }
    }

    Context 'When Azure is not connected' {
        It 'Should not throw and log message' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Azure.Connected = $false

                { Disconnect-MSCloudLoginAzure } | Should -Not -Throw
            }
        }
    }
}

AfterAll {
    Remove-Module MSCloudLoginAssistant
}
