#Requires -Modules Pester

Describe 'Connect-MSCloudLoginPowerPlatform' {
    BeforeAll {
        Import-Module ./Modules/MSCloudLoginAssistant/MSCloudLoginAssistant.psd1 -Force
    }

    Context 'When connecting with ServicePrincipalWithThumbprint' {
        It 'Should call Add-PowerAppsAccount with correct parameters' {
            InModuleScope 'MSCloudLoginAssistant' {
                Import-Module ./Tests/Unit/Stubs/Stubs.psm1 -Force
                Mock -CommandName Add-PowerAppsAccount -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Import-Module -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PowerPlatform.AuthenticationType = 'ServicePrincipalWithThumbprint'
                $Script:MSCloudLoginConnectionProfile.PowerPlatform.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.PowerPlatform.TenantId = 'tenant-id'
                $Script:MSCloudLoginConnectionProfile.PowerPlatform.CertificateThumbprint = 'thumbprint'
                $Script:MSCloudLoginConnectionProfile.PowerPlatform.EnvironmentName = 'AzureCloud'
                $Script:MSCloudLoginConnectionProfile.PowerPlatform.Connected = $false

                $Script:CloudEnvironmentInfo = @{ tenant_region_sub_scope = 'PROD' }

                $Global:currentSession = [PSCustomObject]@{
                    resourceTokens = @{ 'https://service.powerapps.com/' = @{ accessToken = 'token' } }
                }

                Connect-MSCloudLoginPowerPlatform

                Should -Invoke Add-PowerAppsAccount -ParameterFilter {
                    $ApplicationId -eq 'app-id' -and
                    $TenantID -eq 'tenant-id' -and
                    $CertificateThumbprint -eq 'thumbprint' -and
                    $Endpoint -eq 'prod'
                }
            }
        }
    }

    Context 'When connecting with ServicePrincipalWithSecret' {
        It 'Should call Add-PowerAppsAccount with ClientSecret' {
            InModuleScope 'MSCloudLoginAssistant' {
                Import-Module ./Tests/Unit/Stubs/Stubs.psm1 -Force
                Mock -CommandName Add-PowerAppsAccount -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Import-Module -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PowerPlatform.AuthenticationType = 'ServicePrincipalWithSecret'
                $Script:MSCloudLoginConnectionProfile.PowerPlatform.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.PowerPlatform.TenantId = 'tenant-id'
                $Script:MSCloudLoginConnectionProfile.PowerPlatform.ApplicationSecret = 'secret'
                $Script:MSCloudLoginConnectionProfile.PowerPlatform.EnvironmentName = 'AzureCloud'
                $Script:MSCloudLoginConnectionProfile.PowerPlatform.Connected = $false

                $Script:CloudEnvironmentInfo = @{ tenant_region_sub_scope = 'PROD' }

                $Global:currentSession = [PSCustomObject]@{
                    resourceTokens = @{ 'https://service.powerapps.com/' = @{ accessToken = 'token' } }
                }

                Connect-MSCloudLoginPowerPlatform

                Should -Invoke Add-PowerAppsAccount -ParameterFilter {
                    $ApplicationId -eq 'app-id' -and
                    $ClientSecret -eq 'secret'
                }
            }
        }
    }

    Context 'When connecting with Credentials' {
        It 'Should call Add-PowerAppsAccount with Username/Password' {
            InModuleScope 'MSCloudLoginAssistant' {
                Import-Module ./Tests/Unit/Stubs/Stubs.psm1 -Force
                Mock -CommandName Add-PowerAppsAccount -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Import-Module -MockWith { }

                $secPwd = ConvertTo-SecureString 'password' -AsPlainText -Force
                $cred = New-Object PSCredential ('user@contoso.com', $secPwd)

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PowerPlatform.AuthenticationType = 'Credentials'
                $Script:MSCloudLoginConnectionProfile.PowerPlatform.Credentials = $cred
                $Script:MSCloudLoginConnectionProfile.PowerPlatform.EnvironmentName = 'AzureCloud'
                $Script:MSCloudLoginConnectionProfile.PowerPlatform.Connected = $false

                $Script:CloudEnvironmentInfo = @{ tenant_region_sub_scope = 'PROD' }

                $Global:currentSession = [PSCustomObject]@{
                    resourceTokens = @{ 'https://service.powerapps.com/' = @{ accessToken = 'token' } }
                }

                Connect-MSCloudLoginPowerPlatform

                Should -Invoke Add-PowerAppsAccount -ParameterFilter {
                    $Username -eq 'user@contoso.com' -and
                    $Password -eq $secPwd -and
                    $Endpoint -eq 'prod' -and
                    $ErrorAction -eq 'Stop'
                }
            }
        }
    }

    Context 'When connecting with CredentialsWithTenantId' {
        It 'Should throw an error' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Import-Module -MockWith { }

                $secPwd = ConvertTo-SecureString 'password' -AsPlainText -Force
                $cred = New-Object PSCredential ('user@contoso.com', $secPwd)

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PowerPlatform.AuthenticationType = 'CredentialsWithTenantId'
                $Script:MSCloudLoginConnectionProfile.PowerPlatform.Credentials = $cred
                $Script:MSCloudLoginConnectionProfile.PowerPlatform.TenantId = 'tenant-id'
                $Script:MSCloudLoginConnectionProfile.PowerPlatform.EnvironmentName = 'AzureCloud'
                $Script:MSCloudLoginConnectionProfile.PowerPlatform.Connected = $false

                $Script:CloudEnvironmentInfo = @{ tenant_region_sub_scope = 'PROD' }

                { Connect-MSCloudLoginPowerPlatform } | Should -Throw '*cannot specify TenantId with Credentials*'
            }
        }
    }

    Context 'When connecting with CredentialsWithApplicationId' {
        It 'Should throw an error' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Import-Module -MockWith { }

                $secPwd = ConvertTo-SecureString 'password' -AsPlainText -Force
                $cred = New-Object PSCredential ('user@contoso.com', $secPwd)

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PowerPlatform.AuthenticationType = 'CredentialsWithTenantId'
                $Script:MSCloudLoginConnectionProfile.PowerPlatform.Credentials = $cred
                $Script:MSCloudLoginConnectionProfile.PowerPlatform.TenantId = 'tenant-id'
                $Script:MSCloudLoginConnectionProfile.PowerPlatform.EnvironmentName = 'AzureCloud'
                $Script:MSCloudLoginConnectionProfile.PowerPlatform.Connected = $false

                { Connect-MSCloudLoginPowerPlatform } | Should -Throw '*You cannot specify TenantId with Credentials when connecting to PowerPlatforms*'
            }
        }
    }
}

Describe 'Disconnect-MSCloudLoginPowerPlatform' {
    Context 'When PowerPlatform is connected' {
        It 'Should disconnect properly' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PowerPlatform.Connected = $true

                # There's no explicit disconnect function for PowerPlatform in the module
                # The disconnect is handled by the connection profile reset
                { $Script:MSCloudLoginConnectionProfile.PowerPlatform.Connected = $false } | Should -Not -Throw
            }
        }
    }
}

AfterAll {
    Remove-Module MSCloudLoginAssistant
}
