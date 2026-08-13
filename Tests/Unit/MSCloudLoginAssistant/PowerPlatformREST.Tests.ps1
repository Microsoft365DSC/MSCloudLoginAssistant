#Requires -Modules Pester

Describe 'Connect-MSCloudLoginPowerPlatformREST' {
    BeforeAll {
        Import-Module ./Modules/MSCloudLoginAssistant/MSCloudLoginAssistant.psd1 -Force
    }

    Context 'When connecting with ServicePrincipalWithThumbprint' {
        It 'Should call Connect-MSCloudLoginRESTWorkload with correct parameters' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MSCloudLoginRESTWorkload -MockWith { return @{} }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Invoke-WebRequest -MockWith { return @{} }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AuthenticationType = 'ServicePrincipalWithThumbprint'
                $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AuthorizationUrl = 'https://login.microsoftonline.com'
                $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.Scope = 'https://service.powerapps.com/.default'
                $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.ClientId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.CertificateThumbprint = 'thumbprint'
                $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.Connected = $false
                $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AccessToken = $null

                Connect-MSCloudLoginPowerPlatformREST

                Should -Invoke Connect-MSCloudLoginRESTWorkload -ParameterFilter {
                    $WorkloadName -eq 'PowerPlatformREST' -and
                    $AuthorizationUrl -eq 'https://login.microsoftonline.com' -and
                    $Scope -eq 'https://service.powerapps.com/.default' -and
                    $ClientId -eq 'app-id'
                }
            }
        }
    }

Context 'When token probe succeeds' {
        It 'Should still call Connect-MSCloudLoginRESTWorkload' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Invoke-WebRequest -MockWith { return @{ StatusCode = 200 } }
                Mock -CommandName Connect-MSCloudLoginRESTWorkload -MockWith { return @{} }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AuthenticationType = 'ServicePrincipalWithThumbprint'
                $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.BapEndpoint = 'bap.endpoint.com'
                $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AccessToken = 'Bearer token123'
                $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.Connected = $true

                Connect-MSCloudLoginPowerPlatformREST

                Should -Invoke Connect-MSCloudLoginRESTWorkload
            }
        }
    }

    Context 'When token probe fails' {
        It 'Should force reconnection' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Invoke-WebRequest -MockWith { throw '401 Unauthorized' }
                Mock -CommandName Connect-MSCloudLoginRESTWorkload -MockWith { return @{} }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AuthenticationType = 'ServicePrincipalWithThumbprint'
                $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.BapEndpoint = 'bap.endpoint.com'
                $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AccessToken = 'Bearer token123'
                $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.Connected = $true
                $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AuthorizationUrl = 'https://login.microsoftonline.com'
                $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.Scope = 'https://service.powerapps.com/.default'
                $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.ClientId = 'app-id'

                Connect-MSCloudLoginPowerPlatformREST

                Should -Invoke Connect-MSCloudLoginRESTWorkload
            }
        }
    }

Context 'When unsupported authentication type is provided' {
        It 'Should throw an error' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AuthenticationType = 'Interactive'

                { Connect-MSCloudLoginPowerPlatformREST } | Should -Throw "*Authentication method 'Interactive' is not supported for workload 'PowerPlatformREST'*"
            }
        }
    }
}

Describe 'Disconnect-MSCloudLoginPowerPlatformREST' {
    Context 'When PowerPlatformREST is connected' {
        It 'Should call Disconnect-MSCloudLoginRESTWorkload' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Disconnect-MSCloudLoginRESTWorkload -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.Connected = $true

                Disconnect-MSCloudLoginPowerPlatformREST

                Should -Invoke Disconnect-MSCloudLoginRESTWorkload -ParameterFilter {
                    $WorkloadName -eq 'PowerPlatformREST'
                }
            }
        }
    }

    Context 'When PowerPlatformREST is not connected' {
        It 'Should not throw and log message' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.Connected = $false

                { Disconnect-MSCloudLoginPowerPlatformREST } | Should -Not -Throw
            }
        }
    }
}

AfterAll {
    Remove-Module MSCloudLoginAssistant
}
