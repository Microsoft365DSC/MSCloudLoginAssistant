#Requires -Modules Pester

Describe 'Connect-MSCloudLoginDefenderForEndpoint' {
    BeforeAll {
        Import-Module ./Modules/MSCloudLoginAssistant/MSCloudLoginAssistant.psd1 -Force
    }

    Context 'When connecting with ServicePrincipalWithThumbprint' {
        It 'Should call Connect-MSCloudLoginRESTWorkload with correct parameters' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MSCloudLoginRESTWorkload -MockWith { return @{} }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.DefenderForEndpoint.AuthenticationType = 'ServicePrincipalWithThumbprint'
                $Script:MSCloudLoginConnectionProfile.DefenderForEndpoint.AuthorizationUrl = 'https://login.microsoftonline.com'
                $Script:MSCloudLoginConnectionProfile.DefenderForEndpoint.Scope = 'https://api.securitycenter.microsoft.com/.default'
                $Script:MSCloudLoginConnectionProfile.DefenderForEndpoint.ApplicationId = 'app-id'

                Connect-MSCloudLoginDefenderForEndpoint

                Should -Invoke Connect-MSCloudLoginRESTWorkload -ParameterFilter {
                    $WorkloadName -eq 'DefenderForEndpoint' -and
                    $AuthorizationUrl -eq 'https://login.microsoftonline.com' -and
                    $Scope -eq 'https://api.securitycenter.microsoft.com/.default' -and
                    $ClientId -eq 'app-id'
                }
            }
        }
    }

Context 'When unsupported authentication type is provided' {
        It 'Should throw an error' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.DefenderForEndpoint.AuthenticationType = 'Interactive'

                { Connect-MSCloudLoginDefenderForEndpoint } | Should -Throw "*Authentication method 'Interactive' is not supported for workload 'DefenderForEndpoint'*"
            }
        }
    }
}

Describe 'Disconnect-MSCloudLoginDefenderForEndpoint' {
    Context 'When DefenderForEndpoint is connected' {
        It 'Should call Disconnect-MSCloudLoginRESTWorkload' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Disconnect-MSCloudLoginRESTWorkload -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.DefenderForEndpoint.Connected = $true

                Disconnect-MSCloudLoginDefenderForEndpoint

                Should -Invoke Disconnect-MSCloudLoginRESTWorkload -ParameterFilter {
                    $WorkloadName -eq 'DefenderForEndpoint'
                }
            }
        }
    }

    Context 'When DefenderForEndpoint is not connected' {
        It 'Should not throw and log message' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.DefenderForEndpoint.Connected = $false

                { Disconnect-MSCloudLoginDefenderForEndpoint } | Should -Not -Throw
            }
        }
    }
}

AfterAll {
    Remove-Module MSCloudLoginAssistant
}
