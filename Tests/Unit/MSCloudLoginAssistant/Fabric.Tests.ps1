#Requires -Modules Pester

Describe 'Connect-MSCloudLoginFabric' {
    BeforeAll {
        Import-Module ./Modules/MSCloudLoginAssistant/MSCloudLoginAssistant.psd1 -Force
    }

    Context 'When connecting with ServicePrincipalWithThumbprint' {
        It 'Should call Connect-MSCloudLoginRESTWorkload with correct parameters' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MSCloudLoginRESTWorkload -MockWith { return @{} }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Fabric.AuthenticationType = 'ServicePrincipalWithThumbprint'
                $Script:MSCloudLoginConnectionProfile.Fabric.AuthorizationUrl = 'https://login.microsoftonline.com'
                $Script:MSCloudLoginConnectionProfile.Fabric.Scope = 'https://analysis.windows.net/powerbi/api/.default'
                $Script:MSCloudLoginConnectionProfile.Fabric.ApplicationId = 'app-id'

                Connect-MSCloudLoginFabric

                Should -Invoke Connect-MSCloudLoginRESTWorkload -ParameterFilter {
                    $WorkloadName -eq 'Fabric' -and
                    $AuthorizationUrl -eq 'https://login.microsoftonline.com' -and
                    $Scope -eq 'https://analysis.windows.net/powerbi/api/.default' -and
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
                $Script:MSCloudLoginConnectionProfile.Fabric.AuthenticationType = 'Interactive'

                { Connect-MSCloudLoginFabric } | Should -Throw "*Authentication method 'Interactive' is not supported for workload 'Fabric'*"
            }
        }
    }
}

Describe 'Disconnect-MSCloudLoginFabric' {
    Context 'When Fabric is connected' {
        It 'Should call Disconnect-MSCloudLoginRESTWorkload' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Disconnect-MSCloudLoginRESTWorkload -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Fabric.Connected = $true

                Disconnect-MSCloudLoginFabric

                Should -Invoke Disconnect-MSCloudLoginRESTWorkload -ParameterFilter {
                    $WorkloadName -eq 'Fabric'
                }
            }
        }
    }

    Context 'When Fabric is not connected' {
        It 'Should not throw and log message' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Fabric.Connected = $false

                { Disconnect-MSCloudLoginFabric } | Should -Not -Throw
            }
        }
    }
}

AfterAll {
    Remove-Module MSCloudLoginAssistant
}
