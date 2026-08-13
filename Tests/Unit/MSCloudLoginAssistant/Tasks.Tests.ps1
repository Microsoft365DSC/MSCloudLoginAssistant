#Requires -Modules Pester

Describe 'Connect-MSCloudLoginTasks' {
    BeforeAll {
        Import-Module ./Modules/MSCloudLoginAssistant/MSCloudLoginAssistant.psd1 -Force
    }

    Context 'When connecting with ServicePrincipalWithThumbprint' {
        It 'Should call Connect-MSCloudLoginRESTWorkload with correct parameters' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MSCloudLoginRESTWorkload -MockWith { return @{} }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Tasks.AuthenticationType = 'ServicePrincipalWithThumbprint'
                $Script:MSCloudLoginConnectionProfile.Tasks.AuthorizationUrl = 'https://login.microsoftonline.com'
                $Script:MSCloudLoginConnectionProfile.Tasks.Scope = 'https://graph.microsoft.com/.default'
                $Script:MSCloudLoginConnectionProfile.Tasks.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.Tasks.CertificateThumbprint = 'thumbprint'

                Connect-MSCloudLoginTasks

                Should -Invoke Connect-MSCloudLoginRESTWorkload -ParameterFilter {
                    $WorkloadName -eq 'Tasks' -and
                    $AuthorizationUrl -eq 'https://login.microsoftonline.com' -and
                    $Scope -eq 'https://graph.microsoft.com/.default' -and
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
                $Script:MSCloudLoginConnectionProfile.Tasks.AuthenticationType = 'Interactive'

                { Connect-MSCloudLoginTasks } | Should -Throw "*Authentication method 'Interactive' is not supported for workload 'Tasks'*"
            }
        }
    }
}

Describe 'Disconnect-MSCloudLoginTasks' {
    Context 'When Tasks is connected' {
        It 'Should call Disconnect-MSCloudLoginRESTWorkload' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Disconnect-MSCloudLoginRESTWorkload -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Tasks.Connected = $true

                Disconnect-MSCloudLoginTasks

                Should -Invoke Disconnect-MSCloudLoginRESTWorkload -ParameterFilter {
                    $WorkloadName -eq 'Tasks'
                }
            }
        }
    }

    Context 'When Tasks is not connected' {
        It 'Should not throw and log message' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Tasks.Connected = $false

                { Disconnect-MSCloudLoginTasks } | Should -Not -Throw
            }
        }
    }
}

AfterAll {
    Remove-Module MSCloudLoginAssistant
}
