#Requires -Modules Pester

Describe 'Connect-MSCloudLoginAzureDevOPS' {
    BeforeAll {
        Import-Module ./Modules/MSCloudLoginAssistant/MSCloudLoginAssistant.psd1 -Force
    }

    Context 'When connecting with ServicePrincipalWithThumbprint' {
        It 'Should call Connect-MSCloudLoginRESTWorkload with correct parameters' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MSCloudLoginRESTWorkload -MockWith { return @{} }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.AzureDevOPS.AuthenticationType = 'ServicePrincipalWithThumbprint'
                $Script:MSCloudLoginConnectionProfile.AzureDevOPS.AuthorizationUrl = 'https://login.microsoftonline.com'
                $Script:MSCloudLoginConnectionProfile.AzureDevOPS.Scope = 'https://management.azure.com/.default'
                $Script:MSCloudLoginConnectionProfile.AzureDevOPS.ApplicationId = 'app-id'

                Connect-MSCloudLoginAzureDevOPS

                Should -Invoke Connect-MSCloudLoginRESTWorkload -ParameterFilter {
                    $WorkloadName -eq 'AzureDevOPS' -and
                    $AuthorizationUrl -eq 'https://login.microsoftonline.com' -and
                    $Scope -eq 'https://management.azure.com/.default' -and
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
                $Script:MSCloudLoginConnectionProfile.AzureDevOPS.AuthenticationType = 'Interactive'

                { Connect-MSCloudLoginAzureDevOPS } | Should -Throw "*Authentication method 'Interactive' is not supported for workload 'AzureDevOPS'*"
            }
        }
    }
}

Describe 'Disconnect-MSCloudLoginAzureDevOPS' {
    Context 'When AzureDevOPS is connected' {
        It 'Should call Disconnect-MSCloudLoginRESTWorkload' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Disconnect-MSCloudLoginRESTWorkload -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.AzureDevOPS.Connected = $true

                Disconnect-MSCloudLoginAzureDevOPS

                Should -Invoke Disconnect-MSCloudLoginRESTWorkload -ParameterFilter {
                    $WorkloadName -eq 'AzureDevOPS'
                }
            }
        }
    }

    Context 'When AzureDevOPS is not connected' {
        It 'Should not throw and log message' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.AzureDevOPS.Connected = $false

                { Disconnect-MSCloudLoginAzureDevOPS } | Should -Not -Throw
            }
        }
    }
}

AfterAll {
    Remove-Module MSCloudLoginAssistant
}
