#Requires -Modules Pester

Describe 'Connect-MSCloudLoginAdminAPI' {
    BeforeAll {
        Import-Module ./Modules/MSCloudLoginAssistant/MSCloudLoginAssistant.psd1 -Force
    }

    Context 'When connecting with ServicePrincipalWithThumbprint' {
        It 'Should call Connect-MSCloudLoginRESTWorkload with correct parameters' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MSCloudLoginRESTWorkload -MockWith { return @{} }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.AdminAPI.AuthenticationType = 'ServicePrincipalWithThumbprint'
                $Script:MSCloudLoginConnectionProfile.AdminAPI.AuthorizationUrl = 'https://login.microsoftonline.com'
                $Script:MSCloudLoginConnectionProfile.AdminAPI.Scope = 'https://graph.microsoft.com/.default'
                $Script:MSCloudLoginConnectionProfile.AdminAPI.ApplicationId = 'app-id'

                Connect-MSCloudLoginAdminAPI

                Should -Invoke Connect-MSCloudLoginRESTWorkload -ParameterFilter {
                    $WorkloadName -eq 'AdminAPI' -and
                    $AuthorizationUrl -eq 'https://login.microsoftonline.com' -and
                    $Scope -eq 'https://graph.microsoft.com/.default' -and
                    $ClientId -eq 'app-id'
                }
            }
        }
    }

    Context 'When connecting with ServicePrincipalWithSecret' {
        It 'Should call Connect-MSCloudLoginRESTWorkload with correct parameters' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MSCloudLoginRESTWorkload -MockWith { return @{} }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.AdminAPI.AuthenticationType = 'ServicePrincipalWithSecret'
                $Script:MSCloudLoginConnectionProfile.AdminAPI.AuthorizationUrl = 'https://login.microsoftonline.com'
                $Script:MSCloudLoginConnectionProfile.AdminAPI.Scope = 'https://graph.microsoft.com/.default'
                $Script:MSCloudLoginConnectionProfile.AdminAPI.ApplicationId = 'app-id'

                Connect-MSCloudLoginAdminAPI

                Should -Invoke Connect-MSCloudLoginRESTWorkload -ParameterFilter {
                    $WorkloadName -eq 'AdminAPI'
                }
            }
        }
    }

    Context 'When connecting with Credentials' {
        It 'Should call Connect-MSCloudLoginRESTWorkload with correct parameters' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MSCloudLoginRESTWorkload -MockWith { return @{} }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $secPwd = ConvertTo-SecureString 'password' -AsPlainText -Force
                $cred = New-Object PSCredential ('user@contoso.com', $secPwd)

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.AdminAPI.AuthenticationType = 'Credentials'
                $Script:MSCloudLoginConnectionProfile.AdminAPI.AuthorizationUrl = 'https://login.microsoftonline.com'
                $Script:MSCloudLoginConnectionProfile.AdminAPI.Scope = 'https://graph.microsoft.com/.default'
                $Script:MSCloudLoginConnectionProfile.AdminAPI.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.AdminAPI.Credentials = $cred

                Connect-MSCloudLoginAdminAPI

                Should -Invoke Connect-MSCloudLoginRESTWorkload -ParameterFilter {
                    $WorkloadName -eq 'AdminAPI'
                }
            }
        }
    }

Context 'When unsupported authentication type is provided' {
        It 'Should throw an error' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.AdminAPI.AuthenticationType = 'Interactive'

                { Connect-MSCloudLoginAdminAPI } | Should -Throw "*Authentication method 'Interactive' is not supported for workload 'AdminAPI'*"
            }
        }
    }
}

Describe 'Disconnect-MSCloudLoginAdminAPI' {
    Context 'When AdminAPI is connected' {
        It 'Should call Disconnect-MSCloudLoginRESTWorkload' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Disconnect-MSCloudLoginRESTWorkload -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.AdminAPI.Connected = $true

                Disconnect-MSCloudLoginAdminAPI

                Should -Invoke Disconnect-MSCloudLoginRESTWorkload -ParameterFilter {
                    $WorkloadName -eq 'AdminAPI'
                }
            }
        }
    }

    Context 'When AdminAPI is not connected' {
        It 'Should not throw and log message' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.AdminAPI.Connected = $false

                { Disconnect-MSCloudLoginAdminAPI } | Should -Not -Throw
            }
        }
    }
}

AfterAll {
    Remove-Module MSCloudLoginAssistant
}
