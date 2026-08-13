#Requires -Modules Pester

Describe 'Connect-MSCloudLoginSharePointOnlineREST' {
    BeforeAll {
        Import-Module ./Modules/MSCloudLoginAssistant/MSCloudLoginAssistant.psd1 -Force
    }

    Context 'When connecting with ServicePrincipalWithThumbprint' {
        It 'Should call Connect-MSCloudLoginRESTWorkload with correct parameters' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MSCloudLoginRESTWorkload -MockWith { return @{} }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.AuthenticationType = 'ServicePrincipalWithThumbprint'
                $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.AuthorizationUrl = 'https://login.microsoftonline.com'
                $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.Scope = 'https://contoso.sharepoint.com/.default'
                $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.CertificateThumbprint = 'thumbprint'

                Connect-MSCloudLoginSharePointOnlineREST

                Should -Invoke Connect-MSCloudLoginRESTWorkload -ParameterFilter {
                    $WorkloadName -eq 'SharePointOnlineREST' -and
                    $AuthorizationUrl -eq 'https://login.microsoftonline.com' -and
                    $Scope -eq 'https://contoso.sharepoint.com/.default' -and
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
                $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.AuthenticationType = 'Interactive'

                { Connect-MSCloudLoginSharePointOnlineREST } | Should -Throw "*Authentication method 'Interactive' is not supported for workload 'SharePointOnlineREST'*"
            }
        }
    }
}

Describe 'Disconnect-MSCloudLoginSharePointOnlineREST' {
    Context 'When SharePointOnlineREST is connected' {
        It 'Should call Disconnect-MSCloudLoginRESTWorkload' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Disconnect-MSCloudLoginRESTWorkload -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.Connected = $true

                Disconnect-MSCloudLoginSharePointOnlineREST

                Should -Invoke Disconnect-MSCloudLoginRESTWorkload -ParameterFilter {
                    $WorkloadName -eq 'SharePointOnlineREST'
                }
            }
        }
    }

    Context 'When SharePointOnlineREST is not connected' {
        It 'Should not throw and log message' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST.Connected = $false

                { Disconnect-MSCloudLoginSharePointOnlineREST } | Should -Not -Throw
            }
        }
    }
}

AfterAll {
    Remove-Module MSCloudLoginAssistant
}
