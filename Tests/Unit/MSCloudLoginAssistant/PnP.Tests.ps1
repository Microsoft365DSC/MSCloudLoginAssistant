#Requires -Modules Pester

Describe 'Connect-MSCloudLoginPnP' {
    BeforeAll {
        Import-Module ./Modules/MSCloudLoginAssistant/MSCloudLoginAssistant.psd1 -Force
    }

    Context 'When connecting with AccessTokens' {
        It 'Should call Connect-PnPOnline with AccessToken' {
            InModuleScope 'MSCloudLoginAssistant' {
                Import-Module ./Tests/Unit/Stubs/Stubs.psm1 -Force
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-Module -MockWith {
                    param ($Name, $ListAvailable)
                    if ($ListAvailable -and $Name -eq 'PnP.PowerShell') {
                        return @(
                            [pscustomobject]@{
                                Name = 'PnP.PowerShell'
                                Version = '1.10.0'
                                CompatiblePSEditions = @('Desktop', 'Core')
                            }
                        )
                    }
                    if ($Name -eq 'PnP.PowerShell') { return $null }
                    return @()
                }
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Get-MSCloudLoginSPOUrlFromTenantId -MockWith { return @{ ConnectionUrl = 'https://contoso.sharepoint.com'; AdminUrl = 'https://contoso-admin.sharepoint.com' } }

                $secPwd = ConvertTo-SecureString 'password' -AsPlainText -Force

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl = 'https://contoso-admin.sharepoint.com'
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'AccessTokens'
                $Script:MSCloudLoginConnectionProfile.PnP.AccessTokens = @( $secPwd )
                $Script:MSCloudLoginConnectionProfile.PnP.TenantId = 'tenant-id'
                $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment = 'Production'
                $Script:MSCloudLoginConnectionProfile.PnP.EnvironmentName = 'AzureCloud'
                $Script:MSCloudLoginConnectionProfile.PnP.Connected = $false

                Connect-MSCloudLoginPnP

                Should -Invoke Connect-PnPOnline -ParameterFilter {
                    $Url -eq 'https://contoso-admin.sharepoint.com' -and
                    $AccessToken -eq $secPwd -and
                    $AzureEnvironment -eq 'Production'
                }
            }
        }
    }
}

Describe 'Disconnect-MSCloudLoginPnP' {
    Context 'When PnP is connected' {
        It 'Should call Disconnect-PnPOnline' {
            InModuleScope 'MSCloudLoginAssistant' {
                Import-Module ./Tests/Unit/Stubs/Stubs.psm1 -Force
                Mock -CommandName Disconnect-PnPOnline -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.Connected = $true

                Disconnect-MSCloudLoginPnP

                Should -Invoke Disconnect-PnPOnline
            }
        }
    }

    Context 'When PnP is not connected' {
        It 'Should not throw and log message' {
            InModuleScope 'MSCloudLoginAssistant' {
                Import-Module ./Tests/Unit/Stubs/Stubs.psm1 -Force
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.Connected = $false

                { Disconnect-MSCloudLoginPnP } | Should -Not -Throw
            }
        }
    }
}

AfterAll {
    Remove-Module MSCloudLoginAssistant
}
