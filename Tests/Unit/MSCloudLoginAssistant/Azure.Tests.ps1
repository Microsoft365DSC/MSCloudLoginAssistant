#Requires -Modules Pester

Describe 'Connect-MSCloudLoginAzure' {
    BeforeAll {
        Import-Module ./Modules/MSCloudLoginAssistant/MSCloudLoginAssistant.psd1 -Force
    }

    Context 'When the sign-in fails' {
        BeforeEach {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Azure.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.Azure.TenantId = 'tenant-id'
                $Script:MSCloudLoginConnectionProfile.Azure.EnvironmentName = 'AzureCloud'
                $Script:MSCloudLoginConnectionProfile.Azure.Connected = $false
            }
        }

        It 'Should rethrow a <Kind> <AuthenticationType> sign-in error after <ExpectedSignIns> sign-in(s) and stay disconnected' -TestCases @(
            @{
                Kind               = 'non-terminating'
                AuthenticationType = 'ServicePrincipalWithThumbprint'
                Settings           = @{ CertificateThumbprint = 'thumbprint' }
                SignInError        = 'The provided account app-id does not have access to subscription ID'
                ExpectedSignIns    = 1
            }
            @{
                Kind               = 'terminating'
                AuthenticationType = 'ServicePrincipalWithSecret'
                Settings           = @{ ApplicationSecret = 'secret' }
                SignInError        = 'AADSTS7000215: Invalid client secret'
                ExpectedSignIns    = 1
            }
            @{
                Kind               = 'non-terminating'
                AuthenticationType = 'Identity'
                Settings           = @{}
                SignInError        = 'No managed identity endpoint found'
                ExpectedSignIns    = 1
            }
            @{
                Kind               = 'non-MFA'
                AuthenticationType = 'Credentials'
                Settings           = @{ Credentials = New-Object PSCredential ('user@contoso.com', (ConvertTo-SecureString 'password' -AsPlainText -Force)) }
                SignInError        = 'AADSTS50126: Invalid username or password'
                ExpectedSignIns    = 1
            }
            @{
                Kind               = 'unsupported'
                AuthenticationType = 'Interactive'
                Settings           = @{}
                SignInError        = 'Specified authentication method is not supported'
                ExpectedSignIns    = 0
            }
        ) {
            InModuleScope 'MSCloudLoginAssistant' -Parameters $_ {
                param ($Kind, $AuthenticationType, $Settings, $SignInError, $ExpectedSignIns)

                Mock -CommandName Connect-AzAccount -MockWith {
                    if ($Kind -eq 'non-terminating')
                    {
                        Write-Error $SignInError
                        return
                    }
                    throw $SignInError
                }

                $Script:MSCloudLoginConnectionProfile.Azure.AuthenticationType = $AuthenticationType
                foreach ($setting in $Settings.GetEnumerator())
                {
                    $Script:MSCloudLoginConnectionProfile.Azure.($setting.Key) = $setting.Value
                }

                { Connect-MSCloudLoginAzure } | Should -Throw "*$SignInError*"
                $Script:MSCloudLoginConnectionProfile.Azure.Connected | Should -BeFalse
                Should -Invoke Connect-AzAccount -Exactly $ExpectedSignIns
                Should -Invoke Add-MSCloudLoginAssistantEvent -ParameterFilter {
                    $EntryType -eq 'Error' -and $Message -like 'Failed to connect to Azure:*'
                }
            }
        }
    }
}

AfterAll {
    Remove-Module MSCloudLoginAssistant
}
