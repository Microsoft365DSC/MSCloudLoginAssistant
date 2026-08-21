#Requires -Modules Pester

Describe 'Connect-MSCloudLoginMicrosoftGraph' {
    BeforeAll {
        # Plain function stubs keep command resolution (and therefore Pester mock
        # creation) off the real SDK modules, which would otherwise be discovered
        # and imported on first use at a multi second cost.
        Import-Module (Join-Path $PSScriptRoot '..\Stubs\Stubs.psm1') -Force -Global -WarningAction SilentlyContinue
        Import-Module ./Modules/MSCloudLoginAssistant/MSCloudLoginAssistant.psd1 -Force

        # Compile and instantiate the workload classes once here so that the cost
        # does not show up inside the first test of this file.
        InModuleScope 'MSCloudLoginAssistant' {
            $null = New-Object MSCloudLoginConnectionProfile
        }
    }

    Context 'When connecting with Credentials' {
        It 'Should call Connect-MSCloudLoginMSGraphWithUser' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MSCloudLoginMSGraphWithUser -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AuthenticationType = 'Credentials'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected = $false

                Connect-MSCloudLoginMicrosoftGraph

                Should -Invoke Connect-MSCloudLoginMSGraphWithUser
            }
        }
    }

    Context 'When connecting with CredentialsWithTenantId' {
        It 'Should call Connect-MSCloudLoginMSGraphWithUser with TenantId' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MSCloudLoginMSGraphWithUser -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AuthenticationType = 'CredentialsWithTenantId'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.TenantId = 'tenant-id'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected = $false

                Connect-MSCloudLoginMicrosoftGraph

                Should -Invoke Connect-MSCloudLoginMSGraphWithUser -ParameterFilter {
                    $TenantId -eq 'tenant-id'
                }
            }
        }
    }

    Context 'When connecting with Identity' {
        It 'Should call Connect-MgGraph with AccessToken' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MgGraph -MockWith { }
                Mock -CommandName Get-MgContext -MockWith { return @{ TenantId = 'context-tenant-id' } }
                Mock -CommandName Get-AuthToken -MockWith { return 'access-token' }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AuthenticationType = 'Identity'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ResourceUrl = 'https://graph.microsoft.com'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.GraphEnvironment = 'Global'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected = $false

                Connect-MSCloudLoginMicrosoftGraph

                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected | Should -BeTrue
                # The tenant id must be adopted from the live Graph context.
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.TenantId | Should -Be 'context-tenant-id'

                Should -Invoke Connect-MgGraph -ParameterFilter {
                    $AccessToken -ne $null -and
                    $Environment -eq 'Global'
                }
                Should -Invoke Get-AuthToken -Exactly 1 -ParameterFilter {
                    $Resource -eq 'https://graph.microsoft.com' -and $Identity.IsPresent
                }
            }
        }
    }

    Context 'When connecting with ServicePrincipalWithThumbprint' {
        It 'Should call Connect-MgGraph with Certificate' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MgGraph -MockWith { }
                Mock -CommandName Get-MSCloudLoginCertificate -MockWith { return New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AuthenticationType = 'ServicePrincipalWithThumbprint'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.TenantId = 'tenant-id'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.CertificateThumbprint = 'thumbprint'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.GraphEnvironment = 'Global'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected = $false

                Connect-MSCloudLoginMicrosoftGraph

                Should -Invoke Connect-MgGraph -ParameterFilter {
                    $ClientId -eq 'app-id' -and
                    $TenantId -eq 'tenant-id' -and
                    $Environment -eq 'Global'
                }
            }
        }
    }

    Context 'When connecting with ServicePrincipalWithSecret' {
        It 'Should call Connect-MgGraph with ClientSecret' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MgGraph -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AuthenticationType = 'ServicePrincipalWithSecret'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.TenantId = 'tenant-id'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ApplicationSecret = 'secret'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.GraphEnvironment = 'Global'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected = $false

                Connect-MSCloudLoginMicrosoftGraph

                Should -Invoke Connect-MgGraph -ParameterFilter {
                    $ClientSecretCredential -ne $null -and
                    $Environment -eq 'Global'
                }
            }
        }
    }

    Context 'When connecting with AccessTokens' {
        It 'Should call Connect-MgGraph with AccessToken' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MgGraph -MockWith { }
                Mock -CommandName Get-MSCloudLoginAccessTokenValue -MockWith { return 'token-value' }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AuthenticationType = 'AccessTokens'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AccessTokens = @(@{ access_token = 'token123' })
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.GraphEnvironment = 'Global'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected = $false

                Connect-MSCloudLoginMicrosoftGraph

                Should -Invoke Connect-MgGraph -ParameterFilter {
                    $AccessToken -ne $null -and
                    $Environment -eq 'Global'
                }
            }
        }
    }

    Context 'When an existing Graph context is still valid' {
        It 'Should reuse the connection instead of connecting again' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MgGraph -MockWith { }
                Mock -CommandName Connect-MSCloudLoginMSGraphWithUser -MockWith { }
                Mock -CommandName Get-MgContext -MockWith { return @{ Account = 'admin@contoso.com' } }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AuthenticationType = 'Credentials'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected = $true
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ConnectedDateTime = [System.DateTime]::Now.ToString()

                Connect-MSCloudLoginMicrosoftGraph

                Should -Invoke Connect-MgGraph -Exactly 0
                Should -Invoke Connect-MSCloudLoginMSGraphWithUser -Exactly 0
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected | Should -BeTrue
            }
        }

        It 'Should reconnect when the SDK context is gone' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MgGraph -MockWith { }
                Mock -CommandName Connect-MSCloudLoginMSGraphWithUser -MockWith { }
                Mock -CommandName Get-MgContext -MockWith { return $null }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AuthenticationType = 'Credentials'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected = $true
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ConnectedDateTime = [System.DateTime]::Now.ToString()

                Connect-MSCloudLoginMicrosoftGraph

                Should -Invoke Connect-MSCloudLoginMSGraphWithUser -Exactly 1
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected | Should -BeFalse
            }
        }
    }

    Context 'When registering a custom environment' {
        It 'Should register the custom environment when it does not exist yet' {
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{ ModuleRoot = (Resolve-Path "$PSScriptRoot\..\..\..\Modules\MSCloudLoginAssistant").Path } {
                param ($ModuleRoot)

                Mock -CommandName Connect-MSCloudLoginMSGraphWithUser -MockWith { }
                Mock -CommandName Get-MgEnvironment -MockWith { return @() }
                Mock -CommandName Add-MgEnvironment -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }

                $Script:CustomEnvConfig.CustomEnvironment = $true
                try
                {
                    $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                    $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AuthenticationType = 'Credentials'

                    Connect-MSCloudLoginMicrosoftGraph

                    Should -Invoke Add-MgEnvironment -Exactly 1 -ParameterFilter {
                        $Name -eq 'Custom' -and
                        $GraphEndpoint -eq 'https://graph.microsoft.com/' -and
                        $AzureADEndPoint -eq 'https://login.microsoftonline.com'
                    }
                }
                finally
                {
                    $Script:CustomEnvConfig = Import-PowerShellDataFile -Path (Join-Path $ModuleRoot 'CustomEnvironment.psd1')
                    $Script:LoadedCustomEnvFileName = 'CustomEnvironment.psd1'
                }
            }
        }
    }

    Context 'When connecting with ServicePrincipalWithThumbprint in a custom environment' {
        It 'Should acquire the token locally and pass it to Connect-MgGraph' {
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{ ModuleRoot = (Resolve-Path "$PSScriptRoot\..\..\..\Modules\MSCloudLoginAssistant").Path } {
                param ($ModuleRoot)

                Mock -CommandName Connect-MgGraph -MockWith { }
                Mock -CommandName Get-MgEnvironment -MockWith { return @([PSCustomObject]@{ Name = 'Custom' }) }
                Mock -CommandName Add-MgEnvironment -MockWith { }
                Mock -CommandName Get-MSCloudLoginAccessToken -MockWith { return 'locally-issued-token' }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }

                $Script:CustomEnvConfig.CustomEnvironment = $true
                try
                {
                    $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                    $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AuthenticationType = 'ServicePrincipalWithThumbprint'
                    $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ApplicationId = 'app-id'
                    $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.TenantId = 'contoso.local'
                    $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.CertificateThumbprint = 'thumbprint'
                    $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.GraphEnvironment = 'Custom'
                    $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Scope = 'https://graph.contoso.local/.default'
                    $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.TokenUrl = 'https://login.contoso.local/oauth2/v2.0/token'
                    $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AuthorizationUrl = 'https://login.contoso.local'
                    $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected = $false

                    Connect-MSCloudLoginMicrosoftGraph

                    $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected | Should -BeTrue
                    Should -Invoke Get-MSCloudLoginAccessToken -Exactly 1 -ParameterFilter {
                        $ApplicationId -eq 'app-id' -and $TenantId -eq 'contoso.local'
                    }
                    Should -Invoke Connect-MgGraph -Exactly 1 -ParameterFilter {
                        $AccessToken -is [System.Security.SecureString] -and $Environment -eq 'Custom'
                    }
                    Should -Invoke Add-MgEnvironment -Exactly 0
                    Should -Invoke Add-MSCloudLoginAssistantEvent -ParameterFilter {
                        $Message -like '*Successfully connected to the Microsoft Graph API using Certificate Thumbprint*'
                    }
                }
                finally
                {
                    $Script:CustomEnvConfig = Import-PowerShellDataFile -Path (Join-Path $ModuleRoot 'CustomEnvironment.psd1')
                    $Script:LoadedCustomEnvFileName = 'CustomEnvironment.psd1'
                }
            }
        }
    }

    Context 'When connecting with ServicePrincipalWithPath' {
        It 'Should load the certificate from disk and pass it to Connect-MgGraph' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MgGraph -MockWith { }
                Mock -CommandName Get-MSCloudLoginCertificate -MockWith { return New-Object System.Security.Cryptography.X509Certificates.X509Certificate2 }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AuthenticationType = 'ServicePrincipalWithPath'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.TenantId = 'tenant-id'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.CertificatePath = 'C:\certs\contoso.pfx'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.CertificatePassword = ConvertTo-SecureString 'cert-password' -AsPlainText -Force
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.GraphEnvironment = 'Global'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected = $false

                Connect-MSCloudLoginMicrosoftGraph

                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected | Should -BeTrue
                Should -Invoke Get-MSCloudLoginCertificate -ParameterFilter {
                    $CertificatePath -eq 'C:\certs\contoso.pfx'
                }
                Should -Invoke Connect-MgGraph -ParameterFilter {
                    $TenantId -eq 'tenant-id' -and
                    $ClientId -eq 'app-id' -and
                    $Certificate -is [System.Security.Cryptography.X509Certificates.X509Certificate2]
                }
            }
        }
    }

    Context 'When the authentication type is not supported' {
        It 'Should throw an error naming the unsupported type' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AuthenticationType = 'Interactive'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected = $false

                { Connect-MSCloudLoginMicrosoftGraph } |
                    Should -Throw "*Authentication type 'Interactive' is not supported for workload 'MicrosoftGraph'*"
            }
        }
    }

    Context 'When the connection fails' {
        It 'Should surface the underlying error and leave the workload disconnected' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-MgGraph -MockWith { throw 'the graph tenant is unreachable' }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AuthenticationType = 'ServicePrincipalWithSecret'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.TenantId = 'tenant-id'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ApplicationSecret = 'secret'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.GraphEnvironment = 'Global'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected = $false

                { Connect-MSCloudLoginMicrosoftGraph } | Should -Throw '*unreachable*'

                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected | Should -BeFalse
                Should -Invoke Add-MSCloudLoginAssistantEvent -ParameterFilter {
                    $EntryType -eq 'Error' -and $Message -like '*Failed to connect to Microsoft Graph*'
                }
            }
        }
    }
}

Describe 'Connect-MSCloudLoginMSGraphWithUser' {

    It 'Should disconnect the stale account, register the custom environment and connect with the acquired token' {
        InModuleScope 'MSCloudLoginAssistant' -Parameters @{ ModuleRoot = (Resolve-Path "$PSScriptRoot\..\..\..\Modules\MSCloudLoginAssistant").Path } {
            param ($ModuleRoot)

            Mock -CommandName Get-MgContext -MockWith { return $null }
            Mock -CommandName Get-MgEnvironment -MockWith { return @([PSCustomObject]@{ Name = 'OtherEnvironment' }) }
            Mock -CommandName Add-MgEnvironment -MockWith { }
            Mock -CommandName Disconnect-MgGraph -MockWith { }
            Mock -CommandName Get-AuthToken -MockWith { return @{ access_token = 'user-token' } }
            Mock -CommandName Connect-MgGraph -MockWith { }
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

            $Script:CustomEnvConfig.CustomEnvironment = $true
            try
            {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $graphProfile = $Script:MSCloudLoginConnectionProfile.MicrosoftGraph
                $graphProfile.Credentials = New-Object PSCredential ('admin@contoso.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))
                $graphProfile.AuthorizationUrl = 'https://login.microsoftonline.com'
                $graphProfile.Scope = 'https://graph.microsoft.com/.default'

                Connect-MSCloudLoginMSGraphWithUser

                $graphProfile.ApplicationId | Should -Be '14d82eec-204b-4c2f-b7e8-296a70dab67e'
                $graphProfile.AccessTokens | Should -Be 'user-token'
                $graphProfile.Connected | Should -BeTrue

                Should -Invoke Add-MgEnvironment -Exactly 1 -ParameterFilter {
                    $Name -eq 'Custom' -and $GraphEndpoint -eq 'https://graph.microsoft.com/'
                }
                Should -Invoke Disconnect-MgGraph -Exactly 1
                Should -Invoke Get-AuthToken -Exactly 1 -ParameterFilter {
                    $null -ne $Credentials -and $ClientId -eq '14d82eec-204b-4c2f-b7e8-296a70dab67e'
                }
                Should -Invoke Connect-MgGraph -Exactly 1 -ParameterFilter {
                    $AccessToken -is [System.Security.SecureString] -and $Environment -eq 'Global'
                }
            }
            finally
            {
                $Script:CustomEnvConfig = Import-PowerShellDataFile -Path (Join-Path $ModuleRoot 'CustomEnvironment.psd1')
                $Script:LoadedCustomEnvFileName = 'CustomEnvironment.psd1'
            }
        }
    }

    It 'Should tolerate a failing disconnect of the previous session' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Get-MgContext -MockWith { return @{ Account = 'other@contoso.com' } }
            Mock -CommandName Disconnect-MgGraph -MockWith { throw 'no session to remove' }
            Mock -CommandName Get-AuthToken -MockWith { return @{ access_token = 'user-token' } }
            Mock -CommandName Connect-MgGraph -MockWith { }
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

            $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
            $graphProfile = $Script:MSCloudLoginConnectionProfile.MicrosoftGraph
            $graphProfile.Credentials =
                New-Object PSCredential ('admin@contoso.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))
            $graphProfile.AuthorizationUrl = 'https://login.microsoftonline.com'
            $graphProfile.Scope = 'https://graph.microsoft.com/.default'

            { Connect-MSCloudLoginMSGraphWithUser } | Should -Not -Throw

            $graphProfile.Connected | Should -BeTrue
            Should -Invoke Add-MSCloudLoginAssistantEvent -ParameterFilter {
                $Message -like '*Disconnecting from Microsoft Graph failed*'
            }
        }
    }

    It 'Should fall back to the device code flow when MFA is required' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Get-MgContext -MockWith { return $null }
            Mock -CommandName Disconnect-MgGraph -MockWith { }
            Mock -CommandName Get-AuthToken -MockWith { throw 'AADSTS50076: multi-factor authentication is required' }
            Mock -CommandName Connect-MgGraph -MockWith { }
            Mock -CommandName Connect-MSCloudLoginMSGraphWithUserMFA -MockWith { }
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

            $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
            $graphProfile = $Script:MSCloudLoginConnectionProfile.MicrosoftGraph
            $graphProfile.Credentials =
                New-Object PSCredential ('admin@contoso.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))
            $graphProfile.AuthorizationUrl = 'https://login.microsoftonline.com'
            $graphProfile.Scope = 'https://graph.microsoft.com/.default'

            { Connect-MSCloudLoginMSGraphWithUser } | Should -Not -Throw

            Should -Invoke Connect-MSCloudLoginMSGraphWithUserMFA -Exactly 1
        }
    }

    It 'Should explain the missing application registration on a bad request in a non interactive shell' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Get-MgContext -MockWith { return $null }
            Mock -CommandName Disconnect-MgGraph -MockWith { }
            Mock -CommandName Get-AuthToken -MockWith {
                throw [System.Net.WebException]::new('System.Net.WebException: The remote server returned an error: (400) Bad Request.')
            }
            Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $true }
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

            $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
            $graphProfile = $Script:MSCloudLoginConnectionProfile.MicrosoftGraph
            $graphProfile.Credentials =
                New-Object PSCredential ('admin@contoso.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))
            $graphProfile.AuthorizationUrl = 'https://login.microsoftonline.com'
            $graphProfile.Scope = 'https://graph.microsoft.com/.default'

            { Connect-MSCloudLoginMSGraphWithUser } |
                Should -Throw "*Unable to retrieve AccessToken. Have you registered the 'Microsoft Graph PowerShell' application already*"

            $graphProfile.Connected | Should -BeFalse
        }
    }

    It 'Should reconnect without an environment when the environment specific connect fails' {
        InModuleScope 'MSCloudLoginAssistant' {
            $Script:microsoftGraphConnectCalls = 0
            Mock -CommandName Get-MgContext -MockWith { return $null }
            Mock -CommandName Disconnect-MgGraph -MockWith { }
            Mock -CommandName Get-AuthToken -MockWith { return @{ access_token = 'user-token' } }
            Mock -CommandName Test-MSCloudLoginMFARequiredError -MockWith { return $false }
            Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $false }
            Mock -CommandName Connect-MgGraph -MockWith {
                $Script:microsoftGraphConnectCalls++
                if ($Script:microsoftGraphConnectCalls -eq 1)
                {
                    # The first attempt with an explicit environment fails.
                    throw 'the environment specific endpoint rejected the connection'
                }
                # The fallback attempt without an environment succeeds.
            }
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

            $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
            $graphProfile = $Script:MSCloudLoginConnectionProfile.MicrosoftGraph
            $graphProfile.Credentials =
                New-Object PSCredential ('admin@contoso.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))
            $graphProfile.AuthorizationUrl = 'https://login.microsoftonline.com'
            $graphProfile.Scope = 'https://graph.microsoft.com/.default'

            { Connect-MSCloudLoginMSGraphWithUser } | Should -Not -Throw

            $graphProfile.Connected | Should -BeTrue
            $graphProfile.AccessTokens | Should -Be 'user-token'
            Should -Invoke Add-MSCloudLoginAssistantEvent -ParameterFilter {
                $Message -like '*Attempting to connect without specifying the Environment*'
            }
            $Script:microsoftGraphConnectCalls | Should -Be 2
            # The successful retry must be the call that omits the environment.
            Should -Invoke Connect-MgGraph -Exactly 1 -ParameterFilter {
                $AccessToken -is [System.Security.SecureString] -and $null -eq $Environment
            }
        }
    }

    It 'Should refuse the interactive fallback in a non interactive session' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Get-MgContext -MockWith { return $null }
            Mock -CommandName Disconnect-MgGraph -MockWith { }
            Mock -CommandName Get-AuthToken -MockWith { return @{ access_token = 'user-token' } }
            Mock -CommandName Test-MSCloudLoginMFARequiredError -MockWith { return $false }
            Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $true }
            Mock -CommandName Connect-MgGraph -MockWith { throw 'token connect failed' }
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

            $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
            $graphProfile = $Script:MSCloudLoginConnectionProfile.MicrosoftGraph
            $graphProfile.Credentials =
                New-Object PSCredential ('admin@contoso.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))
            $graphProfile.AuthorizationUrl = 'https://login.microsoftonline.com'
            $graphProfile.Scope = 'https://graph.microsoft.com/.default'

            # The bare throw in the non-interactive guard re-throws the original
            # connection failure; the refusal itself is recorded as an error event.
            { Connect-MSCloudLoginMSGraphWithUser } | Should -Throw '*token connect failed*'

            $graphProfile.Connected | Should -BeFalse
            Should -Invoke Add-MSCloudLoginAssistantEvent -ParameterFilter {
                $Message -like '*Error connecting*' -and $EntryType -eq 'Error'
            }
            Should -Invoke Add-MSCloudLoginAssistantEvent -ParameterFilter {
                $Message -like '*Unable to connect to Microsoft Graph and interactive fallback is not possible*' -and
                $EntryType -eq 'Error'
            }
        }
    }

    It 'Should reconnect interactively when both token based attempts fail' {
        InModuleScope 'MSCloudLoginAssistant' {
            $Script:microsoftGraphConnectCalls = 0
            Mock -CommandName Get-MgContext -MockWith { return $null }
            Mock -CommandName Disconnect-MgGraph -MockWith { }
            Mock -CommandName Get-AuthToken -MockWith { return @{ access_token = 'user-token' } }
            Mock -CommandName Test-MSCloudLoginMFARequiredError -MockWith { return $false }
            Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $false }
            Mock -CommandName Connect-MgGraph -MockWith {
                $Script:microsoftGraphConnectCalls++
                if ($Script:microsoftGraphConnectCalls -le 2)
                {
                    # Both token based attempts fail; only the interactive sign-in works.
                    throw 'the token endpoint could not be reached'
                }
            }
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

            $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
            $graphProfile = $Script:MSCloudLoginConnectionProfile.MicrosoftGraph
            $graphProfile.Credentials =
                New-Object PSCredential ('admin@contoso.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))
            $graphProfile.AuthorizationUrl = 'https://login.microsoftonline.com'
            $graphProfile.Scope = 'https://graph.microsoft.com/.default'

            { Connect-MSCloudLoginMSGraphWithUser } | Should -Not -Throw

            $graphProfile.Connected | Should -BeTrue
            $graphProfile.MultiFactorAuthentication | Should -BeFalse
            Should -Invoke Add-MSCloudLoginAssistantEvent -ParameterFilter {
                $Message -like '*Connecting to Microsoft Graph interactively*'
            }
            $Script:microsoftGraphConnectCalls | Should -Be 3
        }
    }

    It 'Should recreate the graph context file and retry when saving the context fails' {
        InModuleScope 'MSCloudLoginAssistant' {
            $contextDirectory = Join-Path $env:TEMP "opencode\mscloudlogin-graphctx-$PID\.graph"
            $contextPath = Join-Path $contextDirectory 'GraphContext.json'
            $Script:graphContextRetryCount = 0

            Mock -CommandName Get-MgContext -MockWith { return $null }
            Mock -CommandName Disconnect-MgGraph -MockWith { }
            Mock -CommandName Get-AuthToken -MockWith { return @{ access_token = 'user-token' } }
            Mock -CommandName Test-MSCloudLoginMFARequiredError -MockWith { return $false }
            Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $false }
            Mock -CommandName Connect-MgGraph -MockWith {
                if ($null -ne $Scopes)
                {
                    $Script:graphContextRetryCount++
                    if (-not (Test-Path -LiteralPath $contextPath))
                    {
                        # The first interactive attempt fails because the context file cannot be written.
                        throw "Failed to save graph context to file at '$contextPath'."
                    }
                    return
                }
                throw 'the token endpoint could not be reached'
            }
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

            New-Item -Path $contextDirectory -ItemType Directory -Force | Out-Null
            try
            {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $graphProfile = $Script:MSCloudLoginConnectionProfile.MicrosoftGraph
                $graphProfile.Credentials =
                    New-Object PSCredential ('admin@contoso.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))
                $graphProfile.AuthorizationUrl = 'https://login.microsoftonline.com'
                $graphProfile.Scope = 'https://graph.microsoft.com/.default'

                { Connect-MSCloudLoginMSGraphWithUser } | Should -Not -Throw

                Test-Path -LiteralPath $contextPath | Should -BeTrue
                $Script:graphContextRetryCount | Should -Be 2
                $graphProfile.Connected | Should -BeTrue
            }
            finally
            {
                Remove-Item -Path (Split-Path $contextDirectory -Parent) -Recurse -Force -ErrorAction SilentlyContinue
                Remove-Variable -Name graphContextRetryCount -Scope Script -ErrorAction SilentlyContinue
            }
        }
    }

    It 'Should explain the app permissions when the device code terminal timed out' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Get-MgContext -MockWith { return $null }
            Mock -CommandName Disconnect-MgGraph -MockWith { }
            Mock -CommandName Get-AuthToken -MockWith { return @{ access_token = 'user-token' } }
            Mock -CommandName Test-MSCloudLoginMFARequiredError -MockWith { return $false }
            Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $false }
            Mock -CommandName Connect-MgGraph -MockWith {
                if ($null -ne $Scopes)
                {
                    throw 'Device code terminal timed-out after 120 seconds. Please try again.'
                }
                throw 'the token endpoint could not be reached'
            }
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

            $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
            $graphProfile = $Script:MSCloudLoginConnectionProfile.MicrosoftGraph
            $graphProfile.Credentials =
                New-Object PSCredential ('admin@contoso.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))
            $graphProfile.AuthorizationUrl = 'https://login.microsoftonline.com'
            $graphProfile.Scope = 'https://graph.microsoft.com/.default'

            { Connect-MSCloudLoginMSGraphWithUser } |
                Should -Throw '*Please make sure the app permissions are setup correctly*'

            $graphProfile.Connected | Should -BeFalse
        }
    }

    It 'Should surface any other interactive failure and leave the workload disconnected' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Get-MgContext -MockWith { return $null }
            Mock -CommandName Disconnect-MgGraph -MockWith { }
            Mock -CommandName Get-AuthToken -MockWith { return @{ access_token = 'user-token' } }
            Mock -CommandName Test-MSCloudLoginMFARequiredError -MockWith { return $false }
            Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $false }
            Mock -CommandName Connect-MgGraph -MockWith {
                if ($null -ne $Scopes)
                {
                    throw 'sign-in was cancelled by the user'
                }
                throw 'the token endpoint could not be reached'
            }
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

            $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
            $graphProfile = $Script:MSCloudLoginConnectionProfile.MicrosoftGraph
            $graphProfile.Credentials =
                New-Object PSCredential ('admin@contoso.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))
            $graphProfile.AuthorizationUrl = 'https://login.microsoftonline.com'
            $graphProfile.Scope = 'https://graph.microsoft.com/.default'

            { Connect-MSCloudLoginMSGraphWithUser } | Should -Throw '*cancelled by the user*'

            $graphProfile.Connected | Should -BeFalse
            Should -Invoke Add-MSCloudLoginAssistantEvent -ParameterFilter {
                $Message -like '*Failed to connect to Microsoft Graph interactively*cancelled by the user*' -and
                $EntryType -eq 'Error'
            }
        }
    }
}

Describe 'Connect-MSCloudLoginMSGraphWithUserMFA' {

    It 'Should derive the tenant from the credential when no tenant id is set' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Get-AuthToken -MockWith { return @{ access_token = 'mfa-token' } }
            Mock -CommandName Connect-MgGraph -MockWith { }
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

            $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
            $graphProfile = $Script:MSCloudLoginConnectionProfile.MicrosoftGraph
            $graphProfile.Credentials =
                New-Object PSCredential ('admin@contoso.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))
            $graphProfile.TenantId = ''
            $graphProfile.Scope = 'https://graph.microsoft.com/.default'
            $graphProfile.AuthorizationUrl = 'https://login.microsoftonline.com'

            Connect-MSCloudLoginMSGraphWithUserMFA

            $graphProfile.Connected | Should -BeTrue
            $graphProfile.MultiFactorAuthentication | Should -BeTrue
            $graphProfile.AccessTokens | Should -Be 'mfa-token'
            Should -Invoke Get-AuthToken -Exactly 1 -ParameterFilter {
                $TenantId -eq 'contoso.com' -and $DeviceCode.IsPresent
            }
            Should -Invoke Connect-MgGraph -Exactly 1 -ParameterFilter {
                $AccessToken -is [System.Security.SecureString] -and $Environment -eq 'Global'
            }
        }
    }

    It 'Should use the tenant id stored on the profile when it is set' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Get-AuthToken -MockWith { return @{ access_token = 'mfa-token' } }
            Mock -CommandName Connect-MgGraph -MockWith { }
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

            $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
            $graphProfile = $Script:MSCloudLoginConnectionProfile.MicrosoftGraph
            $graphProfile.Credentials =
                New-Object PSCredential ('admin@contoso.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))
            $graphProfile.TenantId = 'contoso.onmicrosoft.com'
            $graphProfile.Scope = 'https://graph.microsoft.com/.default'
            $graphProfile.AuthorizationUrl = 'https://login.microsoftonline.com'

            Connect-MSCloudLoginMSGraphWithUserMFA

            Should -Invoke Get-AuthToken -Exactly 1 -ParameterFilter {
                $TenantId -eq 'contoso.onmicrosoft.com'
            }
        }
    }
}

Describe 'Disconnect-MSCloudLoginMicrosoftGraph' {
    Context 'When MicrosoftGraph is connected' {
        It 'Should call Disconnect-MgGraph' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Disconnect-MgGraph -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected = $true

                Disconnect-MSCloudLoginMicrosoftGraph

                Should -Invoke Disconnect-MgGraph
            }
        }
    }

    Context 'When MicrosoftGraph is not connected' {
        It 'Should not throw and log message' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected = $false

                { Disconnect-MSCloudLoginMicrosoftGraph } | Should -Not -Throw
            }
        }
    }
}

AfterAll {
    Remove-Module MSCloudLoginAssistant
}
