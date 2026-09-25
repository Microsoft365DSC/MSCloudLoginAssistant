#Requires -Modules Pester

BeforeAll {
    $graphModuleName = 'Microsoft.Graph.Beta.Identity.DirectoryManagement'
    if (-not (Get-Module -Name $graphModuleName -ListAvailable))
    {
        $script:tempModuleBase = Join-Path $env:TEMP 'MSCloudLoginTestModules'
        $tempModuleDir = Join-Path $script:tempModuleBase $graphModuleName
        if (-not (Test-Path $tempModuleDir))
        {
            New-Item -Path $tempModuleDir -ItemType Directory -Force | Out-Null
        }
        $manifestPath = Join-Path $tempModuleDir "$graphModuleName.psd1"
        if (-not (Test-Path $manifestPath))
        {
            New-ModuleManifest -Path $manifestPath -ModuleVersion '1.0.0' -Description 'Test stub'
        }
        $env:PSModulePath = $script:tempModuleBase + [IO.Path]::PathSeparator + $env:PSModulePath
    }

    Import-Module (Join-Path $PSScriptRoot '..\Stubs\Stubs.psm1') -Force -Global -WarningAction SilentlyContinue
    $script:moduleRoot = Resolve-Path (Join-Path $PSScriptRoot '..\..\..\Modules\MSCloudLoginAssistant')
    Import-Module (Join-Path $script:moduleRoot 'MSCloudLoginAssistant.psd1') -Force
}

AfterAll {
    if ($script:tempModuleBase -and (Test-Path $script:tempModuleBase))
    {
        Remove-Item -Path $script:tempModuleBase -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Describe 'Connect-MSCloudLoginMicrosoftGraph failure handling' {

    BeforeEach {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
            $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
            $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AuthorizationUrl = 'https://login.microsoftonline.com'
            $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Scope = 'https://graph.microsoft.com/.default'
        }
    }

    It 'Should reject an authentication type it does not support' {
        InModuleScope 'MSCloudLoginAssistant' {
            $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AuthenticationType = 'Interactive'

            { Connect-MSCloudLoginMicrosoftGraph } |
                Should -Throw "*'Interactive' is not supported for workload 'MicrosoftGraph'*"
            $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Connected | Should -BeFalse
        }
    }

    It 'Should rethrow a failing certificate based sign-in' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Get-MgContext -MockWith { return $null }
            Mock -CommandName Get-MSCloudLoginCertificate -MockWith {
                return New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
            }
            Mock -CommandName Connect-MgGraph -MockWith { throw 'AADSTS700027: certificate is not valid' }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.MicrosoftGraph
            $workloadProfile.AuthenticationType = 'ServicePrincipalWithThumbprint'
            $workloadProfile.ApplicationId = 'app-id'
            $workloadProfile.TenantId = 'contoso.onmicrosoft.com'
            $workloadProfile.CertificateThumbprint = 'thumbprint'

            { Connect-MSCloudLoginMicrosoftGraph } | Should -Throw '*AADSTS700027*'
            $workloadProfile.Connected | Should -BeFalse
        }
    }

    It 'Should retry without an environment when the environment specific sign-in fails' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Get-MgContext -MockWith { return $null }
            Mock -CommandName Disconnect-MgGraph -MockWith { }
            Mock -CommandName Get-AuthToken -MockWith { return @{ access_token = 'delegated-token' } }
            Mock -CommandName Connect-MgGraph -MockWith {
                if (-not [System.String]::IsNullOrEmpty($Environment))
                {
                    throw 'the environment is unknown'
                }
            }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.MicrosoftGraph
            $workloadProfile.AuthenticationType = 'Credentials'
            $workloadProfile.Credentials = New-Object PSCredential ('admin@contoso.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))

            Connect-MSCloudLoginMicrosoftGraph

            $workloadProfile.Connected | Should -BeTrue
            $workloadProfile.AccessTokens | Should -Be @('delegated-token')
        }
    }

    It 'Should give up in a non interactive session when every sign-in attempt fails' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Get-MgContext -MockWith { return $null }
            Mock -CommandName Disconnect-MgGraph -MockWith { }
            Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $true }
            Mock -CommandName Get-AuthToken -MockWith { return @{ access_token = 'delegated-token' } }
            Mock -CommandName Connect-MgGraph -MockWith { throw 'the tenant does not exist' }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.MicrosoftGraph
            $workloadProfile.AuthenticationType = 'Credentials'
            $workloadProfile.Credentials = New-Object PSCredential ('admin@contoso.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))

            { Connect-MSCloudLoginMicrosoftGraph } | Should -Throw '*the tenant does not exist*'
            $workloadProfile.Connected | Should -BeFalse
        }
    }

    It 'Should sign in interactively as a last resort in an interactive session' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Get-MgContext -MockWith { return $null }
            Mock -CommandName Disconnect-MgGraph -MockWith { }
            Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $false }
            Mock -CommandName Get-AuthToken -MockWith { return @{ access_token = 'delegated-token' } }
            Mock -CommandName Connect-MgGraph -MockWith {
                if ($null -eq $Scopes)
                {
                    throw 'the token was rejected'
                }
            }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.MicrosoftGraph
            $workloadProfile.AuthenticationType = 'Credentials'
            $workloadProfile.Credentials = New-Object PSCredential ('admin@contoso.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))

            Connect-MSCloudLoginMicrosoftGraph

            $workloadProfile.Connected | Should -BeTrue
            Should -Invoke Connect-MgGraph -ParameterFilter { $Scopes -contains 'Domain.Read.All' }
        }
    }

    It 'Should translate a device code timeout into an actionable error' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Get-MgContext -MockWith { return $null }
            Mock -CommandName Disconnect-MgGraph -MockWith { }
            Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $false }
            Mock -CommandName Get-AuthToken -MockWith { return @{ access_token = 'delegated-token' } }
            Mock -CommandName Connect-MgGraph -MockWith {
                if ($null -ne $Scopes)
                {
                    throw 'Device code terminal timed-out after 120 seconds. Please try again.'
                }
                throw 'the token was rejected'
            }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.MicrosoftGraph
            $workloadProfile.AuthenticationType = 'Credentials'
            $workloadProfile.Credentials = New-Object PSCredential ('admin@contoso.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))

            { Connect-MSCloudLoginMicrosoftGraph } | Should -Throw '*Update-M365DSCAllowedGraphScopes*'
            $workloadProfile.Connected | Should -BeFalse
        }
    }

    It 'Should keep going when disconnecting a foreign Graph account fails' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Get-MgContext -MockWith { return @{ Account = 'someone.else@contoso.com' } }
            Mock -CommandName Disconnect-MgGraph -MockWith { throw 'there is no active session' }
            Mock -CommandName Get-AuthToken -MockWith { return @{ access_token = 'delegated-token' } }
            Mock -CommandName Connect-MgGraph -MockWith { }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.MicrosoftGraph
            $workloadProfile.AuthenticationType = 'Credentials'
            $workloadProfile.Credentials = New-Object PSCredential ('admin@contoso.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))

            Connect-MSCloudLoginMSGraphWithUser

            $workloadProfile.Connected | Should -BeTrue
        }
    }

    It 'Should derive the tenant from the credential for the device code sign-in' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Connect-MgGraph -MockWith { }
            Mock -CommandName Get-AuthToken -MockWith { return @{ access_token = 'device-code-token' } }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.MicrosoftGraph
            $workloadProfile.Credentials = New-Object PSCredential ('admin@contoso.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))

            Connect-MSCloudLoginMSGraphWithUserMFA

            $workloadProfile.MultiFactorAuthentication | Should -BeTrue
            Should -Invoke Get-AuthToken -ParameterFilter { $TenantId -eq 'contoso.com' -and $DeviceCode.IsPresent }
        }
    }
}

Describe 'Connect-MSCloudLoginExchangeOnline failure handling' {

    BeforeEach {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
            Mock -CommandName Remove-MSCloudLoginProxyModule -MockWith { }
            Mock -CommandName Disconnect-ExchangeOnline -MockWith { }
            $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
            $Script:MSCloudLoginCurrentLoadedModule = $null
        }
    }

    It 'Should return immediately when the workload is already connected' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Get-ConnectionInformation -MockWith { return @() }
            Mock -CommandName Connect-ExchangeOnline -MockWith { }
            Mock -CommandName Restore-MSCloudLoginProxyModule -MockWith { return $true }

            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CompleteConnection()
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.LoadedAllCmdlets = $true
            $Script:MSCloudLoginCurrentLoadedModule = 'EXO'

            Connect-MSCloudLoginExchangeOnline

            Should -Invoke Connect-ExchangeOnline -Exactly 0
            Should -Invoke Restore-MSCloudLoginProxyModule -Exactly 0
        }
    }

    It 'Should restore the Exchange Online proxy module after a Security & Compliance connection' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Get-ConnectionInformation -MockWith { return @() }
            Mock -CommandName Connect-ExchangeOnline -MockWith { }
            Mock -CommandName Restore-MSCloudLoginProxyModule -MockWith { return $true }

            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CompleteConnection()
            $Script:MSCloudLoginCurrentLoadedModule = 'SC'

            Connect-MSCloudLoginExchangeOnline

            $Script:MSCloudLoginCurrentLoadedModule | Should -Be 'EXO'
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected | Should -BeTrue
            Should -Invoke Restore-MSCloudLoginProxyModule -Exactly 1 -ParameterFilter { $ProbeCommand -eq 'Get-AcceptedDomain' }
            Should -Invoke Connect-ExchangeOnline -Exactly 0
        }
    }

    It 'Should reconnect when the Exchange Online proxy module is no longer loaded and <CurrentLoadedModule> was loaded last' -TestCases @(
        @{ CurrentLoadedModule = 'SC' }
        @{ CurrentLoadedModule = 'EXO' }
    ) {
        param ($CurrentLoadedModule)
        InModuleScope 'MSCloudLoginAssistant' -Parameters @{ CurrentLoadedModule = $CurrentLoadedModule } {
            param ($CurrentLoadedModule)

            Mock -CommandName Get-ConnectionInformation -MockWith { return @() }
            Mock -CommandName Connect-ExchangeOnline -MockWith { }
            Mock -CommandName Restore-MSCloudLoginProxyModule -MockWith { return $false }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.ExchangeOnline
            $workloadProfile.AuthenticationType = 'ServicePrincipalWithThumbprint'
            $workloadProfile.ApplicationId = 'app-id'
            $workloadProfile.TenantId = 'contoso.onmicrosoft.com'
            $workloadProfile.CertificateThumbprint = 'thumbprint'
            $workloadProfile.CompleteConnection()
            $Script:MSCloudLoginCurrentLoadedModule = $CurrentLoadedModule

            Connect-MSCloudLoginExchangeOnline

            $Script:MSCloudLoginCurrentLoadedModule | Should -Be 'EXO'
            $workloadProfile.Connected | Should -BeTrue
            Should -Invoke Connect-ExchangeOnline -Exactly 1
        }
    }

    It 'Should not adopt the proxy module of a Security & Compliance session' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Connect-ExchangeOnline -MockWith { }
            Mock -CommandName Import-Module -MockWith { }
            Mock -CommandName Get-ConnectionInformation -MockWith {
                return @([PSCustomObject]@{
                        Name         = 'ExchangeOnline_2'
                        AppId        = 'app-id'
                        Organization = 'contoso.onmicrosoft.com'
                        ModuleName   = 'tmpEXO_compliance'
                        IsEopSession = $true
                    })
            }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.ExchangeOnline
            $workloadProfile.AuthenticationType = 'ServicePrincipalWithThumbprint'
            $workloadProfile.ApplicationId = 'app-id'
            $workloadProfile.TenantId = 'contoso.onmicrosoft.com'
            $workloadProfile.CertificateThumbprint = 'thumbprint'

            Connect-MSCloudLoginExchangeOnline

            Should -Invoke Import-Module -Exactly 0 -ParameterFilter { $Name -eq 'tmpEXO_compliance' }
            Should -Invoke Disconnect-ExchangeOnline -Exactly 0
            Should -Invoke Connect-ExchangeOnline -Exactly 1
        }
    }

    It 'Should reconnect when a requested cmdlet is not loaded yet' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Get-ConnectionInformation -MockWith { return @() }
            Mock -CommandName Connect-ExchangeOnline -MockWith { }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.ExchangeOnline
            $workloadProfile.AuthenticationType = 'ServicePrincipalWithThumbprint'
            $workloadProfile.ApplicationId = 'app-id'
            $workloadProfile.TenantId = 'contoso.onmicrosoft.com'
            $workloadProfile.CertificateThumbprint = 'thumbprint'
            $workloadProfile.CmdletsToLoad = @('Get-Mailbox')
            $workloadProfile.LoadedCmdlets = @('Get-AcceptedDomain')
            $Script:MSCloudLoginCurrentLoadedModule = 'EXO'

            Connect-MSCloudLoginExchangeOnline

            Should -Invoke Connect-ExchangeOnline -Exactly 1 -ParameterFilter {
                $CommandName -contains 'Get-Mailbox' -and $CommandName -contains 'Get-AcceptedDomain'
            }
        }
    }

    It 'Should adopt an existing session that belongs to the same user' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Connect-ExchangeOnline -MockWith { }
            Mock -CommandName Import-Module -MockWith { }
            Mock -CommandName Get-ConnectionInformation -MockWith {
                return @([PSCustomObject]@{
                    Name              = 'ExchangeOnline_1'
                    UserPrincipalName = 'admin@contoso.onmicrosoft.com'
                    ModuleName        = 'tmpEXO_userabc'
                })
            }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.ExchangeOnline
            $workloadProfile.AuthenticationType = 'Credentials'
            $workloadProfile.Credentials = New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))

            Connect-MSCloudLoginExchangeOnline

            $workloadProfile.Connected | Should -BeTrue
            Should -Invoke Import-Module -Exactly 1 -ParameterFilter { $Name -eq 'tmpEXO_userabc' }
            Should -Invoke Connect-ExchangeOnline -Exactly 0
        }
    }

    It 'Should rethrow and disconnect when <AuthenticationType> fails' -TestCases @(
        @{ AuthenticationType = 'ServicePrincipalWithThumbprint' }
        @{ AuthenticationType = 'Identity' }
        @{ AuthenticationType = 'AccessTokens' }
    ) {
        param ($AuthenticationType)
        InModuleScope 'MSCloudLoginAssistant' -Parameters @{ AuthenticationType = $AuthenticationType } {
            param ($AuthenticationType)

            Mock -CommandName Get-ConnectionInformation -MockWith { return @() }
            Mock -CommandName Connect-ExchangeOnline -MockWith { throw 'the service is unavailable' }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.ExchangeOnline
            $workloadProfile.AuthenticationType = $AuthenticationType
            $workloadProfile.ApplicationId = 'app-id'
            $workloadProfile.TenantId = 'contoso.onmicrosoft.com'
            $workloadProfile.CertificateThumbprint = 'thumbprint'
            $workloadProfile.AccessTokens = @('token')

            { Connect-MSCloudLoginExchangeOnline } | Should -Throw '*the service is unavailable*'
            $workloadProfile.Connected | Should -BeFalse
        }
    }

    It 'Should reject an authentication type it does not support' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Get-ConnectionInformation -MockWith { return @() }
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.AuthenticationType = 'Interactive'

            { Connect-MSCloudLoginExchangeOnline } | Should -Throw 'No valid authentication type found'
        }
    }

    It 'Should rethrow a credential sign-in failure that is unrelated to MFA' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Get-ConnectionInformation -MockWith { return @() }
            Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $true }
            Mock -CommandName Connect-ExchangeOnline -MockWith { throw 'AADSTS50126: Invalid username or password.' }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.ExchangeOnline
            $workloadProfile.AuthenticationType = 'Credentials'
            $workloadProfile.Credentials = New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))

            { Connect-MSCloudLoginExchangeOnline } | Should -Throw '*AADSTS50126*'
            $workloadProfile.Connected | Should -BeFalse
        }
    }

    It 'Should retry a delegated credential sign-in through the MFA flow' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Get-ConnectionInformation -MockWith { return @() }
            Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $false }
            Mock -CommandName Connect-ExchangeOnline -MockWith {
                if ($null -ne $Credential)
                {
                    throw 'WAM Error 3399614467'
                }
            }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.ExchangeOnline
            $workloadProfile.AuthenticationType = 'CredentialsWithTenantId'
            $workloadProfile.TenantId = 'fabrikam.onmicrosoft.com'
            $workloadProfile.Credentials = New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))

            Connect-MSCloudLoginExchangeOnline

            $workloadProfile.Connected | Should -BeTrue
            $workloadProfile.MultiFactorAuthentication | Should -BeTrue
            Should -Invoke Connect-ExchangeOnline -Exactly 1 -ParameterFilter {
                $UserPrincipalName -eq 'admin@contoso.onmicrosoft.com' -and
                $DelegatedOrganization -eq 'fabrikam.onmicrosoft.com'
            }
        }
    }

    It 'Should rethrow when the MFA sign-in itself fails' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Connect-ExchangeOnline -MockWith { throw 'the sign-in window was closed' }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.ExchangeOnline
            $workloadProfile.Credentials = New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))

            { Connect-MSCloudLoginExchangeOnlineMFA -Credentials $workloadProfile.Credentials } |
                Should -Throw '*the sign-in window was closed*'
            $workloadProfile.Connected | Should -BeFalse
        }
    }
}

Describe 'Connect-MSCloudLoginSecurityCompliance failure handling' {

    BeforeEach {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
            Mock -CommandName Remove-MSCloudLoginProxyModule -MockWith { }
            Mock -CommandName Get-PSSession -MockWith { return @() }
            $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
        }
    }

    It 'Should rethrow and disconnect when the <AuthenticationType> sign-in fails' -TestCases @(
        @{ AuthenticationType = 'ServicePrincipalWithThumbprint' }
        @{ AuthenticationType = 'ServicePrincipalWithPath' }
    ) {
        param ($AuthenticationType)
        InModuleScope 'MSCloudLoginAssistant' -Parameters @{ AuthenticationType = $AuthenticationType } {
            param ($AuthenticationType)

            Mock -CommandName Connect-IPPSSession -MockWith { throw 'the compliance endpoint refused the connection' }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter
            $workloadProfile.AuthenticationType = $AuthenticationType
            $workloadProfile.ApplicationId = 'app-id'
            $workloadProfile.TenantId = 'contoso.onmicrosoft.com'
            $workloadProfile.CertificateThumbprint = 'thumbprint'
            $workloadProfile.CertificatePath = 'C:\certificates\contoso.pfx'

            { Connect-MSCloudLoginSecurityCompliance } | Should -Throw '*refused the connection*'
            $workloadProfile.Connected | Should -BeFalse
        }
    }

    It 'Should retry a delegated credential sign-in through the MFA flow' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $false }
            Mock -CommandName Connect-IPPSSession -MockWith {
                if ($null -ne $Credential)
                {
                    throw 'AADSTS50076: multi-factor authentication is required.'
                }
            }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter
            $workloadProfile.AuthenticationType = 'CredentialsWithTenantId'
            $workloadProfile.TenantId = 'contoso.onmicrosoft.com'
            $workloadProfile.ConnectionUrl = 'https://ps.compliance.protection.outlook.com/powershell-liveid/'
            $workloadProfile.AzureADAuthorizationEndpointUri = 'https://login.microsoftonline.com/organizations'
            $workloadProfile.Credentials = New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))

            Connect-MSCloudLoginSecurityCompliance

            $workloadProfile.Connected | Should -BeTrue
            $workloadProfile.MultiFactorAuthentication | Should -BeTrue
            Should -Invoke Connect-IPPSSession -Exactly 1 -ParameterFilter {
                $UserPrincipalName -eq 'admin@contoso.onmicrosoft.com' -and
                $DelegatedOrganization -eq 'contoso.onmicrosoft.com'
            }
        }
    }

    It 'Should rethrow a delegated credential sign-in failure that is unrelated to MFA' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $true }
            Mock -CommandName Connect-IPPSSession -MockWith { throw 'AADSTS50126: Invalid username or password.' }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter
            $workloadProfile.AuthenticationType = 'CredentialsWithTenantId'
            $workloadProfile.TenantId = 'contoso.onmicrosoft.com'
            $workloadProfile.AzureADAuthorizationEndpointUri = 'https://login.microsoftonline.com/organizations'
            $workloadProfile.Credentials = New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))

            { Connect-MSCloudLoginSecurityCompliance } | Should -Throw '*AADSTS50126*'
            $workloadProfile.Connected | Should -BeFalse
        }
    }

    It 'Should rethrow a credential sign-in failure that is unrelated to MFA' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $true }
            Mock -CommandName Connect-IPPSSession -MockWith { throw 'AADSTS50126: Invalid username or password.' }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter
            $workloadProfile.AuthenticationType = 'Credentials'
            $workloadProfile.Credentials = New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))

            { Connect-MSCloudLoginSecurityCompliance } | Should -Throw '*AADSTS50126*'
            $workloadProfile.Connected | Should -BeFalse
        }
    }

    Context 'When connecting with an access token' {
        BeforeAll {
            # Unsigned JSON Web Token that only carries the exp claim.
            function New-TestAccessToken([System.DateTime]$ExpiresOn)
            {
                $encode = { param ($Value) [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes(($Value | ConvertTo-Json -Compress))).TrimEnd('=').Replace('+', '-').Replace('/', '_') }
                return '{0}.{1}.signature' -f (& $encode @{ alg = 'none' }), (& $encode @{ exp = [System.DateTimeOffset]::new($ExpiresOn).ToUnixTimeSeconds() })
            }
            $script:validToken = New-TestAccessToken -ExpiresOn ([System.DateTime]::Now.AddMinutes(60))
            $script:shortLivedToken = New-TestAccessToken -ExpiresOn ([System.DateTime]::Now.AddMinutes(2))
            $script:expiredToken = New-TestAccessToken -ExpiresOn ([System.DateTime]::Now.AddMinutes(-2))
        }

        BeforeEach {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-IPPSSession -MockWith { }
                Mock -CommandName Connect-M365Tenant -MockWith { }
                $workloadProfile = $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter
                $workloadProfile.TenantId = 'contoso.onmicrosoft.com'
                $workloadProfile.ConnectionUrl = 'https://ps.compliance.protection.outlook.com/powershell-liveid/'
                $workloadProfile.AzureADAuthorizationEndpointUri = 'https://login.microsoftonline.com/organizations'
                $workloadProfile.ResourceUrl = 'https://ps.compliance.protection.outlook.com'
                $workloadProfile.EnableSearchOnlySession = $true
                $Script:MSCloudLoginCurrentLoadedModule = $null
            }
        }

        It 'Should connect through Connect-IPPSSession with the supplied access token' {
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{ Token = $script:validToken } {
                param ($Token)

                $workloadProfile = $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter
                $workloadProfile.AuthenticationType = 'AccessTokens'
                $workloadProfile.AccessTokens = @("Bearer $Token")

                Connect-MSCloudLoginSecurityCompliance

                $workloadProfile.Connected | Should -BeTrue
                $workloadProfile.TokenExpiresOn | Should -Be (Get-MSCloudLoginAccessTokenExpiry -Token $Token)
                Should -Invoke Connect-M365Tenant -Exactly 0
                Should -Invoke Connect-IPPSSession -Exactly 1 -ParameterFilter {
                    $AccessToken -eq $Token -and
                    $Organization -eq 'contoso.onmicrosoft.com' -and
                    $ConnectionUri -eq 'https://ps.compliance.protection.outlook.com/powershell-liveid/' -and
                    $AzureADAuthorizationEndpointUri -eq 'https://login.microsoftonline.com/organizations' -and
                    $EnableSearchOnlySession.IsPresent
                }
            }
        }

        It 'Should acquire a managed identity token for the resource of the environment' {
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{ Token = $script:validToken } {
                param ($Token)

                Mock -CommandName Get-AuthToken -MockWith { return $Token }.GetNewClosure()

                $workloadProfile = $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter
                $workloadProfile.AuthenticationType = 'Identity'

                Connect-MSCloudLoginSecurityCompliance

                $workloadProfile.Connected | Should -BeTrue
                Should -Invoke Get-AuthToken -Exactly 1 -ParameterFilter {
                    $Resource -eq 'https://ps.compliance.protection.outlook.com' -and $Identity.IsPresent
                }
                Should -Invoke Connect-IPPSSession -Exactly 1 -ParameterFilter {
                    $AccessToken -eq $Token -and $Organization -eq 'contoso.onmicrosoft.com'
                }
            }
        }

        It 'Should refuse a managed identity connection without a resource URL' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Get-AuthToken -MockWith { }

                $workloadProfile = $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter
                $workloadProfile.AuthenticationType = 'Identity'
                $workloadProfile.EnvironmentName = 'Custom'
                $workloadProfile.ResourceUrl = $null

                { Connect-MSCloudLoginSecurityCompliance } | Should -Throw '*No Security & Compliance resource URL*'

                $workloadProfile.Connected | Should -BeFalse
                Should -Invoke Get-AuthToken -Exactly 0
                Should -Invoke Connect-IPPSSession -Exactly 0
            }
        }

        It 'Should refuse a tenant GUID as organization for <AuthenticationType>' -TestCases @(
            @{ AuthenticationType = 'AccessTokens' }
            @{ AuthenticationType = 'Identity' }
        ) {
            param ($AuthenticationType)
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{ AuthenticationType = $AuthenticationType; Token = $script:validToken } {
                param ($AuthenticationType, $Token)

                Mock -CommandName Get-AuthToken -MockWith { return $Token }.GetNewClosure()

                $workloadProfile = $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter
                $workloadProfile.AuthenticationType = $AuthenticationType
                $workloadProfile.AccessTokens = @($Token)
                $workloadProfile.TenantId = '588132a0-32ad-4a63-b89f-0e7f2e003683'

                { Connect-MSCloudLoginSecurityCompliance } | Should -Throw '*TenantId must be the initial domain*'

                $workloadProfile.Connected | Should -BeFalse
                Should -Invoke Connect-IPPSSession -Exactly 0
            }
        }

        It 'Should refuse an expired supplied access token' {
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{ Token = $script:expiredToken } {
                param ($Token)

                $workloadProfile = $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter
                $workloadProfile.AuthenticationType = 'AccessTokens'
                $workloadProfile.AccessTokens = @($Token)

                { Connect-MSCloudLoginSecurityCompliance } | Should -Throw '*expired*Provide a new access token*'

                $workloadProfile.Connected | Should -BeFalse
                Should -Invoke Connect-IPPSSession -Exactly 0
            }
        }

        It 'Should stay disconnected when Connect-IPPSSession fails' {
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{ Token = $script:validToken } {
                param ($Token)

                Mock -CommandName Connect-IPPSSession -MockWith { throw 'the token was rejected' }

                $workloadProfile = $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter
                $workloadProfile.AuthenticationType = 'AccessTokens'
                $workloadProfile.AccessTokens = @($Token)

                { Connect-MSCloudLoginSecurityCompliance } | Should -Throw '*the token was rejected*'

                $workloadProfile.Connected | Should -BeFalse
                Should -Invoke Add-MSCloudLoginAssistantEvent -ParameterFilter {
                    $EntryType -eq 'Error' -and $Message -like '*Failed to connect to Security & Compliance with Access Token*'
                }
            }
        }

        It 'Should acquire a new managed identity token when the current one expires within five minutes' {
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{ Token = $script:validToken; ShortLivedToken = $script:shortLivedToken } {
                param ($Token, $ShortLivedToken)

                Mock -CommandName Get-AuthToken -MockWith { return $Token }.GetNewClosure()
                Mock -CommandName Restore-MSCloudLoginProxyModule -MockWith { return $true }

                $workloadProfile = $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter
                $workloadProfile.AuthenticationType = 'Identity'
                $workloadProfile.CompleteConnection($false, (Get-MSCloudLoginAccessTokenExpiry -Token $ShortLivedToken))
                $Script:MSCloudLoginCurrentLoadedModule = 'SC'

                Connect-MSCloudLoginSecurityCompliance

                $workloadProfile.Connected | Should -BeTrue
                $workloadProfile.TokenExpiresOn | Should -Be (Get-MSCloudLoginAccessTokenExpiry -Token $Token)
                Should -Invoke Get-AuthToken -Exactly 1
                Should -Invoke Connect-IPPSSession -Exactly 1 -ParameterFilter { $AccessToken -eq $Token }
            }
        }

        It 'Should keep a supplied access token connection until the token expires' {
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{ ShortLivedToken = $script:shortLivedToken } {
                param ($ShortLivedToken)

                $workloadProfile = $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter
                $workloadProfile.AuthenticationType = 'AccessTokens'
                $workloadProfile.AccessTokens = @($ShortLivedToken)
                $workloadProfile.CompleteConnection($false, (Get-MSCloudLoginAccessTokenExpiry -Token $ShortLivedToken))
                $Script:MSCloudLoginCurrentLoadedModule = 'SC'

                Connect-MSCloudLoginSecurityCompliance

                $workloadProfile.Connected | Should -BeTrue
                Should -Invoke Connect-IPPSSession -Exactly 0
            }
        }

        It 'Should report an expired supplied access token instead of reusing the connection' {
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{ Token = $script:expiredToken } {
                param ($Token)

                $workloadProfile = $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter
                $workloadProfile.AuthenticationType = 'AccessTokens'
                $workloadProfile.AccessTokens = @($Token)
                $workloadProfile.CompleteConnection($false, (Get-MSCloudLoginAccessTokenExpiry -Token $Token))
                $Script:MSCloudLoginCurrentLoadedModule = 'SC'

                { Connect-MSCloudLoginSecurityCompliance } | Should -Throw '*expired*'

                $workloadProfile.Connected | Should -BeFalse
                Should -Invoke Connect-IPPSSession -Exactly 0
            }
        }
    }

    It 'Should rethrow when the MFA sign-in itself fails' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Connect-IPPSSession -MockWith { throw 'the sign-in window was closed' }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter
            $workloadProfile.Credentials = New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))

            { Connect-MSCloudLoginSecurityComplianceMFA } | Should -Throw '*the sign-in window was closed*'
            $workloadProfile.Connected | Should -BeFalse
        }
    }

    It 'Should return immediately when the compliance proxy module is the current one' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Connect-IPPSSession -MockWith { }
            Mock -CommandName Restore-MSCloudLoginProxyModule -MockWith { return $true }

            $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.CompleteConnection()
            $Script:MSCloudLoginCurrentLoadedModule = 'SC'

            Connect-MSCloudLoginSecurityCompliance

            Should -Invoke Connect-IPPSSession -Exactly 0
            Should -Invoke Restore-MSCloudLoginProxyModule -Exactly 0
        }
    }

    It 'Should restore the compliance proxy module after an Exchange Online connection' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Connect-IPPSSession -MockWith { }
            Mock -CommandName Restore-MSCloudLoginProxyModule -MockWith { return $true }

            $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.CompleteConnection()
            $Script:MSCloudLoginCurrentLoadedModule = 'EXO'

            Connect-MSCloudLoginSecurityCompliance

            $Script:MSCloudLoginCurrentLoadedModule | Should -Be 'SC'
            $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Connected | Should -BeTrue
            Should -Invoke Restore-MSCloudLoginProxyModule -Exactly 1 -ParameterFilter { $ProbeCommand -eq 'Get-ComplianceSearch' }
            Should -Invoke Connect-IPPSSession -Exactly 0
        }
    }

    It 'Should reconnect when the compliance proxy module is no longer loaded' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Connect-IPPSSession -MockWith { }
            Mock -CommandName Restore-MSCloudLoginProxyModule -MockWith { return $false }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter
            $workloadProfile.AuthenticationType = 'ServicePrincipalWithThumbprint'
            $workloadProfile.ApplicationId = 'app-id'
            $workloadProfile.TenantId = 'contoso.onmicrosoft.com'
            $workloadProfile.CertificateThumbprint = 'thumbprint'
            $workloadProfile.CompleteConnection()
            $Script:MSCloudLoginCurrentLoadedModule = 'EXO'

            Connect-MSCloudLoginSecurityCompliance

            $Script:MSCloudLoginCurrentLoadedModule | Should -Be 'SC'
            $workloadProfile.Connected | Should -BeTrue
            Should -Invoke Connect-IPPSSession -Exactly 1
        }
    }

    It 'Should mark the compliance proxy module as current after re-importing an open session' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Connect-IPPSSession -MockWith { }
            Mock -CommandName Import-PSSession -MockWith { return 'tmpSCC_abcdefgh' }
            Mock -CommandName Import-Module -MockWith { }
            Mock -CommandName Get-PSSession -MockWith {
                return @([PSCustomObject]@{ ComputerName = 'ps.compliance.protection.outlook.com'; State = 'Opened' })
            }
            $Script:MSCloudLoginCurrentLoadedModule = 'EXO'

            Connect-MSCloudLoginSecurityCompliance

            $Script:MSCloudLoginCurrentLoadedModule | Should -Be 'SC'
            Should -Invoke Connect-IPPSSession -Exactly 0
        }
    }
}

Describe 'Connect-MSCloudLoginTeams failure handling' {

    BeforeEach {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
            $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
        }
    }

    It 'Should keep an existing connection when the session probe returns nothing usable' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $true }
            Mock -CommandName Connect-MicrosoftTeams -MockWith { }

            $Script:MSCloudLoginConnectionProfile.Teams.CompleteConnection()

            Connect-MSCloudLoginTeams

            Should -Invoke Connect-MicrosoftTeams -Exactly 0
        }
    }

    It 'Should rethrow a failing certificate based sign-in' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Get-CsTeamsCallingPolicy -MockWith { throw 'no session' }
            Mock -CommandName Connect-MicrosoftTeams -MockWith { throw 'the application is not consented' }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.Teams
            $workloadProfile.AuthenticationType = 'ServicePrincipalWithThumbprint'
            $workloadProfile.ApplicationId = 'app-id'
            $workloadProfile.TenantId = 'contoso.onmicrosoft.com'
            $workloadProfile.CertificateThumbprint = 'thumbprint'

            { Connect-MSCloudLoginTeams } | Should -Throw '*not consented*'
            $workloadProfile.Connected | Should -BeFalse
        }
    }

    It 'Should reject an authentication type it does not support' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Get-CsTeamsCallingPolicy -MockWith { throw 'no session' }
            $Script:MSCloudLoginConnectionProfile.Teams.AuthenticationType = 'Interactive'

            { Connect-MSCloudLoginTeams } | Should -Throw "*'Interactive' is not supported for workload 'MicrosoftTeams'*"
        }
    }

    It 'Should pass the government environment to the MFA sign-in' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Disconnect-MicrosoftTeams -MockWith { }
            Mock -CommandName Connect-MicrosoftTeams -MockWith { }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.Teams
            $workloadProfile.EnvironmentName = 'AzureUSGovernment'
            $workloadProfile.TenantId = 'contoso.onmicrosoft.com'

            Connect-MSCloudLoginTeamsMFA

            $workloadProfile.MultiFactorAuthentication | Should -BeTrue
            Should -Invoke Connect-MicrosoftTeams -Exactly 1 -ParameterFilter {
                $TeamsEnvironmentName -eq 'TeamsGCCH' -and $TenantId -eq 'contoso.onmicrosoft.com'
            }
        }
    }

    It 'Should rethrow when the MFA sign-in itself fails' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Disconnect-MicrosoftTeams -MockWith { }
            Mock -CommandName Connect-MicrosoftTeams -MockWith { throw 'the sign-in window was closed' }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.Teams
            $workloadProfile.EnvironmentName = 'AzureDOD'

            { Connect-MSCloudLoginTeamsMFA } | Should -Throw '*the sign-in window was closed*'
            $workloadProfile.Connected | Should -BeFalse
        }
    }
}

Describe 'Connect-MSCloudLoginPowerPlatform failure handling' {

    BeforeEach {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
            Mock -CommandName Import-Module -MockWith { }
            $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
            $Script:CloudEnvironmentInfo = $null
        }
    }

    It 'Should reject an authentication type it does not support' {
        InModuleScope 'MSCloudLoginAssistant' {
            $Script:MSCloudLoginConnectionProfile.PowerPlatform.AuthenticationType = 'Identity'

            { Connect-MSCloudLoginPowerPlatform } |
                Should -Throw "*'Identity' is not supported for workload 'PowerPlatform'*"
            $Script:MSCloudLoginConnectionProfile.PowerPlatform.Connected | Should -BeFalse
        }
    }

    It 'Should retry a service principal sign-in against the preview endpoint' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Add-PowerAppsAccount -MockWith {
                if ($Endpoint -ne 'preview')
                {
                    throw 'unknown_user_type: Unknown User Type'
                }
            }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.PowerPlatform
            $workloadProfile.AuthenticationType = 'ServicePrincipalWithThumbprint'
            $workloadProfile.ApplicationId = 'app-id'
            $workloadProfile.TenantId = 'contoso.onmicrosoft.com'
            $workloadProfile.CertificateThumbprint = 'thumbprint'

            Connect-MSCloudLoginPowerPlatform

            $workloadProfile.Connected | Should -BeTrue
            Should -Invoke Add-PowerAppsAccount -Exactly 1 -ParameterFilter {
                $Endpoint -eq 'preview' -and $CertificateThumbprint -eq 'thumbprint'
            }
        }
    }

    It 'Should give up in a non interactive session when the preview endpoint also fails' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $true }
            Mock -CommandName Add-PowerAppsAccount -MockWith { throw 'unknown_user_type: Unknown User Type' }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.PowerPlatform
            $workloadProfile.AuthenticationType = 'ServicePrincipalWithSecret'
            $workloadProfile.ApplicationId = 'app-id'
            $workloadProfile.TenantId = 'contoso.onmicrosoft.com'
            $workloadProfile.ApplicationSecret = 'secret'
            $workloadProfile.Credentials = New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))

            { Connect-MSCloudLoginPowerPlatform } | Should -Throw '*unknown_user_type*'
            $workloadProfile.Connected | Should -BeFalse
        }
    }

    It 'Should rethrow when the MFA sign-in itself fails' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Add-PowerAppsAccount -MockWith { throw 'the sign-in window was closed' }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.PowerPlatform

            { Connect-MSCloudLoginPowerPlatformMFA } | Should -Throw '*the sign-in window was closed*'
            $workloadProfile.Connected | Should -BeFalse
        }
    }
}

Describe 'Connect-MSCloudLoginAzure failure handling' {

    BeforeEach {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
            $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
        }
    }

    It 'Should reuse a live Azure context' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Connect-AzAccount -MockWith { }
            Mock -CommandName Get-AzContext -MockWith {
                return @{
                    Account     = @{ Id = '00000000-0000-0000-0000-000000000001'; Type = 'ServicePrincipal' }
                    Environment = @{ ResourceManagerUrl = 'https://management.azure.com/' }
                }
            }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.Azure
            $workloadProfile.AuthenticationType = 'ServicePrincipalWithThumbprint'
            $workloadProfile.ApplicationId = '00000000-0000-0000-0000-000000000001'
            $workloadProfile.CompleteConnection()

            Connect-MSCloudLoginAzure

            Should -Invoke Connect-AzAccount -Exactly 0
        }
    }

    It 'Should reconnect when another application connected the Azure context' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Connect-AzAccount -MockWith { }
            Mock -CommandName Get-AzContext -MockWith {
                return @{
                    Account     = @{ Id = '00000000-0000-0000-0000-000000000002'; Type = 'ServicePrincipal' }
                    Environment = @{ ResourceManagerUrl = 'https://management.azure.com/' }
                }
            }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.Azure
            $workloadProfile.AuthenticationType = 'ServicePrincipalWithThumbprint'
            $workloadProfile.ApplicationId = '00000000-0000-0000-0000-000000000001'
            $workloadProfile.CompleteConnection()

            Connect-MSCloudLoginAzure

            Should -Invoke Connect-AzAccount -Exactly 1
        }
    }

    It 'Should rethrow a credential sign-in failure that is unrelated to MFA' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $true }
            Mock -CommandName Get-AzContext -MockWith { return $null }
            Mock -CommandName Connect-AzAccount -MockWith { throw 'AADSTS50126: Invalid username or password.' }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.Azure
            $workloadProfile.AuthenticationType = 'Credentials'
            $workloadProfile.Credentials = New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))

            { Connect-MSCloudLoginAzure } | Should -Throw '*AADSTS50126*'
            $workloadProfile.Connected | Should -BeFalse
        }
    }
}

Describe 'Connect-MSCloudLoginPnP failure handling' {

    BeforeEach {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
            Mock -CommandName Import-Module -MockWith { }
            $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
        }
    }

    It 'Should return immediately when the connection can be reused' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Connect-PnPOnline -MockWith { }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.PnP
            $workloadProfile.AuthenticationType = 'ServicePrincipalWithThumbprint'
            $workloadProfile.ConnectionUrl = 'https://contoso-admin.sharepoint.com'
            $workloadProfile.CompleteConnection()

            Connect-MSCloudLoginPnP

            Should -Invoke Connect-PnPOnline -Exactly 0
        }
    }

    It 'Should continue when the Microsoft Graph authentication module cannot be imported' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Connect-PnPOnline -MockWith { }
            Mock -CommandName Get-Module -MockWith { return $null }
            Mock -CommandName Import-Module -MockWith { throw 'the module is not installed' }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.PnP
            $workloadProfile.AuthenticationType = 'ServicePrincipalWithSecret'
            $workloadProfile.ApplicationId = 'app-id'
            $workloadProfile.ApplicationSecret = 'secret'
            $workloadProfile.ConnectionUrl = 'https://contoso-admin.sharepoint.com'

            Connect-MSCloudLoginPnP

            $workloadProfile.Connected | Should -BeTrue
            Should -Invoke Add-MSCloudLoginAssistantEvent -ParameterFilter {
                $Message -like 'Failed to import Microsoft.Graph.Authentication*' -and $EntryType -eq 'Warning'
            }
        }
    }

    It 'Should load PnP.PowerShell version 1 through Windows PowerShell' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Connect-PnPOnline -MockWith { }
            Mock -CommandName Get-Module -MockWith {
                if ($ListAvailable.IsPresent)
                {
                    return @([PSCustomObject]@{
                        Name                  = 'PnP.PowerShell'
                        Version               = [System.Version]'1.12.0'
                        CompatiblePSEditions  = @('Desktop', 'Core')
                    })
                }
                if ($Name -eq 'Microsoft.Graph.Authentication')
                {
                    return @([PSCustomObject]@{ Name = 'Microsoft.Graph.Authentication' })
                }
                return $null
            }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.PnP
            $workloadProfile.AuthenticationType = 'ServicePrincipalWithSecret'
            $workloadProfile.ApplicationId = 'app-id'
            $workloadProfile.ApplicationSecret = 'secret'
            $workloadProfile.ConnectionUrl = 'https://contoso-admin.sharepoint.com'

            Connect-MSCloudLoginPnP

            $workloadProfile.Connected | Should -BeTrue
            Should -Invoke Import-Module -Exactly 1 -ParameterFilter {
                $Name -eq 'PnP.PowerShell' -and $UseWindowsPowerShell.IsPresent
            }
        }
    }

    It 'Should explain how to install PnP.PowerShell version 1 for Windows PowerShell' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Connect-PnPOnline -MockWith { }
            Mock -CommandName Get-Module -MockWith {
                if ($ListAvailable.IsPresent)
                {
                    return @([PSCustomObject]@{
                        Name                 = 'PnP.PowerShell'
                        Version              = [System.Version]'1.12.0'
                        CompatiblePSEditions = @('Core')
                    })
                }
                if ($Name -eq 'Microsoft.Graph.Authentication')
                {
                    return @([PSCustomObject]@{ Name = 'Microsoft.Graph.Authentication' })
                }
                return $null
            }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.PnP
            $workloadProfile.AuthenticationType = 'ServicePrincipalWithSecret'
            $workloadProfile.ConnectionUrl = 'https://contoso-admin.sharepoint.com'

            { Connect-MSCloudLoginPnP } | Should -Throw '*Install-Module Pnp.PowerShell -Force -Scope AllUsers*'
        }
    }

    It 'Should take the admin URL as connection URL when only the admin URL is known' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Connect-PnPOnline -MockWith { }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.PnP
            $workloadProfile.AuthenticationType = 'ServicePrincipalWithThumbprint'
            $workloadProfile.ApplicationId = 'app-id'
            $workloadProfile.TenantId = 'contoso.onmicrosoft.com'
            $workloadProfile.CertificateThumbprint = 'thumbprint'
            $workloadProfile.AdminUrl = 'https://contoso-admin.sharepoint.com'

            Connect-MSCloudLoginPnP

            $workloadProfile.ConnectionUrl | Should -Be 'https://contoso-admin.sharepoint.com'
            $workloadProfile.Connected | Should -BeTrue
        }
    }

    It 'Should fail with a clear message when the admin URL cannot be resolved' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Connect-PnPOnline -MockWith { }
            Mock -CommandName Get-SPOAdminUrl -MockWith { return '' }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.PnP
            $workloadProfile.AuthenticationType = 'Credentials'
            $workloadProfile.Credentials = New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))

            { Connect-MSCloudLoginPnP } | Should -Throw '*Unable to retrieve SharePoint Admin Url*'
        }
    }

    It 'Should reject an authentication type it does not support' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Connect-PnPOnline -MockWith { }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.PnP
            $workloadProfile.AuthenticationType = 'Interactive'
            $workloadProfile.ConnectionUrl = 'https://contoso-admin.sharepoint.com'

            { Connect-MSCloudLoginPnP } | Should -Throw "*'Interactive' is not supported for workload 'PnP'*"
        }
    }

    It 'Should fall back to the web login when the interactive MFA sign-in fails' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $false }
            Mock -CommandName Connect-PnPOnline -MockWith {
                if ($UseWebLogin.IsPresent)
                {
                    return
                }
                if ($Interactive.IsPresent)
                {
                    throw 'the interactive sign-in failed'
                }
                throw 'AADSTS50076: multi-factor authentication is required.'
            }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.PnP
            $workloadProfile.AuthenticationType = 'ServicePrincipalWithSecret'
            $workloadProfile.ApplicationId = 'app-id'
            $workloadProfile.ApplicationSecret = 'secret'
            $workloadProfile.ConnectionUrl = 'https://contoso-admin.sharepoint.com'

            Connect-MSCloudLoginPnP

            $workloadProfile.Connected | Should -BeTrue
            $workloadProfile.MultiFactorAuthentication | Should -BeTrue
            Should -Invoke Connect-PnPOnline -Exactly 1 -ParameterFilter { $UseWebLogin.IsPresent }
        }
    }

    It 'Should sign in interactively when the account cannot use the password grant' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $false }
            Mock -CommandName Connect-PnPOnline -MockWith {
                if ($Interactive.IsPresent)
                {
                    return
                }
                throw 'The sign-in name or password does not match one in the Microsoft account system.'
            }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.PnP
            $workloadProfile.AuthenticationType = 'ServicePrincipalWithSecret'
            $workloadProfile.ApplicationId = 'app-id'
            $workloadProfile.ApplicationSecret = 'secret'
            $workloadProfile.ConnectionUrl = 'https://contoso-admin.sharepoint.com'

            Connect-MSCloudLoginPnP

            $workloadProfile.Connected | Should -BeTrue
            $workloadProfile.MultiFactorAuthentication | Should -BeTrue
        }
    }

    It 'Should connect through the web login after granting management shell consent' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $false }
            Mock -CommandName Register-PnPManagementShellAccess -MockWith { }
            Mock -CommandName Connect-PnPOnline -MockWith {
                if ($UseWebLogin.IsPresent)
                {
                    return
                }
                throw 'AADSTS65001: The user or administrator has not consented to use the application with ID app-id'
            }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.PnP
            $workloadProfile.AuthenticationType = 'ServicePrincipalWithSecret'
            $workloadProfile.ApplicationId = 'app-id'
            $workloadProfile.ApplicationSecret = 'secret'
            $workloadProfile.ConnectionUrl = 'https://contoso-admin.sharepoint.com'

            Connect-MSCloudLoginPnP

            $workloadProfile.Connected | Should -BeTrue
            Should -Invoke Register-PnPManagementShellAccess -Exactly 1
        }
    }
}
