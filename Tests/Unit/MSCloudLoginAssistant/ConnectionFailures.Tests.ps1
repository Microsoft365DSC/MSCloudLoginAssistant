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

    It 'Should return immediately when the workload is already connected with <Description>' -TestCases @(
        @{ Description = 'all cmdlets'; CmdletsToLoad = @(); LoadedAllCmdlets = $true }
        @{ Description = 'the requested cmdlets'; CmdletsToLoad = @('Get-Mailbox'); LoadedAllCmdlets = $false }
    ) {
        param ($CmdletsToLoad, $LoadedAllCmdlets)
        InModuleScope 'MSCloudLoginAssistant' -Parameters @{ CmdletsToLoad = $CmdletsToLoad; LoadedAllCmdlets = $LoadedAllCmdlets } {
            param ($CmdletsToLoad, $LoadedAllCmdlets)
            Mock -CommandName Get-ConnectionInformation -MockWith {
                return @([PSCustomObject]@{
                        Name         = 'ExchangeOnline_1'
                        AppId        = 'app-id'
                        ModuleName   = (Get-Command -Name Get-OrganizationConfig).Module.ModuleBase
                        IsEopSession = $false
                    })
            }
            Mock -CommandName Connect-ExchangeOnline -MockWith { }
            Mock -CommandName Restore-MSCloudLoginProxyModule -MockWith { return $true }

            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ApplicationId = 'app-id'
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CompleteConnection()
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CmdletsToLoad = $CmdletsToLoad
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.LoadedCmdlets = @('Get-Mailbox', 'Get-OrganizationConfig')
            $Script:MSCloudLoginConnectionProfile.ExchangeOnline.LoadedAllCmdlets = $LoadedAllCmdlets
            $Script:MSCloudLoginCurrentLoadedModule = 'EXO'

            Connect-MSCloudLoginExchangeOnline

            Should -Invoke Connect-ExchangeOnline -Exactly 0
            Should -Invoke Restore-MSCloudLoginProxyModule -Exactly 0
        }
    }

    It 'Should reconnect when the loaded proxy module belongs to another application' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Get-ConnectionInformation -MockWith {
                return @([PSCustomObject]@{
                        Name         = 'ExchangeOnline_1'
                        AppId        = 'other-app-id'
                        Organization = 'contoso.onmicrosoft.com'
                        ModuleName   = (Get-Command -Name Get-OrganizationConfig).Module.ModuleBase
                        IsEopSession = $false
                    })
            }
            Mock -CommandName Connect-ExchangeOnline -MockWith { }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.ExchangeOnline
            $workloadProfile.AuthenticationType = 'ServicePrincipalWithThumbprint'
            $workloadProfile.ApplicationId = 'app-id'
            $workloadProfile.TenantId = 'contoso.onmicrosoft.com'
            $workloadProfile.CertificateThumbprint = 'thumbprint'
            $workloadProfile.LoadedAllCmdlets = $true
            $Script:MSCloudLoginCurrentLoadedModule = 'EXO'

            Connect-MSCloudLoginExchangeOnline

            Should -Invoke Connect-ExchangeOnline -Exactly 1 -ParameterFilter { $AppId -eq 'app-id' }
        }
    }

    It 'Should connect again when the module of the session matching the <MatchedBy> cannot be imported' -TestCases @(
        @{ MatchedBy = 'application'; AuthenticationType = 'ServicePrincipalWithThumbprint' }
        @{ MatchedBy = 'user principal name'; AuthenticationType = 'Credentials' }
    ) {
        param ($AuthenticationType)
        InModuleScope 'MSCloudLoginAssistant' -Parameters @{ AuthenticationType = $AuthenticationType } {
            param ($AuthenticationType)
            Mock -CommandName Get-ConnectionInformation -MockWith {
                return @([PSCustomObject]@{
                        Name              = 'ExchangeOnline_1'
                        AppId             = 'app-id'
                        Organization      = 'contoso.onmicrosoft.com'
                        UserPrincipalName = 'admin@contoso.onmicrosoft.com'
                        ModuleName        = 'C:\Temp\tmpEXO_deleted'
                        IsEopSession      = $false
                    })
            }
            Mock -CommandName Import-Module -MockWith { throw 'The member FormatsToProcess in the module manifest is not valid' } -ParameterFilter { $Name -eq 'C:\Temp\tmpEXO_deleted' }
            Mock -CommandName Connect-ExchangeOnline -MockWith { }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.ExchangeOnline
            $workloadProfile.AuthenticationType = $AuthenticationType
            if ($AuthenticationType -eq 'Credentials')
            {
                $workloadProfile.Credentials = New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))
            }
            else
            {
                $workloadProfile.ApplicationId = 'app-id'
                $workloadProfile.TenantId = 'contoso.onmicrosoft.com'
                $workloadProfile.CertificateThumbprint = 'thumbprint'
            }

            { Connect-MSCloudLoginExchangeOnline } | Should -Not -Throw
            Should -Invoke Add-MSCloudLoginAssistantEvent -ParameterFilter { $Message -like 'Could not import the module of the active session:*FormatsToProcess*' }

            $workloadProfile.Connected | Should -BeTrue
            Should -Invoke Connect-ExchangeOnline -Exactly 1
        }
    }

    It 'Should switch back to the Exchange Online proxy module and reconnect only when the restore fails (<CurrentLoadedModule> loaded last, restored: <Restored>)' -TestCases @(
        @{ CurrentLoadedModule = 'SC'; Restored = $true; ExpectedConnections = 0 }
        @{ CurrentLoadedModule = 'SC'; Restored = $false; ExpectedConnections = 1 }
        @{ CurrentLoadedModule = 'EXO'; Restored = $false; ExpectedConnections = 1 }
    ) {
        param ($CurrentLoadedModule, $Restored, $ExpectedConnections)
        InModuleScope 'MSCloudLoginAssistant' -Parameters @{ CurrentLoadedModule = $CurrentLoadedModule; Restored = $Restored; ExpectedConnections = $ExpectedConnections } {
            param ($CurrentLoadedModule, $Restored, $ExpectedConnections)

            Mock -CommandName Get-ConnectionInformation -MockWith { return @() }
            Mock -CommandName Connect-ExchangeOnline -MockWith { }
            Mock -CommandName Restore-MSCloudLoginProxyModule -MockWith { return $Restored }.GetNewClosure()

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
            Should -Invoke Restore-MSCloudLoginProxyModule -Exactly 1 -ParameterFilter { $ProbeCommand -eq 'Get-OrganizationConfig' }
            Should -Invoke Connect-ExchangeOnline -Exactly $ExpectedConnections
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
        @{ AuthenticationType = 'ServicePrincipalWithThumbprint'; ExpectedError = '*AADSTS50126*'; ExpectedEvent = $null }
        @{ AuthenticationType = 'Identity'; ExpectedError = '*AADSTS50126*'; ExpectedEvent = $null }
        @{ AuthenticationType = 'AccessTokens'; ExpectedError = '*AADSTS50126*'; ExpectedEvent = $null }
        @{ AuthenticationType = 'Credentials'; ExpectedError = '*AADSTS50126*'; ExpectedEvent = $null }
        @{ AuthenticationType = 'CredentialsWithTenantId'; ExpectedError = '*AADSTS50126*'; ExpectedEvent = '*Failed to connect to Exchange Online with Credentials and TenantId*' }
        @{ AuthenticationType = 'ServicePrincipalWithSecret'; ExpectedError = '*No valid authentication type found*'; ExpectedEvent = $null }
    ) {
        param ($AuthenticationType, $ExpectedError, $ExpectedEvent)
        InModuleScope 'MSCloudLoginAssistant' -Parameters @{ AuthenticationType = $AuthenticationType; ExpectedError = $ExpectedError; ExpectedEvent = $ExpectedEvent } {
            param ($AuthenticationType, $ExpectedError, $ExpectedEvent)

            Mock -CommandName Get-ConnectionInformation -MockWith { return @() }
            Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $true }
            Mock -CommandName Connect-ExchangeOnline -MockWith { throw 'AADSTS50126: Invalid username or password.' }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.ExchangeOnline
            $workloadProfile.AuthenticationType = $AuthenticationType
            $workloadProfile.ApplicationId = 'app-id'
            $workloadProfile.TenantId = 'contoso.onmicrosoft.com'
            $workloadProfile.CertificateThumbprint = 'thumbprint'
            $workloadProfile.AccessTokens = @('token')
            $workloadProfile.Credentials = New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))

            { Connect-MSCloudLoginExchangeOnline } | Should -Throw $ExpectedError
            $workloadProfile.Connected | Should -BeFalse
            if ($null -ne $ExpectedEvent)
            {
                Should -Invoke Add-MSCloudLoginAssistantEvent -ParameterFilter { $EntryType -eq 'Error' -and $Message -like $ExpectedEvent }
            }
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

    It 'Should rethrow and disconnect when the <AuthenticationType> sign-in fails for a reason unrelated to MFA' -TestCases @(
        @{ AuthenticationType = 'ServicePrincipalWithThumbprint' }
        @{ AuthenticationType = 'ServicePrincipalWithPath' }
        @{ AuthenticationType = 'CredentialsWithTenantId' }
        @{ AuthenticationType = 'Credentials' }
    ) {
        param ($AuthenticationType)
        InModuleScope 'MSCloudLoginAssistant' -Parameters @{ AuthenticationType = $AuthenticationType } {
            param ($AuthenticationType)

            Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $true }
            Mock -CommandName Connect-IPPSSession -MockWith { throw 'AADSTS50126: Invalid username or password.' }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter
            $workloadProfile.AuthenticationType = $AuthenticationType
            $workloadProfile.ApplicationId = 'app-id'
            $workloadProfile.TenantId = 'contoso.onmicrosoft.com'
            $workloadProfile.CertificateThumbprint = 'thumbprint'
            $workloadProfile.CertificatePath = 'C:\certificates\contoso.pfx'
            $workloadProfile.AzureADAuthorizationEndpointUri = 'https://login.microsoftonline.com/organizations'
            $workloadProfile.Credentials = New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))

            { Connect-MSCloudLoginSecurityCompliance } | Should -Throw '*AADSTS50126*'
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

        It 'Should refuse an expired supplied access token instead of connecting or reusing the connection' {
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{ Token = $script:expiredToken } {
                param ($Token)

                $workloadProfile = $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter
                $workloadProfile.AuthenticationType = 'AccessTokens'
                $workloadProfile.AccessTokens = @($Token)

                { Connect-MSCloudLoginSecurityCompliance } | Should -Throw '*expired*Provide a new access token*'

                $workloadProfile.CompleteConnection($false, (Get-MSCloudLoginAccessTokenExpiry -Token $Token))
                $Script:MSCloudLoginCurrentLoadedModule = 'SC'

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

    It 'Should switch back to the compliance proxy module after an Exchange Online connection and reconnect only when the restore fails (restored: <Restored>)' -TestCases @(
        @{ Restored = $true; ExpectedConnections = 0 }
        @{ Restored = $false; ExpectedConnections = 1 }
    ) {
        param ($Restored, $ExpectedConnections)
        InModuleScope 'MSCloudLoginAssistant' -Parameters @{ Restored = $Restored; ExpectedConnections = $ExpectedConnections } {
            param ($Restored, $ExpectedConnections)

            Mock -CommandName Connect-IPPSSession -MockWith { }
            Mock -CommandName Restore-MSCloudLoginProxyModule -MockWith { return $Restored }.GetNewClosure()

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
            Should -Invoke Restore-MSCloudLoginProxyModule -Exactly 1 -ParameterFilter { $ProbeCommand -eq 'Get-ComplianceSearch' }
            Should -Invoke Connect-IPPSSession -Exactly $ExpectedConnections
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
}

Describe 'Connect-MSCloudLoginAzure failure handling' {

    BeforeEach {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
            $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
        }
    }

    It 'Should connect <ExpectedConnections> time(s) when <Description> connected the live Azure context' -TestCases @(
        @{ Description = 'the same application'; ContextAccountId = '00000000-0000-0000-0000-000000000001'; ExpectedConnections = 0 }
        @{ Description = 'another application'; ContextAccountId = '00000000-0000-0000-0000-000000000002'; ExpectedConnections = 1 }
    ) {
        param ($ContextAccountId, $ExpectedConnections)
        InModuleScope 'MSCloudLoginAssistant' -Parameters @{ ContextAccountId = $ContextAccountId; ExpectedConnections = $ExpectedConnections } {
            param ($ContextAccountId, $ExpectedConnections)

            Mock -CommandName Connect-AzAccount -MockWith { }
            Mock -CommandName Get-AzContext -MockWith {
                return @{
                    Account     = @{ Id = $ContextAccountId; Type = 'ServicePrincipal' }
                    Environment = @{ ResourceManagerUrl = 'https://management.azure.com/' }
                }
            }.GetNewClosure()

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.Azure
            $workloadProfile.AuthenticationType = 'ServicePrincipalWithThumbprint'
            $workloadProfile.ApplicationId = '00000000-0000-0000-0000-000000000001'
            $workloadProfile.CompleteConnection()

            Connect-MSCloudLoginAzure

            Should -Invoke Connect-AzAccount -Exactly $ExpectedConnections
        }
    }
}

Describe 'Connect-MSCloudLogin* MFA failure handling shared by the workloads' {

    BeforeEach {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
            $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
        }
    }

    It 'Should rethrow and stay disconnected when the <Workload> MFA sign-in itself fails' -TestCases @(
        @{ Workload = 'ExchangeOnline'; Command = 'Connect-MSCloudLoginExchangeOnlineMFA'; PassCredentials = $true }
        @{ Workload = 'SecurityComplianceCenter'; Command = 'Connect-MSCloudLoginSecurityComplianceMFA'; PassCredentials = $false }
        @{ Workload = 'PowerPlatform'; Command = 'Connect-MSCloudLoginPowerPlatformMFA'; PassCredentials = $false }
    ) {
        param ($Workload, $Command, $PassCredentials)
        InModuleScope 'MSCloudLoginAssistant' -Parameters @{ Workload = $Workload; Command = $Command; PassCredentials = $PassCredentials } {
            param ($Workload, $Command, $PassCredentials)

            Mock -CommandName Connect-ExchangeOnline -MockWith { throw 'the sign-in window was closed' }
            Mock -CommandName Connect-IPPSSession -MockWith { throw 'the sign-in window was closed' }
            Mock -CommandName Add-PowerAppsAccount -MockWith { throw 'the sign-in window was closed' }

            $workloadProfile = $Script:MSCloudLoginConnectionProfile.$Workload
            $workloadProfile.Credentials = New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))
            $parameters = @{}
            if ($PassCredentials)
            {
                $parameters['Credentials'] = $workloadProfile.Credentials
            }

            { & $Command @parameters } | Should -Throw '*the sign-in window was closed*'
            $workloadProfile.Connected | Should -BeFalse
        }
    }
}
