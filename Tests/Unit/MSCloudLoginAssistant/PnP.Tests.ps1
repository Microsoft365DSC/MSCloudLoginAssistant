#Requires -Modules Pester

Describe 'Connect-MSCloudLoginPnP' {
    BeforeAll {
        # Plain function stubs keep command resolution off the real SDK modules,
        # which would otherwise be discovered and imported on first use.
        Import-Module ./Tests/Unit/Stubs/Stubs.psm1 -Force -Global -WarningAction SilentlyContinue
        Import-Module ./Modules/MSCloudLoginAssistant/MSCloudLoginAssistant.psd1 -Force

        # Compile and instantiate the workload classes once here so that the cost
        # does not show up inside the first test of this file.
        InModuleScope 'MSCloudLoginAssistant' {
            $null = New-Object MSCloudLoginConnectionProfile
        }
    }

    BeforeEach {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Connect-PnPOnline -MockWith { }
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
            Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
            Mock -CommandName Get-Module -MockWith {
                param ($Name)
                if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                if ($Name -eq 'PnP.PowerShell') { return [pscustomobject]@{ Name = 'PnP.PowerShell' } }
                return $null
            }
            Mock -CommandName Import-Module -MockWith { }
            # An empty ConnectionUrl in the resolved result leaves the profile
            # without a connection URL, so the AdminUrl based path applies.
            Mock -CommandName Get-MSCloudLoginSPOUrlFromTenantId -MockWith {
                return @{
                    AdminUrl      = 'https://contoso-admin.sharepoint.com'
                    ConnectionUrl = ''
                }
            }

            $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
        }
    }

    Context 'When the connection is still reusable' {
        It 'Should return early without connecting again' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $true }

                $Script:MSCloudLoginConnectionProfile.PnP.Connected = $true

                { Connect-MSCloudLoginPnP } | Should -Not -Throw

                Should -Invoke Add-MSCloudLoginAssistantEvent -ParameterFilter {
                    $Message -like '*Already connected to PnP*'
                }
                Should -Invoke Connect-PnPOnline -Exactly 0
            }
        }
    }

    Context 'When loading the modules' {
        It 'Should log a warning when the Graph workaround import fails and not reload an already loaded PnP.PowerShell' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Import-Module -MockWith { throw 'the module could not be loaded' }

                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'AccessTokens'
                $Script:MSCloudLoginConnectionProfile.PnP.AccessTokens = @('raw-token')
                $Script:MSCloudLoginConnectionProfile.PnP.TenantId = 'contoso.onmicrosoft.com'
                $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl = 'https://contoso-admin.sharepoint.com'

                { Connect-MSCloudLoginPnP } | Should -Not -Throw

                Should -Invoke Add-MSCloudLoginAssistantEvent -ParameterFilter {
                    $Message -like '*Failed to import Microsoft.Graph.Authentication*' -and $EntryType -eq 'Warning'
                }
                Should -Invoke Import-Module -Exactly 0 -ParameterFilter {
                    $Name -eq 'PnP.PowerShell'
                }
            }
        }

        It 'Should load the Desktop edition through Windows PowerShell when only v1 is available' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Get-Module -MockWith {
                    param ($Name, $ListAvailable)
                    if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                    if ($Name -ne 'PnP.PowerShell') { return $null }
                    if (-not $ListAvailable) { return $null }
                    return @(
                        [pscustomobject]@{
                            Name                 = 'PnP.PowerShell'
                            Version              = [version]'1.10.0'
                            CompatiblePSEditions = @('Core', 'Desktop')
                        }
                    )
                }

                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'AccessTokens'
                $Script:MSCloudLoginConnectionProfile.PnP.AccessTokens = @('raw-token')
                $Script:MSCloudLoginConnectionProfile.PnP.TenantId = 'contoso.onmicrosoft.com'
                $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl = 'https://contoso-admin.sharepoint.com'

                { Connect-MSCloudLoginPnP } | Should -Not -Throw

                Should -Invoke Import-Module -Exactly 1 -ParameterFilter {
                    $Name -eq 'PnP.PowerShell' -and
                    $RequiredVersion -eq [version]'1.10.0' -and
                    $UseWindowsPowerShell.IsPresent
                }
            }
        }

        It 'Should explain that the Windows PowerShell installation is missing when no Desktop edition exists' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Get-Module -MockWith {
                    param ($Name, $ListAvailable)
                    if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                    if ($Name -ne 'PnP.PowerShell') { return $null }
                    if (-not $ListAvailable) { return $null }
                    return @(
                        [pscustomobject]@{
                            Name                 = 'PnP.PowerShell'
                            Version              = [version]'1.10.0'
                            CompatiblePSEditions = @('Core')
                        }
                    )
                }

                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'AccessTokens'
                $Script:MSCloudLoginConnectionProfile.PnP.AccessTokens = @('raw-token')
                $Script:MSCloudLoginConnectionProfile.PnP.TenantId = 'contoso.onmicrosoft.com'
                $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl = 'https://contoso-admin.sharepoint.com'

                { Connect-MSCloudLoginPnP } |
                    Should -Throw '*Powershell 7+ was detected*-UseWindowsPowerShell*not installed for Windows PowerShell*'
            }
        }
    }

    Context 'When resolving the connection URL' {
        It 'Should adopt the admin URL as connection URL when only the admin URL is set' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'AccessTokens'
                $Script:MSCloudLoginConnectionProfile.PnP.AccessTokens = @('raw-token')
                $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl = 'https://contoso-admin.sharepoint.com'

                { Connect-MSCloudLoginPnP } | Should -Not -Throw

                $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl | Should -Be 'https://contoso-admin.sharepoint.com'
                Should -Invoke Connect-PnPOnline -Exactly 1
            }
        }

        It 'Should fail with a clear message when the admin URL cannot be resolved' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Get-SPOAdminUrl -MockWith { return '' }

                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'Credentials'
                $Script:MSCloudLoginConnectionProfile.PnP.Credentials =
                    New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))

                { Connect-MSCloudLoginPnP } | Should -Throw '*Unable to retrieve SharePoint Admin Url*'
            }
        }
    }

    Context 'When connecting through the custom endpoints or the resolved admin URL' {
        It 'Should connect with <AuthenticationType> <Description>' -TestCases @(
            @{
                AuthenticationType = 'ServicePrincipalWithThumbprint'
                Description        = 'through the connection URL passing the custom endpoints'
                AdminUrl           = 'https://contoso-admin.sharepoint.contoso.local'
                Settings           = @{
                    ApplicationId         = 'app-id'
                    TenantId              = 'contoso.local'
                    CertificateThumbprint = 'thumbprint'
                    EndPoints             = @{
                        AzureADLoginEndPoint   = 'https://login.contoso.local'
                        MicrosoftGraphEndPoint = 'https://graph.contoso.local'
                    }
                    ConnectionUrl         = 'https://contoso-admin.sharepoint.contoso.local'
                    PnPAzureEnvironment   = 'Custom'
                }
                ExpectedParameters = {
                    $Url -eq 'https://contoso-admin.sharepoint.contoso.local' -and
                    $AzureEnvironment -eq 'Custom' -and
                    $AzureADLoginEndPoint -eq 'https://login.contoso.local' -and
                    $MicrosoftGraphEndPoint -eq 'https://graph.contoso.local'
                }
            }
            @{
                AuthenticationType = 'ServicePrincipalWithThumbprint'
                Description        = 'through the admin URL using the tenant GUID instead of the tenant name for AzureChinaCloud'
                AdminUrl           = 'https://contoso-admin.sharepoint.cn'
                Settings           = @{
                    ApplicationId         = 'app-id'
                    TenantId              = 'contoso.partner.onmschina.cn'
                    TenantGUID            = '22222222-2222-2222-2222-222222222222'
                    CertificateThumbprint = 'thumbprint'
                    EnvironmentName       = 'AzureChinaCloud'
                    PnPAzureEnvironment   = 'China'
                }
                ExpectedParameters = {
                    $Tenant -eq '22222222-2222-2222-2222-222222222222' -and
                    $Url -eq 'https://contoso-admin.sharepoint.cn' -and
                    $AzureEnvironment -eq 'China'
                }
            }
            @{
                AuthenticationType = 'ServicePrincipalWithThumbprint'
                Description        = 'through the admin URL passing the custom endpoints'
                AdminUrl           = 'https://contoso-admin.sharepoint.contoso.local'
                Settings           = @{
                    ApplicationId         = 'app-id'
                    TenantId              = 'contoso.local'
                    CertificateThumbprint = 'thumbprint'
                    EndPoints             = @{
                        AzureADLoginEndPoint   = 'https://login.contoso.local'
                        MicrosoftGraphEndPoint = 'https://graph.contoso.local'
                    }
                    PnPAzureEnvironment   = 'Custom'
                }
                ExpectedParameters = {
                    $Url -eq 'https://contoso-admin.sharepoint.contoso.local' -and
                    $AzureADLoginEndPoint -eq 'https://login.contoso.local' -and
                    $MicrosoftGraphEndPoint -eq 'https://graph.contoso.local'
                }
            }
            @{
                AuthenticationType = 'ServicePrincipalWithPath'
                Description        = 'through the admin URL with the certificate path'
                AdminUrl           = 'https://contoso-admin.sharepoint.com'
                Settings           = @{
                    ApplicationId       = 'app-id'
                    CertificatePath     = 'C:\certs\contoso.pfx'
                    CertificatePassword = ConvertTo-SecureString 'cert-password' -AsPlainText -Force
                }
                ExpectedParameters = { $Url -eq 'https://contoso-admin.sharepoint.com' -and $CertificatePath -eq 'C:\certs\contoso.pfx' }
            }
            @{
                AuthenticationType = 'ServicePrincipalWithSecret'
                Description        = 'through the admin URL with the client secret'
                AdminUrl           = 'https://contoso-admin.sharepoint.com'
                Settings           = @{ ApplicationId = 'app-id'; ApplicationSecret = 'secret' }
                ExpectedParameters = { $ClientSecret -eq 'secret' -and $Url -eq 'https://contoso-admin.sharepoint.com' }
            }
            @{
                AuthenticationType = 'CredentialsWithApplicationId'
                Description        = 'through the admin URL with the credential and application id'
                AdminUrl           = 'https://contoso-admin.sharepoint.com'
                Settings           = @{
                    ApplicationId = 'app-id'
                    Credentials   = New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))
                }
                ExpectedParameters = {
                    $ClientId -eq 'app-id' -and $null -ne $Credentials -and
                    $Url -eq 'https://contoso-admin.sharepoint.com'
                }
            }
            @{
                AuthenticationType = 'AccessTokens'
                Description        = 'through the admin URL with the resolved access token'
                AdminUrl           = 'https://contoso-admin.sharepoint.com'
                Settings           = @{ AccessTokens = @('raw-token') }
                ExpectedParameters = { $AccessToken -eq 'resolved-token' -and $Url -eq 'https://contoso-admin.sharepoint.com' }
            }
        ) {
            InModuleScope 'MSCloudLoginAssistant' -Parameters $_ {
                param ($AuthenticationType, $AdminUrl, $Settings, $ExpectedParameters)

                Mock -CommandName Get-MSCloudLoginSPOUrlFromTenantId -MockWith {
                    return @{
                        AdminUrl      = $AdminUrl
                        ConnectionUrl = ''
                    }
                }
                Mock -CommandName Get-MSCloudLoginAccessTokenValue -MockWith { return 'resolved-token' }

                $pnpProfile = $Script:MSCloudLoginConnectionProfile.PnP
                $pnpProfile.AuthenticationType = $AuthenticationType
                $pnpProfile.TenantId = 'contoso.onmicrosoft.com'
                $pnpProfile.EnvironmentName = 'AzureCloud'
                $pnpProfile.PnPAzureEnvironment = 'Production'
                foreach ($setting in $Settings.GetEnumerator())
                {
                    $pnpProfile.($setting.Key) = $setting.Value
                }

                { Connect-MSCloudLoginPnP } | Should -Not -Throw

                $pnpProfile.Connected | Should -BeTrue
                $pnpProfile.AdminUrl | Should -Be $AdminUrl
                Should -Invoke Connect-PnPOnline -Exactly 1 -ParameterFilter ([scriptblock]::Create($ExpectedParameters))
            }
        }

        It 'Should connect with credentials through the admin URL when the type resolves late' {
            InModuleScope 'MSCloudLoginAssistant' {
                # The authentication type flips to Credentials after the URLs have been
                # resolved, so the connection URL stays empty and the admin URL based
                # credentials path applies.
                Mock -CommandName Get-MSCloudLoginSPOUrlFromTenantId -MockWith {
                    $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'Credentials'
                    return @{
                        AdminUrl      = 'https://contoso-admin.sharepoint.com'
                        ConnectionUrl = ''
                    }
                }

                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'AccessTokens'
                $Script:MSCloudLoginConnectionProfile.PnP.AccessTokens = @('raw-token')
                $Script:MSCloudLoginConnectionProfile.PnP.TenantId = 'contoso.onmicrosoft.com'
                $Script:MSCloudLoginConnectionProfile.PnP.EnvironmentName = 'AzureCloud'
                $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment = 'Production'

                { Connect-MSCloudLoginPnP } | Should -Not -Throw

                $Script:MSCloudLoginConnectionProfile.PnP.Connected | Should -BeTrue
                Should -Invoke Add-MSCloudLoginAssistantEvent -ParameterFilter {
                    $Message -like '*using SPOManagementShell and AdminUrl*'
                }
                Should -Invoke Connect-PnPOnline -Exactly 1 -ParameterFilter {
                    $ClientId -eq '9bc3ab49-b65d-410a-85ad-de819febfddc' -and
                    $Url -eq 'https://contoso-admin.sharepoint.com'
                }
            }
        }

        It 'Should request a managed identity token for the resolved admin URL' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Get-AuthToken -MockWith { return 'managed-identity-token' }

                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'Identity'
                $Script:MSCloudLoginConnectionProfile.PnP.TenantId = 'contoso.onmicrosoft.com'
                $Script:MSCloudLoginConnectionProfile.PnP.EnvironmentName = 'AzureCloud'
                $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment = 'Production'

                { Connect-MSCloudLoginPnP } | Should -Not -Throw

                $Script:MSCloudLoginConnectionProfile.PnP.Connected | Should -BeTrue
                Should -Invoke Get-AuthToken -Exactly 1 -ParameterFilter {
                    $Resource -eq 'https://contoso-admin.sharepoint.com' -and $Identity.IsPresent
                }
                Should -Invoke Connect-PnPOnline -Exactly 1 -ParameterFilter {
                    $AccessToken -eq 'managed-identity-token' -and $Url -eq 'https://contoso-admin.sharepoint.com'
                }
            }
        }
    }

    Context 'When the authentication type is not supported' {
        It 'Should throw an error naming the unsupported type' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'Interactive'
                $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl = 'https://contoso-admin.sharepoint.com'

                { Connect-MSCloudLoginPnP } |
                    Should -Throw "*Authentication type 'Interactive' is not supported for workload 'PnP'*"
            }
        }
    }

    Context 'When the sign-in requires MFA' {
        BeforeEach {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $false }

                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'AccessTokens'
                $Script:MSCloudLoginConnectionProfile.PnP.AccessTokens = @('raw-token')
                $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl = 'https://contoso-admin.sharepoint.com'
                $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment = 'Production'
            }
        }

        It 'Should fall back to the web login when the interactive attempt fails' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:pnpConnectCalls = 0
                Mock -CommandName Connect-PnPOnline -MockWith {
                    $Script:pnpConnectCalls++
                    switch ($Script:pnpConnectCalls)
                    {
                        1 { throw 'AADSTS50076: multi-factor authentication is required' }
                        2 { throw 'the interactive window was dismissed' }
                        default { }
                    }
                }

                { Connect-MSCloudLoginPnP } | Should -Not -Throw

                $Script:MSCloudLoginConnectionProfile.PnP.Connected | Should -BeTrue
                $Script:MSCloudLoginConnectionProfile.PnP.MultiFactorAuthentication | Should -BeTrue
                $Script:pnpConnectCalls | Should -Be 3
                Should -Invoke Connect-PnPOnline -Exactly 1 -ParameterFilter {
                    $UseWebLogin.IsPresent
                }
            }
        }

        It 'Should surface the failure when every fallback fails' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-PnPOnline -MockWith { throw 'AADSTS50076: multi-factor authentication is required' }

                { Connect-MSCloudLoginPnP } | Should -Throw '*multi-factor authentication is required*'

                $Script:MSCloudLoginConnectionProfile.PnP.Connected | Should -BeFalse
                Should -Invoke Add-MSCloudLoginAssistantEvent -ParameterFilter {
                    $Message -like '*Failed to connect to PnP interactively after MFA-required error*' -and $EntryType -eq 'Error'
                }
            }
        }
    }

    Context 'When the credentials are rejected because of MFA' {
        BeforeEach {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $false }
            }
        }

        It 'Should retry interactively against the <UrlSource> URL on the password mismatch error' -TestCases @(
            @{
                UrlSource   = 'connection'
                Settings    = @{
                    AuthenticationType = 'Credentials'
                    Credentials        = New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))
                    ConnectionUrl      = 'https://contoso.sharepoint.com'
                }
                ExpectedUrl = 'https://contoso.sharepoint.com'
            }
            @{
                UrlSource   = 'resolved admin'
                Settings    = @{
                    AuthenticationType = 'AccessTokens'
                    AccessTokens       = @('raw-token')
                    TenantId           = 'contoso.onmicrosoft.com'
                    EnvironmentName    = 'AzureCloud'
                }
                ExpectedUrl = 'https://contoso-admin.sharepoint.com'
            }
        ) {
            InModuleScope 'MSCloudLoginAssistant' -Parameters $_ {
                param ($Settings, $ExpectedUrl)

                $Script:pnpConnectCalls = 0
                Mock -CommandName Connect-PnPOnline -MockWith {
                    $Script:pnpConnectCalls++
                    if ($Script:pnpConnectCalls -eq 1)
                    {
                        throw 'The sign-in name or password does not match one in the Microsoft account system.'
                    }
                }

                $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment = 'Production'
                foreach ($setting in $Settings.GetEnumerator())
                {
                    $Script:MSCloudLoginConnectionProfile.PnP.($setting.Key) = $setting.Value
                }

                { Connect-MSCloudLoginPnP } | Should -Not -Throw

                $Script:MSCloudLoginConnectionProfile.PnP.Connected | Should -BeTrue
                $Script:MSCloudLoginConnectionProfile.PnP.MultiFactorAuthentication | Should -BeTrue
                $Script:pnpConnectCalls | Should -Be 2
                Should -Invoke Connect-PnPOnline -Exactly 1 -ParameterFilter {
                    $Interactive.IsPresent -and $Url -eq $ExpectedUrl
                }
            }
        }

        It 'Should surface the failure when the interactive retry also fails' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-PnPOnline -MockWith {
                    throw 'The sign-in name or password does not match one in the Microsoft account system.'
                }

                $pnpProfile = $Script:MSCloudLoginConnectionProfile.PnP
                $pnpProfile.AuthenticationType = 'Credentials'
                $pnpProfile.Credentials =
                    New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))
                $pnpProfile.ConnectionUrl = 'https://contoso-admin.sharepoint.com'
                $pnpProfile.PnPAzureEnvironment = 'Production'

                { Connect-MSCloudLoginPnP } | Should -Throw '*sign-in name or password does not match*'

                $pnpProfile.Connected | Should -BeFalse
                Should -Invoke Add-MSCloudLoginAssistantEvent -ParameterFilter {
                    $Message -like '*Failed to connect to PnP interactively:*' -and $EntryType -eq 'Error'
                }
            }
        }
    }

    Context 'When the application has not been consented' {
        It 'Should register the management shell access and reconnect via web login' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $false }
                $Script:pnpConnectCalls = 0
                Mock -CommandName Register-PnPManagementShellAccess -MockWith { }
                Mock -CommandName Connect-PnPOnline -MockWith {
                    $Script:pnpConnectCalls++
                    if ($Script:pnpConnectCalls -eq 1)
                    {
                        throw 'AADSTS65001: The user or administrator has not consented to use the application with ID abc'
                    }
                }

                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'AccessTokens'
                $Script:MSCloudLoginConnectionProfile.PnP.AccessTokens = @('raw-token')
                $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl = 'https://contoso-admin.sharepoint.com'

                { Connect-MSCloudLoginPnP } | Should -Not -Throw

                $Script:MSCloudLoginConnectionProfile.PnP.Connected | Should -BeTrue
                Should -Invoke Register-PnPManagementShellAccess -Exactly 1
                Should -Invoke Connect-PnPOnline -Exactly 1 -ParameterFilter {
                    $UseWebLogin.IsPresent
                }
            }
        }
    }
}

AfterAll {
    Remove-Module MSCloudLoginAssistant
}
