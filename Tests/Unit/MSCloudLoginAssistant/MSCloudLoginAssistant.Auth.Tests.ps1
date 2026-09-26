#Requires -Modules Pester

BeforeAll {
    # Ensure the Graph dependency check passes during module import.
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

    $moduleRoot = Resolve-Path (Join-Path $PSScriptRoot '..\..\..\Modules\MSCloudLoginAssistant')
    Import-Module (Join-Path $moduleRoot 'MSCloudLoginAssistant.psd1') -Force
}

AfterAll {
    if ($script:tempModuleBase -and (Test-Path $script:tempModuleBase))
    {
        Remove-Item -Path $script:tempModuleBase -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# ---------------------------------------------------------------------------
# Get-AuthToken
# ---------------------------------------------------------------------------
Describe 'Get-AuthToken' {

    BeforeAll {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

            $script:testPfxPath = Join-Path $env:TEMP ('msla-cert-{0}.pfx' -f ([guid]::NewGuid().ToString('N')))
            $rsa = [System.Security.Cryptography.RSA]::Create(2048)
            $req = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new(
                [System.Security.Cryptography.X509Certificates.X500DistinguishedName]::new('CN=MSCloudLoginAssistantTest'),
                $rsa,
                [System.Security.Cryptography.HashAlgorithmName]::SHA256,
                [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
            $cert = $req.CreateSelfSigned([System.DateTimeOffset]::Now.AddDays(-1), [System.DateTimeOffset]::Now.AddDays(1))
            $script:cert = $cert
            $script:testThumbprint = $cert.Thumbprint
            [System.IO.File]::WriteAllBytes($script:testPfxPath, $cert.Export(
                [System.Security.Cryptography.X509Certificates.X509ContentType]::Pfx, 'testpwd'))

            $store = [System.Security.Cryptography.X509Certificates.X509Store]::new('My', 'CurrentUser')
            $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
            $store.Add([System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
                $script:testPfxPath,
                'testpwd',
                [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::UserKeySet))
            $store.Close()

            # $cert.Dispose() # Do not dispose!
            $rsa.Dispose()

            $script:oldAzpsHost = $env:AZUREPS_HOST_ENVIRONMENT
            $script:oldIdentityEndpoint = $env:IDENTITY_ENDPOINT
            $script:oldIdentityHeader = $env:IDENTITY_HEADER
            $script:oldImdsEndpoint = $env:IMDS_ENDPOINT
        }
    }

    AfterAll {
        InModuleScope 'MSCloudLoginAssistant' {
            if ($script:testThumbprint)
            {
                $store = [System.Security.Cryptography.X509Certificates.X509Store]::new('My', 'CurrentUser')
                $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
                $store.Certificates | Where-Object { $_.Thumbprint -eq $script:testThumbprint } |
                    ForEach-Object { $store.Remove($_) }
                $store.Close()
            }
            if ($script:testPfxPath -and (Test-Path $script:testPfxPath))
            {
                Remove-Item $script:testPfxPath -Force -ErrorAction SilentlyContinue
            }
            $env:AZUREPS_HOST_ENVIRONMENT = $script:oldAzpsHost
            $env:IDENTITY_ENDPOINT = $script:oldIdentityEndpoint
            $env:IDENTITY_HEADER = $script:oldIdentityHeader
            $env:IMDS_ENDPOINT = $script:oldImdsEndpoint
        }
    }

    Context 'When using managed identity on an Azure VM or in Azure Automation' {
        It 'Should return the access token from the <Endpoint> endpoint' -TestCases @(
            @{ Endpoint = 'instance metadata'; HostEnvironment = ''; IdentityEndpoint = ''; IdentityHeader = ''; ExpectedUri = 'http://169.254.169.254/metadata/identity/oauth2/token*&resource=https://graph.microsoft.com'; ExpectedIdentityHeader = $null; ExpectedBodyResource = $null }
            @{ Endpoint = 'Azure Automation identity'; HostEnvironment = 'AzureAutomation_Test'; IdentityEndpoint = 'http://localhost:9999/metadata'; IdentityHeader = 'secret-header'; ExpectedUri = 'http://localhost:9999/metadata'; ExpectedIdentityHeader = 'secret-header'; ExpectedBodyResource = 'https://graph.microsoft.com' }
        ) {
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{ HostEnvironment = $HostEnvironment; IdentityEndpoint = $IdentityEndpoint; IdentityHeader = $IdentityHeader; ExpectedUri = $ExpectedUri; ExpectedIdentityHeader = $ExpectedIdentityHeader; ExpectedBodyResource = $ExpectedBodyResource } {
                param ($HostEnvironment, $IdentityEndpoint, $IdentityHeader, $ExpectedUri, $ExpectedIdentityHeader, $ExpectedBodyResource)
                $env:AZUREPS_HOST_ENVIRONMENT = $HostEnvironment
                $env:IDENTITY_ENDPOINT = $IdentityEndpoint
                $env:IDENTITY_HEADER = $IdentityHeader
                $env:IMDS_ENDPOINT = ''
                Mock -CommandName Invoke-RestMethod -MockWith {
                    return @{ access_token = 'identity-token' }
                }
                $result = Get-AuthToken -Identity -Resource 'https://graph.microsoft.com'
                $result | Should -Be 'identity-token'
                Should -Invoke Invoke-RestMethod -Exactly 1 -ParameterFilter {
                    $Uri -like $ExpectedUri -and
                    $Headers.Metadata -and
                    $Headers.'X-IDENTITY-HEADER' -eq $ExpectedIdentityHeader -and
                    $Body.resource -eq $ExpectedBodyResource
                }
            }
        }
    }

    Context 'When using managed identity on an Azure Arc device' {
        It 'Should throw when the secret file cannot be determined' {
            InModuleScope 'MSCloudLoginAssistant' {
                $env:AZUREPS_HOST_ENVIRONMENT = ''
                $env:IDENTITY_ENDPOINT = 'http://localhost:40342/metadata'
                $env:IDENTITY_HEADER = ''
                $env:IMDS_ENDPOINT = 'http://localhost:40342'
                Mock -CommandName Invoke-WebRequest -MockWith {
                    return [PSCustomObject]@{ StatusCode = 200 }
                }
                { Get-AuthToken -Identity -Resource 'https://graph.microsoft.com' } |
                    Should -Throw '*Unable to determine the Azure Arc managed identity secret file*'
            }
        }

        It 'Should retrieve the token after obtaining the challenge secret file' {
            InModuleScope 'MSCloudLoginAssistant' {
                $env:AZUREPS_HOST_ENVIRONMENT = ''
                $env:IDENTITY_ENDPOINT = 'http://localhost:40342/metadata/identity/oauth2/token'
                $env:IDENTITY_HEADER = ''
                $env:IMDS_ENDPOINT = 'http://localhost:40342'

                $script:arcCallCount = 0
                Mock -CommandName Invoke-WebRequest -MockWith {
                    $script:arcCallCount++
                    if ($script:arcCallCount -eq 1)
                    {
                        # First request throws with a WWW-Authenticate challenge header
                        # pointing at the secret file.
                        $ex = [System.Exception]::new('401 Unauthorized')
                        $response = [PSCustomObject]@{
                            Headers = @{ 'WWW-Authenticate' = 'Basic realm=C:\secrets\arc-secret' }
                        }
                        $ex | Add-Member -NotePropertyName Response -NotePropertyValue $response -Force
                        throw $ex
                    }
                    return [PSCustomObject]@{
                        Content = '{ "access_token": "arc-token" }'
                    }
                }
                Mock -CommandName Get-Content -MockWith { return 'arc-secret-value' }

                $result = Get-AuthToken -Identity -Resource 'https://graph.microsoft.com'
                $result | Should -Be 'arc-token'
            }
        }
    }

    Context 'When using a client secret, a refresh token or credentials' {
        It 'Should post the <GrantType> grant for <Scenario>' -TestCases @(
            @{ Scenario = 'a client secret on an explicitly supplied token endpoint'; GrantType = 'client_credentials'; Parameters = @{ ClientSecret = 'secret'; Scope = 'scope/.default'; TokenEndpoint = 'https://login.contoso.local/custom/oauth2/v2.0/token' }; ExpectedUri = 'https://login.contoso.local/custom/oauth2/v2.0/token'; ExpectedBody = @{ client_secret = 'secret'; scope = 'scope/.default' } }
            @{ Scenario = 'a client secret with a resource on the v1.0 endpoint'; GrantType = 'client_credentials'; Parameters = @{ ClientSecret = 'secret'; Resource = 'https://admin.microsoft.com' }; ExpectedUri = 'https://login.microsoftonline.com/tenant/oauth2/token'; ExpectedBody = @{ client_secret = 'secret'; resource = 'https://admin.microsoft.com'; scope = $null } }
            @{ Scenario = 'a refresh token on the v2.0 endpoint'; GrantType = 'refresh_token'; Parameters = @{ RefreshToken = 'rt'; Scope = 'scope/.default' }; ExpectedUri = 'https://login.microsoftonline.com/tenant/oauth2/v2.0/token'; ExpectedBody = @{ refresh_token = 'rt'; scope = 'scope/.default' } }
            @{ Scenario = 'a refresh token with a resource'; GrantType = 'refresh_token'; Parameters = @{ RefreshToken = 'rt'; Resource = 'https://admin.microsoft.com' }; ExpectedUri = 'https://login.microsoftonline.com/tenant/oauth2/token'; ExpectedBody = @{ refresh_token = 'rt'; resource = 'https://admin.microsoft.com' } }
            @{ Scenario = 'credentials on the v2.0 endpoint'; GrantType = 'password'; Parameters = @{ Credentials = New-Object PSCredential ('user@contoso.com', (ConvertTo-SecureString 'pwd' -AsPlainText -Force)); Scope = 'scope/.default' }; ExpectedUri = 'https://login.microsoftonline.com/tenant/oauth2/v2.0/token'; ExpectedBody = @{ username = 'user@contoso.com'; password = 'pwd'; scope = 'scope/.default' } }
            @{ Scenario = 'credentials with a resource'; GrantType = 'password'; Parameters = @{ Credentials = New-Object PSCredential ('user@contoso.com', (ConvertTo-SecureString 'pwd' -AsPlainText -Force)); Resource = 'https://admin.microsoft.com' }; ExpectedUri = 'https://login.microsoftonline.com/tenant/oauth2/token'; ExpectedBody = @{ username = 'user@contoso.com'; password = 'pwd'; resource = 'https://admin.microsoft.com' } }
        ) {
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{ GrantType = $GrantType; Parameters = $Parameters; ExpectedUri = $ExpectedUri; ExpectedBody = $ExpectedBody } {
                param ($GrantType, $Parameters, $ExpectedUri, $ExpectedBody)
                Mock -CommandName Invoke-RestMethod -MockWith {
                    return @{ access_token = 'grant-token'; token_type = 'Bearer' }
                }
                $result = Get-AuthToken -AuthorizationUrl 'https://login.microsoftonline.com' `
                    -TenantId 'tenant' -ClientId 'client' @Parameters
                $result.access_token | Should -Be 'grant-token'
                Should -Invoke Invoke-RestMethod -Exactly 1 -ParameterFilter {
                    $Uri -eq $ExpectedUri -and
                    $Body.grant_type -eq $GrantType -and
                    $Body.client_id -eq 'client' -and
                    @($ExpectedBody.Keys | Where-Object { $Body[$_] -ne $ExpectedBody[$_] }).Count -eq 0
                }
            }
        }
    }

    Context 'When using a certificate' {
        It 'Should sign the JWT assertion from a certificate path or thumbprint' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Invoke-RestMethod -MockWith {
                    return @{ access_token = 'cert-token'; token_type = 'Bearer' }
                }
                $pwd = ConvertTo-SecureString 'testpwd' -AsPlainText -Force
                $pathResult = Get-AuthToken -AuthorizationUrl 'https://login.microsoftonline.com' `
                    -TenantId 'tenant' -ClientId 'client' -CertificatePath $script:testPfxPath -CertificatePassword $pwd -Scope 'scope/.default'

                Mock -CommandName Get-MSCloudLoginCertificate -MockWith {
                    return $script:cert
                }
                $thumbprintResult = Get-AuthToken -AuthorizationUrl 'https://login.microsoftonline.com' `
                    -TenantId 'tenant' -ClientId 'client' -CertificateThumbprint 'dummy-thumb' -Scope 'scope/.default'

                $pathResult.access_token | Should -Be 'cert-token'
                $thumbprintResult.access_token | Should -Be 'cert-token'
                Should -Invoke Invoke-RestMethod -Exactly 1 -ParameterFilter {
                    $Body.client_assertion -and $null -eq $Headers
                }
                Should -Invoke Invoke-RestMethod -Exactly 1 -ParameterFilter {
                    $Body.client_assertion -and $Headers.Authorization -eq "Bearer $($Body.client_assertion)"
                }
            }
        }
    }

    Context 'When using the device code flow' {
        It 'Should keep polling while the authorization is still pending' {
            InModuleScope 'MSCloudLoginAssistant' {
                $script:deviceCodeCalls = 0
                Mock -CommandName Write-Verbose -MockWith { }
                Mock -CommandName Invoke-RestMethod -MockWith {
                    $script:deviceCodeCalls++
                    if ($script:deviceCodeCalls -eq 1)
                    {
                        return @{ device_code = 'device-code'; interval = 0; message = 'Sign in please' }
                    }
                    if ($script:deviceCodeCalls -eq 2)
                    {
                        $errorRecord = [System.Management.Automation.ErrorRecord]::new(
                            [System.Exception]::new('pending'), 'AuthorizationPending', 'NotSpecified', $null)
                        $errorRecord.ErrorDetails = [System.Management.Automation.ErrorDetails]::new('{"error":"authorization_pending"}')
                        throw $errorRecord
                    }
                    return @{ access_token = 'device-token'; token_type = 'Bearer' }
                }

                $result = Get-AuthToken -AuthorizationUrl 'https://login.microsoftonline.com' `
                    -TenantId 'contoso.onmicrosoft.com' -ClientId 'client' -DeviceCode `
                    -Resource 'https://admin.microsoft.com'

                $result.access_token | Should -Be 'device-token'
                $script:deviceCodeCalls | Should -Be 3
                Should -Invoke Invoke-RestMethod -ParameterFilter {
                    $Uri -eq 'https://login.microsoftonline.com/contoso.onmicrosoft.com/oauth2/v2.0/devicecode' -and
                    $Body.scope -eq 'https://admin.microsoft.com'
                }
            }
        }

        It 'Should stop polling immediately on <Description>' -TestCases @(
            @{ Description = 'a terminal OAuth error'; Message = 'the user declined the sign-in'; ErrorDetails = '{"error":"access_denied"}' }
            @{ Description = 'a network failure'; Message = 'the remote name could not be resolved'; ErrorDetails = 'the remote name could not be resolved' }
        ) {
            param ($Message, $ErrorDetails)
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{ Message = $Message; ErrorDetails = $ErrorDetails } {
                param ($Message, $ErrorDetails)
                $script:deviceCodeCalls = 0
                Mock -CommandName Write-Verbose -MockWith { }
                Mock -CommandName Invoke-RestMethod -MockWith {
                    $script:deviceCodeCalls++
                    if ($script:deviceCodeCalls -eq 1)
                    {
                        return @{ device_code = 'device-code'; interval = 0; message = 'Sign in please' }
                    }
                    $errorRecord = [System.Management.Automation.ErrorRecord]::new(
                        [System.Exception]::new($Message), 'DeviceCodeFailure', 'NotSpecified', $null)
                    $errorRecord.ErrorDetails = [System.Management.Automation.ErrorDetails]::new($ErrorDetails)
                    throw $errorRecord
                }

                { Get-AuthToken -AuthorizationUrl 'https://login.microsoftonline.com' `
                    -TenantId 'contoso.onmicrosoft.com' -ClientId 'client' -DeviceCode `
                    -Scope 'https://graph.microsoft.com/.default' } | Should -Throw "*$Message*"

                $script:deviceCodeCalls | Should -Be 2
            }
        }
    }
}

# ---------------------------------------------------------------------------
# Connect-MSCloudLoginRESTWorkload
# ---------------------------------------------------------------------------
Describe 'Connect-MSCloudLoginRESTWorkload' {

    BeforeAll {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
        }
    }

    Context 'When a connection already exists' {
        It 'Should reuse a fresh connection without acquiring a new token and renew an expired one' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $profile = $Script:MSCloudLoginConnectionProfile.AdminAPI
                $profile.AuthenticationType = 'ServicePrincipalWithSecret'
                $profile.RequestedAuthenticationType = 'ServicePrincipalWithSecret'
                $profile.ApplicationId = 'app-id'
                $profile.ApplicationSecret = 'secret'
                $profile.TenantId = 'contoso.onmicrosoft.com'
                $profile.AccessToken = 'Bearer existing-token'
                $profile.CompleteConnection()

                Mock -CommandName Get-AuthToken -MockWith { return @{ token_type = 'Bearer'; access_token = 'renewed-token' } }
                Connect-MSCloudLoginRESTWorkload -WorkloadName 'AdminAPI' -AuthorizationUrl 'https://login.microsoftonline.com' -Scope 's' -ClientId 'c'
                Should -Invoke Get-AuthToken -Exactly 0
                $profile.AccessToken | Should -Be 'Bearer existing-token'

                $profile.ConnectedDateTime = [System.DateTime]::Now.AddMinutes(-90).ToString()
                Connect-MSCloudLoginRESTWorkload -WorkloadName 'AdminAPI' -AuthorizationUrl 'https://login.microsoftonline.com' -Scope 's' -ClientId 'c'
                Should -Invoke Get-AuthToken -Exactly 1
                $profile.AccessToken | Should -Be 'Bearer renewed-token'
            }
        }
    }

    Context 'When the authentication method is not supported' {
        It 'Should throw' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $profile = $Script:MSCloudLoginConnectionProfile.AdminAPI
                $profile.AuthenticationType = 'Interactive'
                $profile.RequestedAuthenticationType = 'Interactive'

                { Connect-MSCloudLoginRESTWorkload -WorkloadName 'AdminAPI' -AuthorizationUrl 'u' -Scope 's' -ClientId 'c' } |
                    Should -Throw "*is not supported for workload 'AdminAPI'*"
            }
        }
    }

    Context 'When a token is acquired through Get-AuthToken' {
        It 'Should connect with <AuthenticationType> as client <ExpectedClientId> for tenant <ExpectedTenantId> and store the bearer token' -TestCases @(
            @{
                AuthenticationType = 'Credentials'
                ProfileValues      = @{ ApplicationId = $null; Credentials = New-Object PSCredential ('user@contoso.com', (ConvertTo-SecureString 'pwd' -AsPlainText -Force)) }
                TokenResponse      = @{ token_type = 'Bearer'; access_token = 'cred-token' }
                ExpectedToken      = 'Bearer cred-token'
                ExpectedTenantId   = 'contoso.com'
                ExpectedClientId   = 'c'
            }
            @{
                AuthenticationType = 'CredentialsWithApplicationId'
                ProfileValues      = @{ ApplicationId = 'app-id'; Credentials = New-Object PSCredential ('user@contoso.com', (ConvertTo-SecureString 'pwd' -AsPlainText -Force)) }
                TokenResponse      = @{ token_type = 'Bearer'; access_token = 'cred-app-token' }
                ExpectedToken      = 'Bearer cred-app-token'
                ExpectedTenantId   = 'contoso.com'
                ExpectedClientId   = 'app-id'
            }
            @{
                AuthenticationType = 'ServicePrincipalWithSecret'
                ProfileValues      = @{ ApplicationId = 'app-id'; ApplicationSecret = 'secret'; TenantId = 'tenant' }
                TokenResponse      = @{ token_type = 'Bearer'; access_token = 'sp-secret-token' }
                ExpectedToken      = 'Bearer sp-secret-token'
                ExpectedTenantId   = 'tenant'
                ExpectedClientId   = 'app-id'
            }
            @{
                AuthenticationType = 'ServicePrincipalWithThumbprint'
                ProfileValues      = @{ ApplicationId = 'app-id'; CertificateThumbprint = 'thumb'; TenantId = 'tenant' }
                TokenResponse      = @{ token_type = 'Bearer'; access_token = 'sp-thumb-token' }
                ExpectedToken      = 'Bearer sp-thumb-token'
                ExpectedTenantId   = 'tenant'
                ExpectedClientId   = 'app-id'
            }
            @{
                AuthenticationType = 'ServicePrincipalWithPath'
                ProfileValues      = @{ ApplicationId = 'app-id'; CertificatePath = 'C:\cert.pfx'; CertificatePassword = (ConvertTo-SecureString 'pwd' -AsPlainText -Force); TenantId = 'tenant' }
                TokenResponse      = @{ token_type = 'Bearer'; access_token = 'sp-path-token' }
                ExpectedToken      = 'Bearer sp-path-token'
                ExpectedTenantId   = 'tenant'
                ExpectedClientId   = 'app-id'
            }
        ) {
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{
                AuthenticationType = $AuthenticationType
                ProfileValues      = $ProfileValues
                TokenResponse      = $TokenResponse
                ExpectedToken      = $ExpectedToken
                ExpectedTenantId   = $ExpectedTenantId
                ExpectedClientId   = $ExpectedClientId
            } {
                param ($AuthenticationType, $ProfileValues, $TokenResponse, $ExpectedToken, $ExpectedTenantId, $ExpectedClientId)
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $profile = $Script:MSCloudLoginConnectionProfile.AdminAPI
                $profile.AuthenticationType = $AuthenticationType
                $profile.RequestedAuthenticationType = $AuthenticationType
                foreach ($key in $ProfileValues.Keys)
                {
                    $profile.$key = $ProfileValues[$key]
                }

                $script:tokenResponse = $TokenResponse
                Mock -CommandName Get-AuthToken -MockWith { return $script:tokenResponse }
                Connect-MSCloudLoginRESTWorkload -WorkloadName 'AdminAPI' -AuthorizationUrl 'u' -Scope 's' -ClientId 'c'

                $profile.Connected | Should -BeTrue
                $profile.AccessToken | Should -Be $ExpectedToken
                $profile.MultiFactorAuthentication | Should -BeFalse
                Should -Invoke Get-AuthToken -Exactly 1 -ParameterFilter { $TenantId -eq $ExpectedTenantId -and $ClientId -eq $ExpectedClientId }
            }
        }

        It 'Should connect using a managed identity token' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $profile = $Script:MSCloudLoginConnectionProfile.AdminAPI
                $profile.AuthenticationType = 'Identity'
                $profile.RequestedAuthenticationType = 'Identity'
                $profile.TenantId = 'tenant'

                Mock -CommandName Get-AuthToken -MockWith { return 'identity-token-raw' }
                Connect-MSCloudLoginRESTWorkload -WorkloadName 'AdminAPI' -AuthorizationUrl 'u' -Scope 'https://graph.microsoft.com/.default' -ClientId 'c'
                $profile.Connected | Should -BeTrue
                $profile.AccessToken | Should -Be 'Bearer identity-token-raw'
            }
        }
    }

    Context 'When credentials require MFA' {
        It 'Should retry with the device code flow and mark the connection as MFA' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $profile = $Script:MSCloudLoginConnectionProfile.AdminAPI
                $profile.AuthenticationType = 'Credentials'
                $profile.RequestedAuthenticationType = 'Credentials'
                $profile.Credentials = New-Object PSCredential ('user@contoso.com', (ConvertTo-SecureString 'pwd' -AsPlainText -Force))

                $script:authCallCount = 0
                Mock -CommandName Get-AuthToken -MockWith {
                    $script:authCallCount++
                    if ($script:authCallCount -eq 1)
                    {
                        throw 'AADSTS50076: Due to a configuration change made by your administrator you must use multi-factor authentication to access this resource.'
                    }
                    return @{ token_type = 'Bearer'; access_token = 'mfa-token' }
                }
                Connect-MSCloudLoginRESTWorkload -WorkloadName 'AdminAPI' -AuthorizationUrl 'u' -Scope 's' -ClientId 'c'
                $profile.Connected | Should -BeTrue
                $profile.MultiFactorAuthentication | Should -BeTrue
                $profile.AccessToken | Should -Be 'Bearer mfa-token'
                $script:authCallCount | Should -Be 2
            }
        }
    }

    Context 'When the token request fails with a non-MFA error' {
        It 'Should rethrow and mark the workload as disconnected' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $profile = $Script:MSCloudLoginConnectionProfile.AdminAPI
                $profile.AuthenticationType = 'Credentials'
                $profile.RequestedAuthenticationType = 'Credentials'
                $profile.Credentials = New-Object PSCredential ('user@contoso.com', (ConvertTo-SecureString 'pwd' -AsPlainText -Force))

                Mock -CommandName Get-AuthToken -MockWith { throw 'access denied' }
                { Connect-MSCloudLoginRESTWorkload -WorkloadName 'AdminAPI' -AuthorizationUrl 'u' -Scope 's' -ClientId 'c' } |
                    Should -Throw 'access denied'
                $profile.Connected | Should -BeFalse
            }
        }
    }

    Context 'When using AccessTokens' {
        It 'Should store <ProvidedToken> as <ExpectedToken>' -TestCases @(
            @{ ProvidedToken = 'raw-token'; ExpectedToken = 'Bearer raw-token' }
            @{ ProvidedToken = 'Bearer prefixed-token'; ExpectedToken = 'Bearer prefixed-token' }
        ) {
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{ ProvidedToken = $ProvidedToken; ExpectedToken = $ExpectedToken } {
                param ($ProvidedToken, $ExpectedToken)
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $profile = $Script:MSCloudLoginConnectionProfile.AdminAPI
                $profile.AuthenticationType = 'AccessTokens'
                $profile.RequestedAuthenticationType = 'AccessTokens'
                $profile.AccessTokens = @($ProvidedToken)

                Connect-MSCloudLoginRESTWorkload -WorkloadName 'AdminAPI' -AuthorizationUrl 'u' -Scope 's' -ClientId 'c'
                $profile.Connected | Should -BeTrue
                $profile.AccessToken | Should -Be $ExpectedToken
            }
        }
    }
}

# ---------------------------------------------------------------------------
# Disconnect-MSCloudLoginRESTWorkload
# ---------------------------------------------------------------------------
Describe 'Disconnect-MSCloudLoginRESTWorkload' {

    BeforeAll {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
        }
    }

    It 'Should clear the connection state and token and not throw when already disconnected' {
        InModuleScope 'MSCloudLoginAssistant' {
            $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
            $profile = $Script:MSCloudLoginConnectionProfile.AdminAPI
            $profile.Connected = $true
            $profile.AccessToken = 'Bearer some-token'

            Disconnect-MSCloudLoginRESTWorkload -WorkloadName 'AdminAPI'
            $profile.Connected | Should -BeFalse
            $profile.AccessToken | Should -BeNullOrEmpty

            { Disconnect-MSCloudLoginRESTWorkload -WorkloadName 'AdminAPI' } | Should -Not -Throw
        }
    }
}
