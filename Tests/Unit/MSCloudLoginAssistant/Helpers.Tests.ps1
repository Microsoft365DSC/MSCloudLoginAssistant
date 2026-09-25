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

    $script:moduleRoot = Resolve-Path (Join-Path $PSScriptRoot '..\..\..\Modules\MSCloudLoginAssistant')
    Import-Module (Join-Path $script:moduleRoot 'MSCloudLoginAssistant.psd1') -Force
}

AfterAll {
    if ($script:tempModuleBase -and (Test-Path $script:tempModuleBase))
    {
        Remove-Item -Path $script:tempModuleBase -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Describe 'Get-MSCloudLoginCertificate' {

    Context 'When a thumbprint is provided' {
        It 'Should return the certificate from the current user store' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Find-MSCloudLoginStoreCertificate -MockWith {
                    if ($StoreLocation -eq 'CurrentUser')
                    {
                        return [PSCustomObject]@{ Thumbprint = 'AA11'; Store = 'CurrentUser' }
                    }
                    return $null
                }

                $certificate = Get-MSCloudLoginCertificate -CertificateThumbprint 'AA11'
                $certificate.Store | Should -Be 'CurrentUser'
                Should -Invoke Find-MSCloudLoginStoreCertificate -Exactly 1
            }
        }

        It 'Should fall back to the local machine store' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Find-MSCloudLoginStoreCertificate -MockWith {
                    if ($StoreLocation -eq 'LocalMachine')
                    {
                        return [PSCustomObject]@{ Thumbprint = 'AA11'; Store = 'LocalMachine' }
                    }
                    return $null
                }

                $certificate = Get-MSCloudLoginCertificate -CertificateThumbprint 'AA11'
                $certificate.Store | Should -Be 'LocalMachine'
                Should -Invoke Find-MSCloudLoginStoreCertificate -Exactly 2
            }
        }

        It 'Should return a store certificate without Cert: drive properties' {
            InModuleScope 'MSCloudLoginAssistant' {
                $thumbprint = (Get-ChildItem -Path 'Cert:\CurrentUser\My' | Select-Object -First 1).Thumbprint
                if ($null -eq $thumbprint)
                {
                    Set-ItResult -Skipped -Because 'the CurrentUser\My store is empty'
                    return
                }

                $certificate = Find-MSCloudLoginStoreCertificate -StoreLocation 'CurrentUser' -CertificateThumbprint $thumbprint
                $certificate.Thumbprint | Should -Be $thumbprint
                $certificate.PSObject.Properties['PSDrive'] | Should -BeNullOrEmpty
            }
        }

        It 'Should return $null for an unknown thumbprint' {
            InModuleScope 'MSCloudLoginAssistant' {
                Find-MSCloudLoginStoreCertificate -StoreLocation 'CurrentUser' -CertificateThumbprint '0000000000000000000000000000000000000000' | Should -BeNullOrEmpty
            }
        }

        It 'Should name both stores when the certificate does not exist' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Find-MSCloudLoginStoreCertificate -MockWith { return $null }

                { Get-MSCloudLoginCertificate -CertificateThumbprint 'AA11' } |
                    Should -Throw "*'AA11' was not found in the CurrentUser\My nor the LocalMachine\My certificate store*"
            }
        }
    }

    Context 'When a certificate path is provided' {
        BeforeAll {
            $script:unprotectedPfxPath = Join-Path $env:TEMP ('msla-helper-{0}.pfx' -f ([guid]::NewGuid().ToString('N')))
            $rsa = [System.Security.Cryptography.RSA]::Create(2048)
            $request = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new(
                [System.Security.Cryptography.X509Certificates.X500DistinguishedName]::new('CN=MSCloudLoginAssistantHelperTest'),
                $rsa,
                [System.Security.Cryptography.HashAlgorithmName]::SHA256,
                [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
            $certificate = $request.CreateSelfSigned([System.DateTimeOffset]::Now.AddDays(-1), [System.DateTimeOffset]::Now.AddDays(1))
            [System.IO.File]::WriteAllBytes($script:unprotectedPfxPath, $certificate.Export(
                [System.Security.Cryptography.X509Certificates.X509ContentType]::Pfx))
            $rsa.Dispose()
        }

        AfterAll {
            Remove-Item -Path $script:unprotectedPfxPath -Force -ErrorAction SilentlyContinue
        }

        It 'Should load a PFX file that has no password' {
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{ PfxPath = $script:unprotectedPfxPath } {
                param ($PfxPath)
                $certificate = Get-MSCloudLoginCertificate -CertificatePath $PfxPath
                $certificate.Subject | Should -Be 'CN=MSCloudLoginAssistantHelperTest'
            }
        }

        It 'Should throw when the file does not exist' {
            InModuleScope 'MSCloudLoginAssistant' {
                { Get-MSCloudLoginCertificate -CertificatePath 'C:\does\not\exist.pfx' } |
                    Should -Throw "*'C:\does\not\exist.pfx' was not found*"
            }
        }
    }
}

Describe 'Remove-MSCloudLoginProxyModule' {

    It 'Should remove every loaded module that exports the probe command' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
            Mock -CommandName Remove-Module -MockWith { }
            Mock -CommandName Get-Module -MockWith {
                $proxyCommands = [System.Collections.Generic.Dictionary[string, object]]::new()
                $proxyCommands.Add('Get-AcceptedDomain', $null)
                $otherCommands = [System.Collections.Generic.Dictionary[string, object]]::new()
                $otherCommands.Add('Get-Something', $null)
                return @(
                    [PSCustomObject]@{ Name = 'tmpEXO_abc'; ExportedCommands = $proxyCommands }
                    [PSCustomObject]@{ Name = 'SomethingElse'; ExportedCommands = $otherCommands }
                )
            }

            Remove-MSCloudLoginProxyModule -ProbeCommand 'Get-AcceptedDomain' -Source 'Test'

            Should -Invoke Remove-Module -Exactly 1
            Should -Invoke Add-MSCloudLoginAssistantEvent -ParameterFilter { $Message -like '*tmpEXO_abc*' }
        }
    }

    It 'Should do nothing when no module exports the probe command' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
            Mock -CommandName Remove-Module -MockWith { }
            Mock -CommandName Get-Module -MockWith {
                $otherCommands = [System.Collections.Generic.Dictionary[string, object]]::new()
                $otherCommands.Add('Get-Something', $null)
                return @([PSCustomObject]@{ Name = 'SomethingElse'; ExportedCommands = $otherCommands })
            }

            Remove-MSCloudLoginProxyModule -ProbeCommand 'Get-AcceptedDomain' -Source 'Test'

            Should -Invoke Remove-Module -Exactly 0
        }
    }
}

Describe 'Restore-MSCloudLoginProxyModule' {

    Context 'When two loaded proxy modules export the same command' {
        BeforeEach {
            # Both proxy modules export Get-MSCLARestoreShared.
            $proxyDefinitions = @{
                tmpEXO_restoreexo = @'
$global:MSCLARestoreExoLoads = 1 + [int]$global:MSCLARestoreExoLoads
function Get-MSCLARestoreShared { 'ExchangeOnline' }
function Get-MSCLARestoreExoProbe { }
Export-ModuleMember -Function Get-MSCLARestoreShared, Get-MSCLARestoreExoProbe
'@
                tmpEXO_restoresc  = @'
$global:MSCLARestoreScLoads = 1 + [int]$global:MSCLARestoreScLoads
function Get-MSCLARestoreShared { 'SecurityCompliance' }
function Get-MSCLARestoreScProbe { }
Export-ModuleMember -Function Get-MSCLARestoreShared, Get-MSCLARestoreScProbe
'@
            }

            $global:MSCLARestoreExoLoads = 0
            $global:MSCLARestoreScLoads = 0
            foreach ($moduleName in @('tmpEXO_restoreexo', 'tmpEXO_restoresc'))
            {
                $modulePath = Join-Path -Path $TestDrive -ChildPath "$moduleName.psm1"
                Set-Content -Path $modulePath -Value $proxyDefinitions[$moduleName]
                Import-Module -Name $modulePath -Global -DisableNameChecking
            }
        }

        AfterEach {
            Remove-Module -Name 'tmpEXO_restoreexo', 'tmpEXO_restoresc' -Force -ErrorAction SilentlyContinue
            Remove-Variable -Name 'MSCLARestoreExoLoads', 'MSCLARestoreScLoads' -Scope Global -ErrorAction SilentlyContinue
        }

        It 'Should give the module that exports the probe command precedence again' {
            Get-MSCLARestoreShared | Should -Be 'SecurityCompliance'

            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                Restore-MSCloudLoginProxyModule -ProbeCommand 'Get-MSCLARestoreExoProbe' -Source 'Test' | Should -BeTrue
            }

            Get-MSCLARestoreShared | Should -Be 'ExchangeOnline'
            (Get-Command -Name 'Get-MSCLARestoreShared').Module.Name | Should -Be 'tmpEXO_restoreexo'
        }

        It 'Should switch the precedence back and forth without reloading either module' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                foreach ($iteration in 1..3)
                {
                    Restore-MSCloudLoginProxyModule -ProbeCommand 'Get-MSCLARestoreExoProbe' -Source 'Test' | Should -BeTrue
                    Get-MSCLARestoreShared | Should -Be 'ExchangeOnline'

                    Restore-MSCloudLoginProxyModule -ProbeCommand 'Get-MSCLARestoreScProbe' -Source 'Test' | Should -BeTrue
                    Get-MSCLARestoreShared | Should -Be 'SecurityCompliance'
                }
            }

            $global:MSCLARestoreExoLoads | Should -Be 1
            $global:MSCLARestoreScLoads | Should -Be 1
            @(Get-Module -Name 'tmpEXO_restore*').Count | Should -Be 2
        }
    }

    It 'Should report a missing proxy module without importing anything' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
            Mock -CommandName Import-Module -MockWith { }
            Mock -CommandName Get-Module -MockWith {
                $otherCommands = [System.Collections.Generic.Dictionary[string, object]]::new()
                $otherCommands.Add('Get-Something', $null)
                return @([PSCustomObject]@{ Name = 'SomethingElse'; ExportedCommands = $otherCommands })
            }

            Restore-MSCloudLoginProxyModule -ProbeCommand 'Get-AcceptedDomain' -Source 'Test' | Should -BeFalse

            Should -Invoke Import-Module -Exactly 0
        }
    }

    It 'Should report a failed import' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
            Mock -CommandName Import-Module -MockWith { throw 'import failed' }
            Mock -CommandName Get-Module -MockWith {
                return New-Module -Name 'tmpEXO_abc' -ScriptBlock { function Get-AcceptedDomain { } }
            }

            Restore-MSCloudLoginProxyModule -ProbeCommand 'Get-AcceptedDomain' -Source 'Test' | Should -BeFalse

            Should -Invoke Add-MSCloudLoginAssistantEvent -ParameterFilter { $Message -like '*Failed to restore proxy module {tmpEXO_abc}*import failed*' }
        }
    }
}

Describe 'Disconnect-MSCloudLoginExchangeConnection' {

    BeforeAll {
        # Stubs prevent the autoload of ExchangeOnlineManagement, whose assemblies block Microsoft.Graph.Authentication in Windows PowerShell.
        function global:Get-ConnectionInformation { }
        function global:Disconnect-ExchangeOnline { param ([System.String[]] $ConnectionId, [switch] $Confirm) }
    }

    AfterAll {
        Remove-Item -Path 'Function:\Get-ConnectionInformation', 'Function:\Disconnect-ExchangeOnline' -ErrorAction SilentlyContinue
    }

    BeforeEach {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
            Mock -CommandName Disconnect-ExchangeOnline -MockWith { }
            Mock -CommandName Get-ConnectionInformation -MockWith {
                return @(
                    [PSCustomObject]@{ ConnectionId = [guid]'11111111-1111-1111-1111-111111111111'; IsEopSession = $false }
                    [PSCustomObject]@{ ConnectionId = [guid]'22222222-2222-2222-2222-222222222222'; IsEopSession = $true }
                    [PSCustomObject]@{ ConnectionId = [guid]'33333333-3333-3333-3333-333333333333'; IsEopSession = $false }
                )
            }
        }
    }

    It 'Should disconnect only the Exchange Online connections' {
        InModuleScope 'MSCloudLoginAssistant' {
            Disconnect-MSCloudLoginExchangeConnection -Source 'Test'

            Should -Invoke Disconnect-ExchangeOnline -Exactly 1 -ParameterFilter {
                ($ConnectionId -join ',') -eq '11111111-1111-1111-1111-111111111111,33333333-3333-3333-3333-333333333333'
            }
        }
    }

    It 'Should disconnect only the Security & Compliance connections' {
        InModuleScope 'MSCloudLoginAssistant' {
            Disconnect-MSCloudLoginExchangeConnection -SecurityCompliance -Source 'Test'

            Should -Invoke Disconnect-ExchangeOnline -Exactly 1 -ParameterFilter {
                ($ConnectionId -join ',') -eq '22222222-2222-2222-2222-222222222222'
            }
        }
    }

    It 'Should not call Disconnect-ExchangeOnline without a matching connection' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Get-ConnectionInformation -MockWith { return $null }

            Disconnect-MSCloudLoginExchangeConnection -Source 'Test'

            Should -Invoke Disconnect-ExchangeOnline -Exactly 0
        }
    }
}

Describe 'Get-MSCloudLoginAccessTokenExpiry' {

    It 'Should read the exp claim of a JSON Web Token with or without the Bearer prefix' {
        InModuleScope 'MSCloudLoginAssistant' {
            $expiresOn = [System.DateTimeOffset]::FromUnixTimeSeconds(1790270000).LocalDateTime
            $payload = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes('{"exp":1790270000,"aud":"https://ps.compliance.protection.outlook.com"}')).TrimEnd('=').Replace('+', '-').Replace('/', '_')
            $token = "eyJhbGciOiJub25lIn0.$payload.signature"

            Get-MSCloudLoginAccessTokenExpiry -Token $token | Should -Be $expiresOn
            Get-MSCloudLoginAccessTokenExpiry -Token "Bearer $token" | Should -Be $expiresOn
        }
    }

    It 'Should return $null for <Description>' -TestCases @(
        @{ Description = 'an opaque token'; Token = 'opaque-token' }
        @{ Description = 'an empty token'; Token = '' }
        @{ Description = 'a token without exp claim'; Token = 'eyJhbGciOiJub25lIn0.eyJhdWQiOiJ4In0.signature' }
        @{ Description = 'a token with an invalid payload'; Token = 'a.%%%.c' }
    ) {
        param ($Token)
        InModuleScope 'MSCloudLoginAssistant' -Parameters @{ Token = $Token } {
            param ($Token)

            Get-MSCloudLoginAccessTokenExpiry -Token $Token | Should -BeNullOrEmpty
        }
    }
}

Describe 'Get-MSCloudLoginEndpointInfo' {

    It 'Should throw when neither the environment nor a default entry is defined' {
        InModuleScope 'MSCloudLoginAssistant' {
            $originalEndpointData = $Script:WorkloadEndpointData
            try
            {
                $Script:WorkloadEndpointData = @{
                    TestWorkload = @{ AzureCloud = @{ HostUrl = 'https://contoso.local' } }
                }
                { Get-MSCloudLoginEndpointInfo -Workload 'TestWorkload' -EnvironmentName 'AzureDOD' } |
                    Should -Throw "*'TestWorkload' in environment 'AzureDOD' and the workload has no default entry*"
            }
            finally
            {
                $Script:WorkloadEndpointData = $originalEndpointData
            }
        }
    }

    It 'Should leave non-string endpoint values untouched' {
        InModuleScope 'MSCloudLoginAssistant' {
            $originalEndpointData = $Script:WorkloadEndpointData
            try
            {
                $Script:WorkloadEndpointData = @{
                    TestWorkload = @{ default = @{ Endpoints = @{ Graph = 'https://graph.contoso.local' }; HostUrl = 'https://{Resource}.contoso.local' } }
                }
                $result = Get-MSCloudLoginEndpointInfo -Workload 'TestWorkload' -EnvironmentName 'AzureCloud' -Replacements @{ Resource = 'api' }
                $result.HostUrl | Should -Be 'https://api.contoso.local'
                $result.Endpoints.Graph | Should -Be 'https://graph.contoso.local'
            }
            finally
            {
                $Script:WorkloadEndpointData = $originalEndpointData
            }
        }
    }
}

Describe 'Test-MSCloudLoginConnectionReusable' {

    Context 'When the token expiry is known' {
        BeforeEach {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
            }
        }

        It 'Should renew a <AuthenticationType> connection whose token expires within five minutes' -TestCases @(
            @{ AuthenticationType = 'Identity' }
            @{ AuthenticationType = 'ServicePrincipalWithThumbprint' }
        ) {
            param ($AuthenticationType)
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{ AuthenticationType = $AuthenticationType } {
                param ($AuthenticationType)

                $workloadProfile = [PSCustomObject]@{
                    Connected          = $true
                    ConnectedDateTime  = [System.DateTime]::Now.ToString()
                    AuthenticationType = $AuthenticationType
                    TokenExpiresOn     = [System.DateTime]::Now.AddMinutes(4)
                }

                Test-MSCloudLoginConnectionReusable -WorkloadProfile $workloadProfile -TokenBasedAuthTypes @() -Source 'Test' | Should -BeFalse
                $workloadProfile.Connected | Should -BeFalse
            }
        }

        It 'Should reuse a connection whose token is valid beyond the renewal window even after 50 minutes' {
            InModuleScope 'MSCloudLoginAssistant' {
                $workloadProfile = [PSCustomObject]@{
                    Connected          = $true
                    ConnectedDateTime  = [System.DateTime]::Now.AddMinutes(-70).ToString()
                    AuthenticationType = 'Identity'
                    TokenExpiresOn     = [System.DateTime]::Now.AddMinutes(20)
                }

                Test-MSCloudLoginConnectionReusable -WorkloadProfile $workloadProfile -Source 'Test' | Should -BeTrue
            }
        }

        It 'Should keep a supplied access token until it expires' {
            InModuleScope 'MSCloudLoginAssistant' {
                $workloadProfile = [PSCustomObject]@{
                    Connected          = $true
                    ConnectedDateTime  = [System.DateTime]::Now.ToString()
                    AuthenticationType = 'AccessTokens'
                    TokenExpiresOn     = [System.DateTime]::Now.AddMinutes(2)
                }

                Test-MSCloudLoginConnectionReusable -WorkloadProfile $workloadProfile -Source 'Test' | Should -BeTrue

                $workloadProfile.TokenExpiresOn = [System.DateTime]::Now.AddSeconds(-1)
                Test-MSCloudLoginConnectionReusable -WorkloadProfile $workloadProfile -Source 'Test' | Should -BeFalse
            }
        }
    }

    It 'Should treat a failing probe as a lost connection' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

            $workloadProfile = New-Object AdminAPI
            $workloadProfile.AuthenticationType = 'ServicePrincipalWithThumbprint'
            $workloadProfile.CompleteConnection()

            $result = Test-MSCloudLoginConnectionReusable -WorkloadProfile $workloadProfile `
                -ProbeScript { throw 'the SDK context is gone' } -Source 'Test'

            $result | Should -BeFalse
            $workloadProfile.Connected | Should -BeFalse
            Should -Invoke Add-MSCloudLoginAssistantEvent -ParameterFilter { $Message -like 'Connection probe failed*' }
        }
    }

    It 'Should reuse the connection when the probe returns a context' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

            $workloadProfile = New-Object AdminAPI
            $workloadProfile.AuthenticationType = 'ServicePrincipalWithThumbprint'
            $workloadProfile.CompleteConnection()

            (Test-MSCloudLoginConnectionReusable -WorkloadProfile $workloadProfile `
                -ProbeScript { return @{ TenantId = 'contoso' } } -Source 'Test') | Should -BeTrue
        }
    }
}

Describe 'Microsoft Graph connection probe' {

    AfterAll {
        [System.AppDomain]::CurrentDomain.SetData('MSCloudLoginAssistant.ConnectionIdentity.MicrosoftGraph', $null)
    }

    It 'Should reject a Graph context of another application' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Get-MgContext -MockWith { [PSCustomObject]@{ ClientId = 'other-app'; Account = $null } }
            $workloadProfile = [PSCustomObject]@{ AuthenticationType = 'ServicePrincipalWithThumbprint'; ApplicationId = 'expected-app'; Credentials = $null }
            Set-MSCloudLoginProcessConnectionIdentity -Workload 'MicrosoftGraph' -Identity (Get-MSCloudLoginConnectionIdentity -WorkloadProfile $workloadProfile)

            & $Script:MSCloudLoginConnectionProbes.MicrosoftGraph $workloadProfile | Should -BeNullOrEmpty
        }
    }

    It 'Should reject a Graph context of another account' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Get-MgContext -MockWith { [PSCustomObject]@{ ClientId = 'app'; Account = 'other@contoso.com' } }
            $credential = [System.Management.Automation.PSCredential]::new('admin@contoso.com', (ConvertTo-SecureString -String 'x' -AsPlainText -Force))
            $workloadProfile = [PSCustomObject]@{ AuthenticationType = 'Credentials'; ApplicationId = 'app'; Credentials = $credential }
            Set-MSCloudLoginProcessConnectionIdentity -Workload 'MicrosoftGraph' -Identity (Get-MSCloudLoginConnectionIdentity -WorkloadProfile $workloadProfile)

            & $Script:MSCloudLoginConnectionProbes.MicrosoftGraph $workloadProfile | Should -BeNullOrEmpty
        }
    }

    It 'Should reject a managed identity profile when another identity connected the process' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Get-MgContext -MockWith { [PSCustomObject]@{ ClientId = 'other-app'; Account = $null } }
            $workloadProfile = [PSCustomObject]@{ AuthenticationType = 'Identity'; TenantId = 'contoso.onmicrosoft.com'; ApplicationId = $null; Credentials = $null }
            Set-MSCloudLoginProcessConnectionIdentity -Workload 'MicrosoftGraph' -Identity 'ServicePrincipalWithThumbprint|contoso.onmicrosoft.com|other-app|'

            & $Script:MSCloudLoginConnectionProbes.MicrosoftGraph $workloadProfile | Should -BeNullOrEmpty
        }
    }

    It 'Should accept the Graph context of the profile' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Get-MgContext -MockWith { [PSCustomObject]@{ ClientId = 'expected-app'; Account = $null } }
            $workloadProfile = [PSCustomObject]@{ AuthenticationType = 'ServicePrincipalWithThumbprint'; ApplicationId = 'expected-app'; Credentials = $null }
            Set-MSCloudLoginProcessConnectionIdentity -Workload 'MicrosoftGraph' -Identity (Get-MSCloudLoginConnectionIdentity -WorkloadProfile $workloadProfile)

            & $Script:MSCloudLoginConnectionProbes.MicrosoftGraph $workloadProfile | Should -Not -BeNullOrEmpty
        }
    }
}

Describe 'Get-MSCloudLoginConnectionIdentity' {

    It 'Should distinguish access tokens of different principals' {
        InModuleScope 'MSCloudLoginAssistant' {
            $newToken = {
                param ($Claims)
                $encode = { param ($Text) [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($Text)).TrimEnd('=').Replace('+', '-').Replace('/', '_') }
                '{0}.{1}.signature' -f (& $encode '{"alg":"none"}'), (& $encode $Claims)
            }
            $first = [PSCustomObject]@{ AuthenticationType = 'AccessTokens'; TenantId = 'contoso.onmicrosoft.com'; AccessTokens = @(& $newToken '{"appid":"app-1","oid":"object-1"}') }
            $second = [PSCustomObject]@{ AuthenticationType = 'AccessTokens'; TenantId = 'contoso.onmicrosoft.com'; AccessTokens = @(& $newToken '{"appid":"app-2","oid":"object-2"}') }
            $renewed = [PSCustomObject]@{ AuthenticationType = 'AccessTokens'; TenantId = 'contoso.onmicrosoft.com'; AccessTokens = @(& $newToken '{"appid":"app-1","oid":"object-1","exp":1}') }

            Get-MSCloudLoginConnectionIdentity -WorkloadProfile $first | Should -Not -Be (Get-MSCloudLoginConnectionIdentity -WorkloadProfile $second)
            Get-MSCloudLoginConnectionIdentity -WorkloadProfile $first | Should -Be (Get-MSCloudLoginConnectionIdentity -WorkloadProfile $renewed)
        }
    }
}

Describe 'Get-MSCloudLoginTenantGuid' {

    BeforeEach {
        InModuleScope 'MSCloudLoginAssistant' {
            $Script:MSCloudLoginTenantGuidCache = @{}
        }
    }

    It 'Should return a TenantId that already is a GUID without a discovery request' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Invoke-WebRequest -MockWith { throw 'Unexpected discovery request' }

            Get-MSCloudLoginTenantGuid -TenantId '22222222-2222-2222-2222-222222222222' | Should -Be '22222222-2222-2222-2222-222222222222'
            Should -Invoke Invoke-WebRequest -Exactly 0
        }
    }

    It 'Should resolve the GUID from the token endpoint once and reuse the cached value' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
            Mock -CommandName Invoke-WebRequest -MockWith {
                return @{ Content = '{ "token_endpoint": "https://login.microsoftonline.com/22222222-2222-2222-2222-222222222222/oauth2/v2.0/token" }' }
            }

            Get-MSCloudLoginTenantGuid -TenantId 'contoso.onmicrosoft.com' | Should -Be '22222222-2222-2222-2222-222222222222'
            Get-MSCloudLoginTenantGuid -TenantId 'contoso.onmicrosoft.com' | Should -Be '22222222-2222-2222-2222-222222222222'
            Should -Invoke Invoke-WebRequest -Exactly 1
        }
    }

    It 'Should reuse the GUID captured by the environment detection' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
            Mock -CommandName Invoke-WebRequest -MockWith {
                return @{ Content = '{ "token_endpoint": "https://login.microsoftonline.com/22222222-2222-2222-2222-222222222222/oauth2/v2.0/token" }' }
            }

            $null = Get-CloudEnvironmentInfo -TenantId 'contoso.onmicrosoft.com'
            Get-MSCloudLoginTenantGuid -TenantId 'contoso.onmicrosoft.com' | Should -Be '22222222-2222-2222-2222-222222222222'
            Should -Invoke Invoke-WebRequest -Exactly 1
        }
    }

    It 'Should return $null when the token endpoint contains no GUID' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
            Mock -CommandName Invoke-WebRequest -MockWith {
                return @{ Content = '{ "token_endpoint": "https://login.microsoftonline.com/t/oauth2/v2.0/token" }' }
            }

            Get-MSCloudLoginTenantGuid -TenantId 'contoso.onmicrosoft.com' | Should -BeNullOrEmpty
        }
    }
}

Describe 'Compare-InputParametersForChange with stored access tokens' {

    It 'Should ignore the token a <AuthenticationType> connection acquired itself' -TestCases @(
        @{ AuthenticationType = 'Credentials'; Workload = 'MicrosoftGraph' }
        @{ AuthenticationType = 'ServicePrincipalWithThumbprint'; Workload = 'MicrosoftTeams' }
    ) {
        param ($AuthenticationType, $Workload)
        InModuleScope 'MSCloudLoginAssistant' -Parameters @{ AuthenticationType = $AuthenticationType; Workload = $Workload } {
            param ($AuthenticationType, $Workload)
            $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
            $profileName = if ($Workload -eq 'MicrosoftTeams') { 'Teams' } else { $Workload }
            $workloadProfile = $Script:MSCloudLoginConnectionProfile.$profileName
            $credential = [System.Management.Automation.PSCredential]::new('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString -String 'x' -AsPlainText -Force))
            $parameters = @{ Workload = $Workload }
            if ($AuthenticationType -eq 'Credentials')
            {
                $workloadProfile.Credentials = $credential
                $workloadProfile.ApplicationId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'
                $workloadProfile.TenantId = 'contoso.onmicrosoft.com'
                $parameters.Credential = $credential
            }
            else
            {
                $workloadProfile.ApplicationId = 'app'
                $workloadProfile.CertificateThumbprint = 'ABC'
                $workloadProfile.TenantId = 'contoso.onmicrosoft.com'
                $parameters += @{ ApplicationId = 'app'; CertificateThumbprint = 'ABC'; TenantId = 'contoso.onmicrosoft.com' }
            }
            $workloadProfile.AuthenticationType = $AuthenticationType
            $workloadProfile.RequestedAuthenticationType = $AuthenticationType
            $workloadProfile.AccessTokens = @('acquired-token')

            Compare-InputParametersForChange -CurrentParamSet $parameters | Should -BeFalse
        }
    }

    It 'Should detect a different access token passed by the caller' {
        InModuleScope 'MSCloudLoginAssistant' {
            $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
            $workloadProfile = $Script:MSCloudLoginConnectionProfile.MicrosoftGraph
            $workloadProfile.AuthenticationType = 'AccessTokens'
            $workloadProfile.RequestedAuthenticationType = 'AccessTokens'
            $workloadProfile.TenantId = 'contoso.onmicrosoft.com'
            $workloadProfile.AccessTokens = @('first-token')

            Compare-InputParametersForChange -CurrentParamSet @{ Workload = 'MicrosoftGraph'; TenantId = 'contoso.onmicrosoft.com'; AccessTokens = @('second-token') } | Should -BeTrue
        }
    }
}

Describe 'Azure connection probe' {

    BeforeAll {
        function global:Get-AzContext { }
    }

    AfterAll {
        Remove-Item -Path 'Function:\Get-AzContext' -ErrorAction SilentlyContinue
    }

    It 'Should reject an Azure context of <Description>' -TestCases @(
        @{ Description = 'another application'; AuthenticationType = 'ServicePrincipalWithThumbprint'; AccountId = 'other-app'; AccountType = 'ServicePrincipal'; SubscriptionId = $null }
        @{ Description = 'another account'; AuthenticationType = 'Credentials'; AccountId = 'other@contoso.com'; AccountType = 'User'; SubscriptionId = $null }
        @{ Description = 'a user instead of the managed identity'; AuthenticationType = 'Identity'; AccountId = 'admin@contoso.com'; AccountType = 'User'; SubscriptionId = $null }
        @{ Description = 'another subscription'; AuthenticationType = 'ServicePrincipalWithThumbprint'; AccountId = 'expected-app'; AccountType = 'ServicePrincipal'; SubscriptionId = 'other-subscription' }
    ) {
        param ($AuthenticationType, $AccountId, $AccountType, $SubscriptionId)
        InModuleScope 'MSCloudLoginAssistant' -Parameters @{ AuthenticationType = $AuthenticationType; AccountId = $AccountId; AccountType = $AccountType; SubscriptionId = $SubscriptionId } {
            param ($AuthenticationType, $AccountId, $AccountType, $SubscriptionId)
            $context = [PSCustomObject]@{
                Account      = [PSCustomObject]@{ Id = $AccountId; Type = $AccountType }
                Subscription = [PSCustomObject]@{ Id = $SubscriptionId }
            }
            Mock -CommandName Get-AzContext -MockWith { $context }
            $credential = [System.Management.Automation.PSCredential]::new('admin@contoso.com', (ConvertTo-SecureString -String 'x' -AsPlainText -Force))
            $workloadProfile = [PSCustomObject]@{
                AuthenticationType = $AuthenticationType
                ApplicationId      = 'expected-app'
                Credentials        = $credential
                SubscriptionId     = if ($null -ne $SubscriptionId) { 'expected-subscription' } else { $null }
            }

            & $Script:MSCloudLoginConnectionProbes.Azure $workloadProfile | Should -BeNullOrEmpty
        }
    }

    It 'Should accept the Azure context of <AuthenticationType>' -TestCases @(
        @{ AuthenticationType = 'ServicePrincipalWithSecret'; AccountId = 'expected-app'; AccountType = 'ServicePrincipal' }
        @{ AuthenticationType = 'CredentialsWithTenantId'; AccountId = 'Admin@contoso.com'; AccountType = 'User' }
        @{ AuthenticationType = 'AccessTokens'; AccountId = 'MSCloudLoginAssistant'; AccountType = 'AccessToken' }
        @{ AuthenticationType = 'Identity'; AccountId = 'MSI@50342'; AccountType = 'ManagedService' }
    ) {
        param ($AuthenticationType, $AccountId, $AccountType)
        InModuleScope 'MSCloudLoginAssistant' -Parameters @{ AuthenticationType = $AuthenticationType; AccountId = $AccountId; AccountType = $AccountType } {
            param ($AuthenticationType, $AccountId, $AccountType)
            $context = [PSCustomObject]@{
                Account      = [PSCustomObject]@{ Id = $AccountId; Type = $AccountType }
                Subscription = [PSCustomObject]@{ Id = 'expected-subscription' }
            }
            Mock -CommandName Get-AzContext -MockWith { $context }
            $credential = [System.Management.Automation.PSCredential]::new('admin@contoso.com', (ConvertTo-SecureString -String 'x' -AsPlainText -Force))
            $workloadProfile = [PSCustomObject]@{
                AuthenticationType = $AuthenticationType
                ApplicationId      = 'expected-app'
                Credentials        = $credential
                SubscriptionId     = 'expected-subscription'
            }

            & $Script:MSCloudLoginConnectionProbes.Azure $workloadProfile | Should -Not -BeNullOrEmpty
        }
    }
}

Describe 'Teams connection probe' {

    BeforeAll {
        function global:Get-CsTeamsCallingPolicy { }
    }

    BeforeEach {
        InModuleScope 'MSCloudLoginAssistant' {
            $Script:MSCloudLoginTeamsVerifiedTime = $null
        }
    }

    AfterAll {
        Remove-Item -Path 'Function:\Get-CsTeamsCallingPolicy' -ErrorAction SilentlyContinue
        [System.AppDomain]::CurrentDomain.SetData('MSCloudLoginAssistant.ConnectionIdentity.Teams', $null)
    }

    It 'Should not call Teams again within 3 minutes of a verification' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Get-CsTeamsCallingPolicy -MockWith { [PSCustomObject]@{ Identity = 'Global' } }
            $workloadProfile = [PSCustomObject]@{ AuthenticationType = 'ServicePrincipalWithThumbprint'; TenantId = 'contoso.onmicrosoft.com'; ApplicationId = 'expected-app'; Credentials = $null }
            Set-MSCloudLoginProcessConnectionIdentity -Workload 'Teams' -Identity (Get-MSCloudLoginConnectionIdentity -WorkloadProfile $workloadProfile)

            & $Script:MSCloudLoginConnectionProbes.Teams $workloadProfile | Should -Not -BeNullOrEmpty
            & $Script:MSCloudLoginConnectionProbes.Teams $workloadProfile | Should -Not -BeNullOrEmpty
            Should -Invoke -CommandName Get-CsTeamsCallingPolicy -Times 1 -Exactly
        }
    }

    It 'Should call Teams again after 3 minutes' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Get-CsTeamsCallingPolicy -MockWith { [PSCustomObject]@{ Identity = 'Global' } }
            $workloadProfile = [PSCustomObject]@{ AuthenticationType = 'ServicePrincipalWithThumbprint'; TenantId = 'contoso.onmicrosoft.com'; ApplicationId = 'expected-app'; Credentials = $null }
            Set-MSCloudLoginProcessConnectionIdentity -Workload 'Teams' -Identity (Get-MSCloudLoginConnectionIdentity -WorkloadProfile $workloadProfile)
            $Script:MSCloudLoginTeamsVerifiedTime = [System.DateTime]::UtcNow.AddMinutes(-4)

            & $Script:MSCloudLoginConnectionProbes.Teams $workloadProfile | Should -Not -BeNullOrEmpty
            Should -Invoke -CommandName Get-CsTeamsCallingPolicy -Times 1 -Exactly
        }
    }

    It 'Should reject another identity within 3 minutes of a verification' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Get-CsTeamsCallingPolicy -MockWith { [PSCustomObject]@{ Identity = 'Global' } }
            $workloadProfile = [PSCustomObject]@{ AuthenticationType = 'ServicePrincipalWithThumbprint'; TenantId = 'contoso.onmicrosoft.com'; ApplicationId = 'expected-app'; Credentials = $null }
            Set-MSCloudLoginProcessConnectionIdentity -Workload 'Teams' -Identity 'ServicePrincipalWithThumbprint|contoso.onmicrosoft.com|other-app|'
            $Script:MSCloudLoginTeamsVerifiedTime = [System.DateTime]::UtcNow

            & $Script:MSCloudLoginConnectionProbes.Teams $workloadProfile | Should -BeNullOrEmpty
        }
    }

    It 'Should reject a Teams session that <Description>' -TestCases @(
        @{ Description = 'another application connected'; RecordedApplicationId = 'other-app' }
        @{ Description = 'was disconnected'; RecordedApplicationId = $null }
    ) {
        param ($RecordedApplicationId)
        InModuleScope 'MSCloudLoginAssistant' -Parameters @{ RecordedApplicationId = $RecordedApplicationId } {
            param ($RecordedApplicationId)
            Mock -CommandName Get-CsTeamsCallingPolicy -MockWith { [PSCustomObject]@{ Identity = 'Global' } }
            $workloadProfile = [PSCustomObject]@{ AuthenticationType = 'ServicePrincipalWithThumbprint'; TenantId = 'contoso.onmicrosoft.com'; ApplicationId = 'expected-app'; Credentials = $null }
            if ($null -ne $RecordedApplicationId)
            {
                $recordedProfile = [PSCustomObject]@{ AuthenticationType = 'ServicePrincipalWithThumbprint'; TenantId = 'contoso.onmicrosoft.com'; ApplicationId = $RecordedApplicationId; Credentials = $null }
                Set-MSCloudLoginProcessConnectionIdentity -Workload 'Teams' -Identity (Get-MSCloudLoginConnectionIdentity -WorkloadProfile $recordedProfile)
            }
            else
            {
                Set-MSCloudLoginProcessConnectionIdentity -Workload 'Teams'
            }

            & $Script:MSCloudLoginConnectionProbes.Teams $workloadProfile | Should -BeNullOrEmpty
            Should -Invoke -CommandName Get-CsTeamsCallingPolicy -Times 0 -Exactly
        }
    }

    It 'Should accept the Teams session of the profile' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Get-CsTeamsCallingPolicy -MockWith { [PSCustomObject]@{ Identity = 'Global' } }
            $workloadProfile = [PSCustomObject]@{ AuthenticationType = 'ServicePrincipalWithThumbprint'; TenantId = 'contoso.onmicrosoft.com'; ApplicationId = 'expected-app'; Credentials = $null }
            Set-MSCloudLoginProcessConnectionIdentity -Workload 'Teams' -Identity (Get-MSCloudLoginConnectionIdentity -WorkloadProfile $workloadProfile)

            & $Script:MSCloudLoginConnectionProbes.Teams $workloadProfile | Should -Not -BeNullOrEmpty
        }
    }

    It 'Should keep the recorded identity visible to other runspaces' {
        InModuleScope 'MSCloudLoginAssistant' {
            Set-MSCloudLoginProcessConnectionIdentity -Workload 'Teams' -Identity 'recorded-identity'
        }
        $runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $runspace.Open()
        try
        {
            $powerShell = [System.Management.Automation.PowerShell]::Create()
            $powerShell.Runspace = $runspace
            $result = $powerShell.AddScript("[System.AppDomain]::CurrentDomain.GetData('MSCloudLoginAssistant.ConnectionIdentity.Teams')").Invoke()
            $powerShell.Dispose()
        }
        finally
        {
            $runspace.Dispose()
        }

        $result | Should -Be 'recorded-identity'
    }
}

Describe 'Test-MSCloudLoginParameterValueEmpty' {

    It 'Should treat <Description> as empty' -TestCases @(
        @{ Description = 'a null value'; Value = $null }
        @{ Description = 'an empty string'; Value = '' }
        @{ Description = 'an unset switch'; Value = [System.Management.Automation.SwitchParameter]::new($false) }
        @{ Description = 'a false boolean'; Value = $false }
        @{ Description = 'an empty secure string'; Value = (New-Object System.Security.SecureString) }
        @{ Description = 'an empty hashtable'; Value = @{} }
        @{ Description = 'an empty array'; Value = @() }
    ) {
        param ($Description, $Value)
        InModuleScope 'MSCloudLoginAssistant' -Parameters @{ Value = $Value } {
            param ($Value)
            (Test-MSCloudLoginParameterValueEmpty -Value $Value) | Should -BeTrue
        }
    }

    It 'Should treat <Description> as populated' -TestCases @(
        @{ Description = 'a non empty string'; Value = 'value' }
        @{ Description = 'a set switch'; Value = [System.Management.Automation.SwitchParameter]::new($true) }
        @{ Description = 'a true boolean'; Value = $true }
        @{ Description = 'a populated hashtable'; Value = @{ Key = 'value' } }
        @{ Description = 'a populated array'; Value = @('value') }
        @{ Description = 'a number'; Value = 42 }
    ) {
        param ($Description, $Value)
        InModuleScope 'MSCloudLoginAssistant' -Parameters @{ Value = $Value } {
            param ($Value)
            (Test-MSCloudLoginParameterValueEmpty -Value $Value) | Should -BeFalse
        }
    }
}

Describe 'Test-MSCloudLoginParameterValueEqual' {

    Context 'Secure strings' {
        It 'Should compare the decrypted values' {
            InModuleScope 'MSCloudLoginAssistant' {
                $left = ConvertTo-SecureString 'same-value' -AsPlainText -Force
                $right = ConvertTo-SecureString 'same-value' -AsPlainText -Force
                $other = ConvertTo-SecureString 'Same-Value' -AsPlainText -Force

                (Test-MSCloudLoginParameterValueEqual -KeyName 'CertificatePassword' -Left $left -Right $right) | Should -BeTrue
                (Test-MSCloudLoginParameterValueEqual -KeyName 'CertificatePassword' -Left $left -Right $other) | Should -BeFalse
            }
        }

        It 'Should never equal a plain string' {
            InModuleScope 'MSCloudLoginAssistant' {
                $secure = ConvertTo-SecureString 'same-value' -AsPlainText -Force
                (Test-MSCloudLoginParameterValueEqual -KeyName 'CertificatePassword' -Left $secure -Right 'same-value') | Should -BeFalse
            }
        }
    }

    Context 'Credentials' {
        It 'Should ignore the casing of the user name but not of the password' {
            InModuleScope 'MSCloudLoginAssistant' {
                $left = New-Object PSCredential ('user@contoso.com', (ConvertTo-SecureString 'Secret' -AsPlainText -Force))
                $sameCredential = New-Object PSCredential ('USER@contoso.com', (ConvertTo-SecureString 'Secret' -AsPlainText -Force))
                $otherPassword = New-Object PSCredential ('user@contoso.com', (ConvertTo-SecureString 'secret' -AsPlainText -Force))
                $otherUser = New-Object PSCredential ('other@contoso.com', (ConvertTo-SecureString 'Secret' -AsPlainText -Force))

                (Test-MSCloudLoginParameterValueEqual -KeyName 'Credentials' -Left $left -Right $sameCredential) | Should -BeTrue
                (Test-MSCloudLoginParameterValueEqual -KeyName 'Credentials' -Left $left -Right $otherPassword) | Should -BeFalse
                (Test-MSCloudLoginParameterValueEqual -KeyName 'Credentials' -Left $left -Right $otherUser) | Should -BeFalse
            }
        }

        It 'Should never equal a plain string' {
            InModuleScope 'MSCloudLoginAssistant' {
                $credential = New-Object PSCredential ('user@contoso.com', (ConvertTo-SecureString 'Secret' -AsPlainText -Force))
                (Test-MSCloudLoginParameterValueEqual -KeyName 'Credentials' -Left $credential -Right 'user@contoso.com') | Should -BeFalse
            }
        }
    }

    Context 'Dictionaries' {
        It 'Should compare the entries recursively' {
            InModuleScope 'MSCloudLoginAssistant' {
                $left = @{ ActiveDirectory = 'https://login.contoso.local'; Graph = @{ Url = 'https://graph.contoso.local' } }
                $same = @{ Graph = @{ Url = 'https://graph.contoso.local' }; ActiveDirectory = 'https://login.contoso.local' }
                $differentValue = @{ ActiveDirectory = 'https://login.fabrikam.local'; Graph = @{ Url = 'https://graph.contoso.local' } }
                $differentKey = @{ ActiveDirectory = 'https://login.contoso.local'; Teams = @{ Url = 'https://graph.contoso.local' } }
                $fewerEntries = @{ ActiveDirectory = 'https://login.contoso.local' }

                (Test-MSCloudLoginParameterValueEqual -KeyName 'Endpoints' -Left $left -Right $same) | Should -BeTrue
                (Test-MSCloudLoginParameterValueEqual -KeyName 'Endpoints' -Left $left -Right $differentValue) | Should -BeFalse
                (Test-MSCloudLoginParameterValueEqual -KeyName 'Endpoints' -Left $left -Right $differentKey) | Should -BeFalse
                (Test-MSCloudLoginParameterValueEqual -KeyName 'Endpoints' -Left $left -Right $fewerEntries) | Should -BeFalse
                (Test-MSCloudLoginParameterValueEqual -KeyName 'Endpoints' -Left $left -Right 'not-a-dictionary') | Should -BeFalse
            }
        }
    }

    Context 'Collections' {
        It 'Should compare access tokens position by position' {
            InModuleScope 'MSCloudLoginAssistant' {
                (Test-MSCloudLoginParameterValueEqual -KeyName 'AccessTokens' -Left @('a', 'b') -Right @('a', 'b')) | Should -BeTrue
                (Test-MSCloudLoginParameterValueEqual -KeyName 'AccessTokens' -Left @('a', 'b') -Right @('b', 'a')) | Should -BeFalse
                (Test-MSCloudLoginParameterValueEqual -KeyName 'AccessTokens' -Left @('a', 'b') -Right @('a')) | Should -BeFalse
            }
        }

        It 'Should compare cmdlet names regardless of order and casing' {
            InModuleScope 'MSCloudLoginAssistant' {
                (Test-MSCloudLoginParameterValueEqual -KeyName 'CmdletsToLoad' -Left @('Get-Mailbox', 'Set-Mailbox') -Right @('set-mailbox', 'get-mailbox')) | Should -BeTrue
                (Test-MSCloudLoginParameterValueEqual -KeyName 'CmdletsToLoad' -Left @('Get-Mailbox') -Right @('Get-User')) | Should -BeFalse
            }
        }
    }

    Context 'Scalars' {
        It 'Should compare booleans and switches by their truth value' {
            InModuleScope 'MSCloudLoginAssistant' {
                (Test-MSCloudLoginParameterValueEqual -KeyName 'Identity' -Left $true -Right ([System.Management.Automation.SwitchParameter]::new($true))) | Should -BeTrue
                (Test-MSCloudLoginParameterValueEqual -KeyName 'Identity' -Left $true -Right $false) | Should -BeFalse
            }
        }

        It 'Should compare identifiers case insensitively and secrets case sensitively' {
            InModuleScope 'MSCloudLoginAssistant' {
                (Test-MSCloudLoginParameterValueEqual -KeyName 'TenantId' -Left 'Contoso.com' -Right 'contoso.com') | Should -BeTrue
                (Test-MSCloudLoginParameterValueEqual -KeyName 'ApplicationSecret' -Left 'Secret' -Right 'secret') | Should -BeFalse
                (Test-MSCloudLoginParameterValueEqual -KeyName 'Credentials.Password' -Left 'Secret' -Right 'secret') | Should -BeFalse
            }
        }
    }
}
