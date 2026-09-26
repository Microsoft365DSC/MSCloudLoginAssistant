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
        It 'Should return the certificate from the <Store> store after <Lookups> lookup(s)' -TestCases @(
            @{ Store = 'CurrentUser'; Lookups = 1 }
            @{ Store = 'LocalMachine'; Lookups = 2 }
        ) {
            param ($Store, $Lookups)
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{ Store = $Store; Lookups = $Lookups } {
                param ($Store, $Lookups)
                Mock -CommandName Find-MSCloudLoginStoreCertificate -MockWith {
                    if ($StoreLocation -eq $Store)
                    {
                        return [PSCustomObject]@{ Thumbprint = 'AA11'; Store = $Store }
                    }
                    return $null
                }

                $certificate = Get-MSCloudLoginCertificate -CertificateThumbprint 'AA11'
                $certificate.Store | Should -Be $Store
                Should -Invoke Find-MSCloudLoginStoreCertificate -Exactly $Lookups
            }
        }

        It 'Should return a store certificate without Cert: drive properties and $null for an unknown thumbprint' {
            InModuleScope 'MSCloudLoginAssistant' {
                Find-MSCloudLoginStoreCertificate -StoreLocation 'CurrentUser' -CertificateThumbprint '0000000000000000000000000000000000000000' | Should -BeNullOrEmpty

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

        It 'Should load a PFX file that has no password and throw for a file that does not exist' {
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{ PfxPath = $script:unprotectedPfxPath } {
                param ($PfxPath)
                $certificate = Get-MSCloudLoginCertificate -CertificatePath $PfxPath
                $certificate.Subject | Should -Be 'CN=MSCloudLoginAssistantHelperTest'

                { Get-MSCloudLoginCertificate -CertificatePath 'C:\does\not\exist.pfx' } |
                    Should -Throw "*'C:\does\not\exist.pfx' was not found*"
            }
        }
    }
}

Describe 'Remove-MSCloudLoginProxyModule' {

    It 'Should remove only the loaded modules that export the probe command' {
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

            Remove-MSCloudLoginProxyModule -ProbeCommand 'Get-OrganizationConfig' -Source 'Test'

            Should -Invoke Remove-Module -Exactly 1
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

        It 'Should give the module that exports the probe command precedence, back and forth, without reloading either module' {
            Get-MSCLARestoreShared | Should -Be 'SecurityCompliance'

            InModuleScope 'MSCloudLoginAssistant' -Parameters @{ ProxyDirectory = $TestDrive } {
                param ($ProxyDirectory)
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Get-ConnectionInformation -MockWith { [PSCustomObject]@{ ModuleName = $ProxyDirectory } }

                Restore-MSCloudLoginProxyModule -ProbeCommand 'Get-MSCLARestoreExoProbe' -Source 'Test' | Should -BeTrue
                (Get-Command -Name 'Get-MSCLARestoreShared').Module.Name | Should -Be 'tmpEXO_restoreexo'

                foreach ($iteration in 1..3)
                {
                    Restore-MSCloudLoginProxyModule -ProbeCommand 'Get-MSCLARestoreScProbe' -Source 'Test' | Should -BeTrue
                    Get-MSCLARestoreShared | Should -Be 'SecurityCompliance'

                    Restore-MSCloudLoginProxyModule -ProbeCommand 'Get-MSCLARestoreExoProbe' -Source 'Test' | Should -BeTrue
                    Get-MSCLARestoreShared | Should -Be 'ExchangeOnline'
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

            Restore-MSCloudLoginProxyModule -ProbeCommand 'Get-OrganizationConfig' -Source 'Test' | Should -BeFalse

            Should -Invoke Import-Module -Exactly 0
        }
    }

    It 'Should report a failed import' {
        InModuleScope 'MSCloudLoginAssistant' {
            $proxyModule = New-Module -Name 'tmpEXO_abc' -ScriptBlock { function Get-OrganizationConfig { } }
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
            Mock -CommandName Import-Module -MockWith { throw 'import failed' }
            Mock -CommandName Get-Module -MockWith { return $proxyModule }
            Mock -CommandName Get-ConnectionInformation -MockWith { [PSCustomObject]@{ ModuleName = $proxyModule.ModuleBase } }

            Restore-MSCloudLoginProxyModule -ProbeCommand 'Get-OrganizationConfig' -Source 'Test' | Should -BeFalse

            Should -Invoke Add-MSCloudLoginAssistantEvent -ParameterFilter { $Message -like '*Failed to restore proxy module {tmpEXO_abc}*import failed*' }
        }
    }

    It 'Should not restore a loaded proxy module whose connection is gone' {
        InModuleScope 'MSCloudLoginAssistant' {
            $proxyModule = New-Module -Name 'tmpEXO_stale' -ScriptBlock { function Get-OrganizationConfig { } }
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
            Mock -CommandName Import-Module -MockWith { }
            Mock -CommandName Get-Module -MockWith { return $proxyModule }
            Mock -CommandName Get-ConnectionInformation -MockWith { return $null }

            Restore-MSCloudLoginProxyModule -ProbeCommand 'Get-OrganizationConfig' -Source 'Test' | Should -BeFalse

            Should -Invoke Import-Module -Exactly 0
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
                    [PSCustomObject]@{ ConnectionId = [guid]'11111111-1111-1111-1111-111111111111'; IsEopSession = $false; ModuleName = 'C:\Temp\tmpEXO_one' }
                    [PSCustomObject]@{ ConnectionId = [guid]'22222222-2222-2222-2222-222222222222'; IsEopSession = $true; ModuleName = 'C:\Temp\tmpEXO_two' }
                    [PSCustomObject]@{ ConnectionId = [guid]'33333333-3333-3333-3333-333333333333'; IsEopSession = $false; ModuleName = 'C:\Temp\tmpEXO_three' }
                    [PSCustomObject]@{ ConnectionId = [guid]'44444444-4444-4444-4444-444444444444'; IsEopSession = $false; ModuleName = 'C:\Temp\tmpEXO_otherrunspace' }
                )
            }
            Mock -CommandName Get-Module -MockWith {
                return @('C:\Temp\tmpEXO_one', 'C:\Temp\tmpEXO_two', 'C:\Temp\tmpEXO_three') | ForEach-Object -Process {
                    [PSCustomObject]@{ ModuleBase = $_ }
                }
            }
        }
    }

    It 'Should disconnect only the <Kind> connections of the modules loaded in this runspace' -TestCases @(
        @{ Kind = 'Exchange Online'; SecurityCompliance = $false; ExpectedConnectionIds = '11111111-1111-1111-1111-111111111111,33333333-3333-3333-3333-333333333333' }
        @{ Kind = 'Security & Compliance'; SecurityCompliance = $true; ExpectedConnectionIds = '22222222-2222-2222-2222-222222222222' }
    ) {
        param ($SecurityCompliance, $ExpectedConnectionIds)
        InModuleScope 'MSCloudLoginAssistant' -Parameters @{ SecurityCompliance = $SecurityCompliance; ExpectedConnectionIds = $ExpectedConnectionIds } {
            param ($SecurityCompliance, $ExpectedConnectionIds)
            Disconnect-MSCloudLoginExchangeConnection -SecurityCompliance:$SecurityCompliance -Source 'Test'

            Should -Invoke Disconnect-ExchangeOnline -Exactly 1 -ParameterFilter {
                ($ConnectionId -join ',') -eq $ExpectedConnectionIds
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

Describe 'Test-MSCloudLoginMFARequiredError' {

    It 'Should return <Expected> for "<Message>" with the additional patterns <AdditionalPatterns>' -TestCases @(
        @{ Message = 'AADSTS50076: Due to a configuration change made by your administrator...'; AdditionalPatterns = @(); Expected = $true }
        @{ Message = 'you must use multi-factor authentication to access this resource'; AdditionalPatterns = @(); Expected = $true }
        @{ Message = 'WAM Error 12345'; AdditionalPatterns = @(); Expected = $false }
        @{ Message = 'WAM Error 12345'; AdditionalPatterns = @('*WAM Error*'); Expected = $true }
        @{ Message = 'The sign-in name or password is incorrect'; AdditionalPatterns = @(); Expected = $false }
    ) {
        InModuleScope 'MSCloudLoginAssistant' -Parameters @{ Message = $Message; AdditionalPatterns = $AdditionalPatterns; Expected = $Expected } {
            param ($Message, $AdditionalPatterns, $Expected)
            $err = $null
            try { throw $Message } catch { $err = $_ }
            (Test-MSCloudLoginMFARequiredError -ErrorRecord $err -AdditionalPatterns $AdditionalPatterns) | Should -Be $Expected
        }
    }
}

Describe 'Get-MSCloudLoginSPOUrlFromTenantId' {

    It 'Should derive <AdminUrl> for <TenantId> in <EnvironmentName>' -TestCases @(
        @{ TenantId = 'contoso.onmicrosoft.com'; EnvironmentName = 'AzureCloud'; AdminUrl = 'https://contoso-admin.sharepoint.com'; ConnectionUrl = 'https://contoso.sharepoint.com' }
        @{ TenantId = 'contoso.onmicrosoft.com'; EnvironmentName = 'AzureUSGovernment'; AdminUrl = 'https://contoso-admin.sharepoint.us'; ConnectionUrl = 'https://contoso.sharepoint.us' }
        @{ TenantId = 'contoso.onmicrosoft.com'; EnvironmentName = 'AzureDOD'; AdminUrl = 'https://contoso-admin.sharepoint-mil.us'; ConnectionUrl = 'https://contoso.sharepoint-mil.us' }
        @{ TenantId = 'contoso.partner.onmschina.cn'; EnvironmentName = 'AzureChinaCloud'; AdminUrl = 'https://contoso-admin.sharepoint.cn'; ConnectionUrl = 'https://contoso.sharepoint.cn' }
        @{ TenantId = 'contoso.onms.fr'; EnvironmentName = 'AzureFranceCloud'; AdminUrl = 'https://contoso-admin.spo.fr'; ConnectionUrl = 'https://contoso.spo.fr' }
    ) {
        InModuleScope 'MSCloudLoginAssistant' -Parameters @{ TenantId = $TenantId; EnvironmentName = $EnvironmentName; AdminUrl = $AdminUrl; ConnectionUrl = $ConnectionUrl } {
            param ($TenantId, $EnvironmentName, $AdminUrl, $ConnectionUrl)
            $result = Get-MSCloudLoginSPOUrlFromTenantId -TenantId $TenantId -EnvironmentName $EnvironmentName
            $result.AdminUrl      | Should -Be $AdminUrl
            $result.ConnectionUrl | Should -Be $ConnectionUrl
        }
    }

    It 'Should throw for an unrecognized tenant format' {
        InModuleScope 'MSCloudLoginAssistant' {
            { Get-MSCloudLoginSPOUrlFromTenantId -TenantId 'contoso.com' -EnvironmentName 'AzureCloud' } | Should -Throw
        }
    }
}

Describe 'Get-MSCloudLoginAccessTokenValue' {

    It 'Should return the plain value of a string, SecureString or PSCredential token' {
        InModuleScope 'MSCloudLoginAssistant' {
            $secure = ConvertTo-SecureString 'secure-token' -AsPlainText -Force
            $cred = New-Object PSCredential ('token', (ConvertTo-SecureString 'cred-token' -AsPlainText -Force))

            (Get-MSCloudLoginAccessTokenValue -Token 'plain-token') | Should -Be 'plain-token'
            (Get-MSCloudLoginAccessTokenValue -Token $secure) | Should -Be 'secure-token'
            (Get-MSCloudLoginAccessTokenValue -Token $cred) | Should -Be 'cred-token'
        }
    }
}

Describe 'Get-MSCloudLoginTenantDomainFromCredentials' {

    It 'Should return the domain part of a UPN and throw when the user name is not a UPN' {
        InModuleScope 'MSCloudLoginAssistant' {
            $upn = New-Object PSCredential ('user@contoso.com', (ConvertTo-SecureString 'pwd' -AsPlainText -Force))
            $downLevel = New-Object PSCredential ('CONTOSO\user', (ConvertTo-SecureString 'pwd' -AsPlainText -Force))

            (Get-MSCloudLoginTenantDomainFromCredentials -Credentials $upn) | Should -Be 'contoso.com'
            { Get-MSCloudLoginTenantDomainFromCredentials -Credentials $downLevel } | Should -Throw
        }
    }
}

Describe 'Get-MSCloudLoginEndpointInfo' {

    It 'Should resolve <Workload>/<Environment> endpoints' -TestCases @(
        @{ Workload = 'AdminAPI'; Environment = 'AzureCloud'; Property = 'AuthorizationUrl'; Expected = 'https://login.microsoftonline.com' }
        @{ Workload = 'AdminAPI'; Environment = 'AzureDOD'; Property = 'AuthorizationUrl'; Expected = 'https://login.microsoftonline.us' }
        @{ Workload = 'AzureDevOPS'; Environment = 'AzureDOD'; Property = 'HostUrl'; Expected = 'https://dev.azure.us' }
        @{ Workload = 'DefenderForEndpoint'; Environment = 'AzureUSGovernment'; Property = 'HostUrl'; Expected = 'https://api-gcc.securitycenter.microsoft.us' }
        @{ Workload = 'Fabric'; Environment = 'AzureCloud'; Property = 'Scope'; Expected = 'https://api.fabric.microsoft.com/.default' }
        @{ Workload = 'Fabric'; Environment = 'SomethingElse'; Property = 'AuthorizationUrl'; Expected = 'https://login.microsoftonline.com' }
        @{ Workload = 'Licensing'; Environment = 'AzureCloud'; Property = 'HostUrl'; Expected = 'https://licensing.m365.microsoft.com' }
        @{ Workload = 'MicrosoftGraph'; Environment = 'AzureCloud'; Property = 'ResourceUrl'; Expected = 'https://graph.microsoft.com/' }
        @{ Workload = 'MicrosoftGraph'; Environment = 'AzureGermanyCloud'; Property = 'GraphEnvironment'; Expected = 'DelosCloud' }
        @{ Workload = 'O365Portal'; Environment = 'AzureDOD'; Property = 'AuthorizationUrl'; Expected = 'https://login.microsoftonline.us' }
        @{ Workload = 'PowerPlatformREST'; Environment = 'AzureDOD'; Property = 'BapEndpoint'; Expected = 'api.bap.appsplatform.us' }
        @{ Workload = 'SecurityComplianceCenter'; Environment = 'AzureChinaCloud'; Property = 'ConnectionUrl'; Expected = 'https://ps.compliance.protection.partner.outlook.cn/powershell-liveid/' }
        @{ Workload = 'SecurityComplianceCenter'; Environment = 'AzureFranceCloud'; Property = 'AuthorizationUrl'; Expected = 'https://login.sovcloud-identity.fr/organizations' }
        @{ Workload = 'Tasks'; Environment = 'AzureUSGovernment'; Property = 'HostUrl'; Expected = 'https://tasks.office365.us' }
        @{ Workload = 'Tasks'; Environment = 'AzureFranceCloud'; Property = 'AuthorizationUrl'; Expected = 'https://login.sovcloud-identity.fr' }
    ) {
        param ($Workload, $Environment, $Property, $Expected)
        InModuleScope 'MSCloudLoginAssistant' -Parameters @{ Workload = $Workload; Environment = $Environment; Property = $Property; Expected = $Expected } {
            $result = Get-MSCloudLoginEndpointInfo -Workload $Workload -EnvironmentName $Environment
            $result[$Property] | Should -Be $Expected
        }
    }

    It 'Should throw for an unknown workload and when neither the environment nor a default entry is defined' {
        InModuleScope 'MSCloudLoginAssistant' {
            $originalEndpointData = $Script:WorkloadEndpointData
            try
            {
                $Script:WorkloadEndpointData = @{
                    TestWorkload = @{ AzureCloud = @{ HostUrl = 'https://contoso.local' } }
                }
                { Get-MSCloudLoginEndpointInfo -Workload 'DoesNotExist' -EnvironmentName 'AzureCloud' } |
                    Should -Throw "No endpoint information is defined for workload 'DoesNotExist'."
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

    BeforeAll {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
        }
    }

    It 'Should reuse a fresh token-based connection and reject one that is not connected, has no timestamp or is expired' {
        InModuleScope 'MSCloudLoginAssistant' {
            $notConnected = New-Object AdminAPI

            $noTimestamp = New-Object AdminAPI
            $noTimestamp.Connected = $true

            $fresh = New-Object AdminAPI
            $fresh.AuthenticationType = 'ServicePrincipalWithSecret'
            $fresh.CompleteConnection()

            $expired = New-Object AdminAPI
            $expired.AuthenticationType = 'ServicePrincipalWithSecret'
            $expired.CompleteConnection()
            $expired.ConnectedDateTime = [System.DateTime]::Now.AddMinutes(-60).ToString()

            (Test-MSCloudLoginConnectionReusable -WorkloadProfile $notConnected -Source 'Test') | Should -BeFalse
            (Test-MSCloudLoginConnectionReusable -WorkloadProfile $noTimestamp -Source 'Test') | Should -BeFalse
            (Test-MSCloudLoginConnectionReusable -WorkloadProfile $fresh -Source 'Test') | Should -BeTrue
            (Test-MSCloudLoginConnectionReusable -WorkloadProfile $expired -Source 'Test') | Should -BeFalse

            $noTimestamp.Connected | Should -BeFalse
            $expired.Connected | Should -BeFalse
        }
    }

    Context 'When the token expiry is known' {
        It 'Should renew a <AuthenticationType> connection only when its token expires within five minutes, even after 50 minutes' -TestCases @(
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

                $workloadProfile.Connected = $true
                $workloadProfile.ConnectedDateTime = [System.DateTime]::Now.AddMinutes(-70).ToString()
                $workloadProfile.TokenExpiresOn = [System.DateTime]::Now.AddMinutes(20)
                Test-MSCloudLoginConnectionReusable -WorkloadProfile $workloadProfile -TokenBasedAuthTypes @('Identity', 'ServicePrincipalWithThumbprint') -Source 'Test' | Should -BeTrue
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

    It 'Should reuse the connection while the probe returns a context and treat a failing probe as a lost connection' {
        InModuleScope 'MSCloudLoginAssistant' {
            $workloadProfile = New-Object AdminAPI
            $workloadProfile.AuthenticationType = 'ServicePrincipalWithThumbprint'
            $workloadProfile.CompleteConnection()

            (Test-MSCloudLoginConnectionReusable -WorkloadProfile $workloadProfile `
                -ProbeScript { return @{ TenantId = 'contoso' } } -Source 'Test') | Should -BeTrue

            $result = Test-MSCloudLoginConnectionReusable -WorkloadProfile $workloadProfile `
                -ProbeScript { throw 'the SDK context is gone' } -Source 'Test'

            $result | Should -BeFalse
            $workloadProfile.Connected | Should -BeFalse
            Should -Invoke Add-MSCloudLoginAssistantEvent -ParameterFilter { $Message -like 'Connection probe failed*' }
        }
    }
}

Describe 'Microsoft Graph connection probe' {

    AfterAll {
        [System.AppDomain]::CurrentDomain.SetData('MSCloudLoginAssistant.ConnectionIdentity.MicrosoftGraph', $null)
    }

    It 'Should reject a Graph context of <Description>' -TestCases @(
        @{ Description = 'another application'; AuthenticationType = 'ServicePrincipalWithThumbprint'; ApplicationId = 'expected-app'; UserName = $null; ClientId = 'other-app'; Account = $null; RecordedIdentity = $null }
        @{ Description = 'another account'; AuthenticationType = 'Credentials'; ApplicationId = 'app'; UserName = 'admin@contoso.com'; ClientId = 'app'; Account = 'other@contoso.com'; RecordedIdentity = $null }
        @{ Description = 'another identity than the managed identity profile'; AuthenticationType = 'Identity'; ApplicationId = $null; UserName = $null; ClientId = 'other-app'; Account = $null; RecordedIdentity = 'ServicePrincipalWithThumbprint|contoso.onmicrosoft.com|other-app|' }
    ) {
        param ($AuthenticationType, $ApplicationId, $UserName, $ClientId, $Account, $RecordedIdentity)
        InModuleScope 'MSCloudLoginAssistant' -Parameters @{ AuthenticationType = $AuthenticationType; ApplicationId = $ApplicationId; UserName = $UserName; ClientId = $ClientId; Account = $Account; RecordedIdentity = $RecordedIdentity } {
            param ($AuthenticationType, $ApplicationId, $UserName, $ClientId, $Account, $RecordedIdentity)
            Mock -CommandName Get-MgContext -MockWith { [PSCustomObject]@{ ClientId = $ClientId; Account = $Account } }
            $credential = $null
            if ($null -ne $UserName)
            {
                $credential = [System.Management.Automation.PSCredential]::new($UserName, (ConvertTo-SecureString -String 'x' -AsPlainText -Force))
            }
            $workloadProfile = [PSCustomObject]@{ AuthenticationType = $AuthenticationType; TenantId = 'contoso.onmicrosoft.com'; ApplicationId = $ApplicationId; Credentials = $credential }
            if ($null -eq $RecordedIdentity)
            {
                $RecordedIdentity = Get-MSCloudLoginConnectionIdentity -WorkloadProfile $workloadProfile
            }
            Set-MSCloudLoginProcessConnectionIdentity -Workload 'MicrosoftGraph' -Identity $RecordedIdentity

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

    It 'Should accept the Teams session of the profile and call Teams again only after 3 minutes' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Get-CsTeamsCallingPolicy -MockWith { [PSCustomObject]@{ Identity = 'Global' } }
            $workloadProfile = [PSCustomObject]@{ AuthenticationType = 'ServicePrincipalWithThumbprint'; TenantId = 'contoso.onmicrosoft.com'; ApplicationId = 'expected-app'; Credentials = $null }
            Set-MSCloudLoginProcessConnectionIdentity -Workload 'Teams' -Identity (Get-MSCloudLoginConnectionIdentity -WorkloadProfile $workloadProfile)

            & $Script:MSCloudLoginConnectionProbes.Teams $workloadProfile | Should -Not -BeNullOrEmpty
            & $Script:MSCloudLoginConnectionProbes.Teams $workloadProfile | Should -Not -BeNullOrEmpty
            Should -Invoke -CommandName Get-CsTeamsCallingPolicy -Times 1 -Exactly

            $Script:MSCloudLoginTeamsVerifiedTime = [System.DateTime]::UtcNow.AddMinutes(-4)
            & $Script:MSCloudLoginConnectionProbes.Teams $workloadProfile | Should -Not -BeNullOrEmpty
            Should -Invoke -CommandName Get-CsTeamsCallingPolicy -Times 2 -Exactly
        }
    }

    It 'Should reject a Teams session that <Description>' -TestCases @(
        @{ Description = 'another application connected'; RecordedApplicationId = 'other-app'; VerifiedMinutesAgo = $null }
        @{ Description = 'another application connected within 3 minutes of a verification'; RecordedApplicationId = 'other-app'; VerifiedMinutesAgo = 0 }
        @{ Description = 'was disconnected'; RecordedApplicationId = $null; VerifiedMinutesAgo = $null }
    ) {
        param ($RecordedApplicationId, $VerifiedMinutesAgo)
        InModuleScope 'MSCloudLoginAssistant' -Parameters @{ RecordedApplicationId = $RecordedApplicationId; VerifiedMinutesAgo = $VerifiedMinutesAgo } {
            param ($RecordedApplicationId, $VerifiedMinutesAgo)
            Mock -CommandName Get-CsTeamsCallingPolicy -MockWith { [PSCustomObject]@{ Identity = 'Global' } }
            if ($null -ne $VerifiedMinutesAgo)
            {
                $Script:MSCloudLoginTeamsVerifiedTime = [System.DateTime]::UtcNow.AddMinutes(-$VerifiedMinutesAgo)
            }
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

    It 'Should treat <Description> as <State>' -TestCases @(
        @{ Description = 'a null value'; Value = $null; State = 'empty' }
        @{ Description = 'an empty string'; Value = ''; State = 'empty' }
        @{ Description = 'an unset switch'; Value = [System.Management.Automation.SwitchParameter]::new($false); State = 'empty' }
        @{ Description = 'a false boolean'; Value = $false; State = 'empty' }
        @{ Description = 'an empty secure string'; Value = (New-Object System.Security.SecureString); State = 'empty' }
        @{ Description = 'an empty hashtable'; Value = @{}; State = 'empty' }
        @{ Description = 'an empty array'; Value = @(); State = 'empty' }
        @{ Description = 'a non empty string'; Value = 'value'; State = 'populated' }
        @{ Description = 'a set switch'; Value = [System.Management.Automation.SwitchParameter]::new($true); State = 'populated' }
        @{ Description = 'a true boolean'; Value = $true; State = 'populated' }
        @{ Description = 'a populated hashtable'; Value = @{ Key = 'value' }; State = 'populated' }
        @{ Description = 'a populated array'; Value = @('value'); State = 'populated' }
        @{ Description = 'a number'; Value = 42; State = 'populated' }
    ) {
        param ($Description, $Value, $State)
        InModuleScope 'MSCloudLoginAssistant' -Parameters @{ Value = $Value; State = $State } {
            param ($Value, $State)
            (Test-MSCloudLoginParameterValueEmpty -Value $Value) | Should -Be ($State -eq 'empty')
        }
    }
}

Describe 'Test-MSCloudLoginParameterValueEqual' {

    Context 'Secure strings' {
        It 'Should compare the decrypted values and never equal a plain string' {
            InModuleScope 'MSCloudLoginAssistant' {
                $left = ConvertTo-SecureString 'same-value' -AsPlainText -Force
                $right = ConvertTo-SecureString 'same-value' -AsPlainText -Force
                $other = ConvertTo-SecureString 'Same-Value' -AsPlainText -Force

                (Test-MSCloudLoginParameterValueEqual -KeyName 'CertificatePassword' -Left $left -Right $right) | Should -BeTrue
                (Test-MSCloudLoginParameterValueEqual -KeyName 'CertificatePassword' -Left $left -Right $other) | Should -BeFalse
                (Test-MSCloudLoginParameterValueEqual -KeyName 'CertificatePassword' -Left $left -Right 'same-value') | Should -BeFalse
            }
        }
    }

    Context 'Credentials' {
        It 'Should ignore the casing of the user name but not of the password and never equal a plain string' {
            InModuleScope 'MSCloudLoginAssistant' {
                $left = New-Object PSCredential ('user@contoso.com', (ConvertTo-SecureString 'Secret' -AsPlainText -Force))
                $sameCredential = New-Object PSCredential ('USER@contoso.com', (ConvertTo-SecureString 'Secret' -AsPlainText -Force))
                $otherPassword = New-Object PSCredential ('user@contoso.com', (ConvertTo-SecureString 'secret' -AsPlainText -Force))
                $otherUser = New-Object PSCredential ('other@contoso.com', (ConvertTo-SecureString 'Secret' -AsPlainText -Force))

                (Test-MSCloudLoginParameterValueEqual -KeyName 'Credentials' -Left $left -Right $sameCredential) | Should -BeTrue
                (Test-MSCloudLoginParameterValueEqual -KeyName 'Credentials' -Left $left -Right $otherPassword) | Should -BeFalse
                (Test-MSCloudLoginParameterValueEqual -KeyName 'Credentials' -Left $left -Right $otherUser) | Should -BeFalse
                (Test-MSCloudLoginParameterValueEqual -KeyName 'Credentials' -Left $left -Right 'user@contoso.com') | Should -BeFalse
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
