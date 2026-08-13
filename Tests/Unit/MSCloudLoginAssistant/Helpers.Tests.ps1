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
                Mock -CommandName Get-Item -MockWith {
                    if ($Path -like 'Cert:\CurrentUser\My\*')
                    {
                        return [PSCustomObject]@{ Thumbprint = 'AA11'; Store = 'CurrentUser' }
                    }
                    return $null
                }

                $certificate = Get-MSCloudLoginCertificate -CertificateThumbprint 'AA11'
                $certificate.Store | Should -Be 'CurrentUser'
                Should -Invoke Get-Item -Exactly 1
            }
        }

        It 'Should fall back to the local machine store' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Get-Item -MockWith {
                    if ($Path -like 'Cert:\LocalMachine\My\*')
                    {
                        return [PSCustomObject]@{ Thumbprint = 'AA11'; Store = 'LocalMachine' }
                    }
                    return $null
                }

                $certificate = Get-MSCloudLoginCertificate -CertificateThumbprint 'AA11'
                $certificate.Store | Should -Be 'LocalMachine'
                Should -Invoke Get-Item -Exactly 2
            }
        }

        It 'Should name both stores when the certificate does not exist' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Get-Item -MockWith { return $null }

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
