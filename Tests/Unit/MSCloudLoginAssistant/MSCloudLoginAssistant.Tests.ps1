#Requires -Modules Pester

BeforeAll {
    # Ensure the Graph dependency check passes during module import.
    # If the real module is not installed, create a temporary stub manifest.
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

    # Import the module under test
    $moduleRoot = Resolve-Path (Join-Path $PSScriptRoot '..\..\..\Modules\MSCloudLoginAssistant')
    Import-Module (Join-Path $moduleRoot 'MSCloudLoginAssistant.psd1') -Force
}

AfterAll {
    # Clean up temporary stub module if we created one
    if ($script:tempModuleBase -and (Test-Path $script:tempModuleBase))
    {
        Remove-Item -Path $script:tempModuleBase -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# ---------------------------------------------------------------------------
# Get-AuthenticationTypeFromParameters
# ---------------------------------------------------------------------------
Describe 'Get-AuthenticationTypeFromParameters' {

    It 'Should return <Expected> for the parameters <Keys>' -TestCases @(
        @{ Keys = @('ApplicationId', 'TenantId', 'CertificateThumbprint'); Expected = 'ServicePrincipalWithThumbprint' }
        @{ Keys = @('ApplicationId', 'TenantId', 'ApplicationSecret'); Expected = 'ServicePrincipalWithSecret' }
        @{ Keys = @('ApplicationId', 'TenantId', 'CertificatePath', 'CertificatePassword'); Expected = 'ServicePrincipalWithPath' }
        @{ Keys = @('Credentials', 'ApplicationId'); Expected = 'CredentialsWithApplicationId' }
        @{ Keys = @('Credentials', 'TenantId'); Expected = 'CredentialsWithTenantId' }
        @{ Keys = @('Credentials'); Expected = 'Credentials' }
        @{ Keys = @('Identity'); Expected = 'Identity' }
        @{ Keys = @('AccessTokens', 'TenantId'); Expected = 'AccessTokens' }
        @{ Keys = @(); Expected = 'Interactive' }
    ) {
        InModuleScope 'MSCloudLoginAssistant' -Parameters @{ Keys = $Keys; Expected = $Expected } {
            param ($Keys, $Expected)
            $secPwd = ConvertTo-SecureString 'pass' -AsPlainText -Force
            $values = @{
                ApplicationId         = 'app-id'
                TenantId              = 'tenant-id'
                CertificateThumbprint = 'thumb'
                ApplicationSecret     = 'secret'
                CertificatePath       = 'C:\cert.pfx'
                CertificatePassword   = $secPwd
                Credentials           = New-Object PSCredential ('user@contoso.com', $secPwd)
                Identity              = $true
                AccessTokens          = @('token1', 'token2')
            }
            $params = @{}
            foreach ($key in $Keys)
            {
                $params[$key] = $values[$key]
            }
            $result = Get-AuthenticationTypeFromParameters -AuthenticationObject $params
            $result | Should -Be $Expected
        }
    }
}

# ---------------------------------------------------------------------------
# MSCloudLoginConnectionProfile class
# ---------------------------------------------------------------------------
Describe 'MSCloudLoginConnectionProfile' {

    Context 'Constructor defaults' {
        It 'Should initialise all workload objects and set CreatedTime' {
            InModuleScope 'MSCloudLoginAssistant' {
                $cloudProfile = New-Object MSCloudLoginConnectionProfile

                $cloudProfile.CreatedTime              | Should -Not -BeNullOrEmpty
                $cloudProfile.AdminAPI                 | Should -Not -BeNullOrEmpty
                $cloudProfile.Azure                    | Should -Not -BeNullOrEmpty
                $cloudProfile.AzureDevOPS              | Should -Not -BeNullOrEmpty
                $cloudProfile.DefenderForEndpoint      | Should -Not -BeNullOrEmpty
                $cloudProfile.EngageHub                | Should -Not -BeNullOrEmpty
                $cloudProfile.ExchangeOnline           | Should -Not -BeNullOrEmpty
                $cloudProfile.Fabric                   | Should -Not -BeNullOrEmpty
                $cloudProfile.Licensing                | Should -Not -BeNullOrEmpty
                $cloudProfile.O365Portal               | Should -Not -BeNullOrEmpty
                $cloudProfile.MicrosoftGraph           | Should -Not -BeNullOrEmpty
                $cloudProfile.PnP                      | Should -Not -BeNullOrEmpty
                $cloudProfile.PowerPlatform            | Should -Not -BeNullOrEmpty
                $cloudProfile.PowerPlatformREST        | Should -Not -BeNullOrEmpty
                $cloudProfile.SecurityComplianceCenter | Should -Not -BeNullOrEmpty
                $cloudProfile.SharePointOnlineREST     | Should -Not -BeNullOrEmpty
                $cloudProfile.Tasks                    | Should -Not -BeNullOrEmpty
                $cloudProfile.Teams                    | Should -Not -BeNullOrEmpty
            }
        }
    }

    Context 'Workload default ApplicationIds' {
        It 'Should set correct default ApplicationId for <Workload>' -TestCases @(
            @{ Workload = 'AdminAPI'; Expected = '1950a258-227b-4e31-a9cf-717495945fc2' }
            @{ Workload = 'Fabric'; Expected = '23d8f6bd-1eb0-4cc2-a08c-7bf525c67bcd' }
            @{ Workload = 'Tasks'; Expected = '9ac8c0b3-2c30-497c-b4bc-cadfe9bd6eed' }
            @{ Workload = 'SharePointOnlineREST'; Expected = '31359c7f-bd7e-475c-86db-fdb8c937548e' }
        ) {
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{ Workload = $Workload; Expected = $Expected } {
                param ($Workload, $Expected)
                $instance = New-Object $Workload
                $instance.ApplicationId | Should -Be $Expected
            }
        }
    }

    Context 'Workload CompleteConnection' {
        It 'Should mark the workload as connected and track MFA usage when specified' {
            InModuleScope 'MSCloudLoginAssistant' {
                $instance = New-Object AdminAPI
                $instance.Connected | Should -BeFalse
                $instance.CompleteConnection()
                $instance.Connected | Should -BeTrue
                $instance.ConnectedDateTime | Should -Not -BeNullOrEmpty
                $instance.MultiFactorAuthentication | Should -BeFalse

                $mfaInstance = New-Object AdminAPI
                $mfaInstance.CompleteConnection($true)
                $mfaInstance.MultiFactorAuthentication | Should -BeTrue
            }
        }
    }

    Context 'Workload Clone' {
        It 'Should return a shallow clone of the workload' {
            InModuleScope 'MSCloudLoginAssistant' {
                $instance = New-Object AdminAPI
                $instance.TenantId = 'test-tenant'
                $clone = $instance.Clone()

                $clone.TenantId | Should -Be 'test-tenant'
                $clone | Should -Not -Be $instance
            }
        }
    }
}

# ---------------------------------------------------------------------------
# Connect-M365Tenant dispatching
# ---------------------------------------------------------------------------
Describe 'Connect-M365Tenant' {

    BeforeAll {
        InModuleScope 'MSCloudLoginAssistant' {
            # Ensure a fresh connection profile exists
            $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile

            # Mock all workload Connect functions to prevent real connections
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
            Mock -CommandName Get-CloudEnvironmentInfo -MockWith {
                return @{ tenant_region_sub_scope = $null; token_endpoint = 'https://login.microsoftonline.com/tenant/oauth2/v2.0/token' }
            }
            Mock -CommandName Get-MSCloudLoginCertificate -MockWith {
                $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
                return $cert
            }
            Mock -CommandName Connect-MSCloudLoginAdminAPI -MockWith { }
            Mock -CommandName Connect-MSCloudLoginAzure -MockWith { }
            Mock -CommandName Connect-MSCloudLoginAzureDevOPS -MockWith { }
            Mock -CommandName Connect-MSCloudLoginDefenderForEndpoint -MockWith { }
            Mock -CommandName Connect-MSCloudLoginEngageHub -MockWith { }
            Mock -CommandName Connect-MSCloudLoginExchangeOnline -MockWith { }
            Mock -CommandName Connect-MSCloudLoginFabric -MockWith { }
            Mock -CommandName Connect-MSCloudLoginLicensing -MockWith { }
            Mock -CommandName Connect-MSCloudLoginO365Portal -MockWith { }
            Mock -CommandName Connect-MSCloudLoginMicrosoftGraph -MockWith { }
            Mock -CommandName Connect-MSCloudLoginPnP -MockWith { }
            Mock -CommandName Connect-MSCloudLoginPowerPlatform -MockWith { }
            Mock -CommandName Connect-MSCloudLoginPowerPlatformREST -MockWith { }
            Mock -CommandName Connect-MSCloudLoginSecurityCompliance -MockWith { }
            Mock -CommandName Connect-MSCloudLoginSharePointOnlineREST -MockWith { }
            Mock -CommandName Connect-MSCloudLoginTasks -MockWith { }
            Mock -CommandName Connect-MSCloudLoginTeams -MockWith { }
            Mock -CommandName Get-ConnectionInformation -MockWith { return $null }
        }
    }

    Context 'When connecting to a workload' {
        It 'Should store the authentication parameters on the <ProfileName> profile and invoke <ConnectFunction> for <Workload>' -TestCases @(
            @{ Workload = 'AdminAPI'; ProfileName = 'AdminAPI'; ConnectFunction = 'Connect-MSCloudLoginAdminAPI'; AuthParameter = 'CertificateThumbprint'; AuthValue = 'thumb' }
            @{ Workload = 'ExchangeOnline'; ProfileName = 'ExchangeOnline'; ConnectFunction = 'Connect-MSCloudLoginExchangeOnline'; AuthParameter = 'CertificateThumbprint'; AuthValue = 'thumb' }
            @{ Workload = 'MicrosoftGraph'; ProfileName = 'MicrosoftGraph'; ConnectFunction = 'Connect-MSCloudLoginMicrosoftGraph'; AuthParameter = 'ApplicationSecret'; AuthValue = 'secret' }
            @{ Workload = 'MicrosoftTeams'; ProfileName = 'Teams'; ConnectFunction = 'Connect-MSCloudLoginTeams'; AuthParameter = 'CertificateThumbprint'; AuthValue = 'thumb' }
            @{ Workload = 'PowerPlatforms'; ProfileName = 'PowerPlatform'; ConnectFunction = 'Connect-MSCloudLoginPowerPlatform'; AuthParameter = 'CertificateThumbprint'; AuthValue = 'thumb' }
            @{ Workload = 'SecurityComplianceCenter'; ProfileName = 'SecurityComplianceCenter'; ConnectFunction = 'Connect-MSCloudLoginSecurityCompliance'; AuthParameter = 'CertificateThumbprint'; AuthValue = 'thumb' }
            @{ Workload = 'Tasks'; ProfileName = 'Tasks'; ConnectFunction = 'Connect-MSCloudLoginTasks'; AuthParameter = 'ApplicationSecret'; AuthValue = 'secret' }
        ) {
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{
                Workload        = $Workload
                ProfileName     = $ProfileName
                ConnectFunction = $ConnectFunction
                AuthParameter   = $AuthParameter
                AuthValue       = $AuthValue
            } {
                param ($Workload, $ProfileName, $ConnectFunction, $AuthParameter, $AuthValue)
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $connectParams = @{
                    Workload      = $Workload
                    ApplicationId = 'app-id'
                    TenantId      = 'tenant-id'
                    $AuthParameter = $AuthValue
                }
                Connect-M365Tenant @connectParams

                Should -Invoke -CommandName $ConnectFunction -Exactly 1
                $Script:MSCloudLoginConnectionProfile.$ProfileName.ApplicationId  | Should -Be 'app-id'
                $Script:MSCloudLoginConnectionProfile.$ProfileName.TenantId       | Should -Be 'tenant-id'
                $Script:MSCloudLoginConnectionProfile.$ProfileName.$AuthParameter | Should -Be $AuthValue
            }
        }
    }

    Context 'When connecting to Azure with a SubscriptionId' {
        It 'Should reconnect only on SubscriptionId drift when <Scenario>' -TestCases @(
            @{ Scenario = 'the same SubscriptionId is provided again'; Sequence = @('sub-A', 'sub-A'); IdentityOnFirstCallOnly = $false; ExpectedStates = @($false, $true); ExpectedSubscriptionId = 'sub-A' }
            @{ Scenario = 'a different SubscriptionId is provided'; Sequence = @('sub-A', 'sub-B'); IdentityOnFirstCallOnly = $false; ExpectedStates = @($false, $false); ExpectedSubscriptionId = 'sub-B' }
            @{ Scenario = 'the SubscriptionId is omitted and the profile clears it'; Sequence = @('sub-A', $null, $null); IdentityOnFirstCallOnly = $false; ExpectedStates = @($false, $false, $true); ExpectedSubscriptionId = '' }
            @{ Scenario = 'the SubscriptionId is set again after being omitted'; Sequence = @('sub-A', $null, 'sub-B'); IdentityOnFirstCallOnly = $false; ExpectedStates = @($false, $false, $false); ExpectedSubscriptionId = 'sub-B' }
            @{ Scenario = 'no identity parameter is repeated'; Sequence = @('sub-A', 'sub-A', 'sub-B'); IdentityOnFirstCallOnly = $true; ExpectedStates = @($false, $true, $false); ExpectedSubscriptionId = 'sub-B' }
        ) {
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{
                Sequence                = $Sequence
                IdentityOnFirstCallOnly = $IdentityOnFirstCallOnly
                ExpectedStates          = $ExpectedStates
                ExpectedSubscriptionId  = $ExpectedSubscriptionId
            } {
                param ($Sequence, $IdentityOnFirstCallOnly, $ExpectedStates, $ExpectedSubscriptionId)
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:AzureConnectedStates = [System.Collections.Generic.List[bool]]::new()
                Mock -CommandName Connect-MSCloudLoginAzure -MockWith {
                    $Script:AzureConnectedStates.Add($Script:MSCloudLoginConnectionProfile.Azure.Connected)
                    $Script:MSCloudLoginConnectionProfile.Azure.CompleteConnection()
                }

                for ($i = 0; $i -lt $Sequence.Count; $i++)
                {
                    $connectParams = @{ Workload = 'Azure' }
                    if ($i -eq 0 -or -not $IdentityOnFirstCallOnly)
                    {
                        $connectParams += @{ ApplicationId = 'app-id'; TenantId = 'tenant-id'; ApplicationSecret = 'secret' }
                    }
                    if ($Sequence[$i])
                    {
                        $connectParams.SubscriptionId = $Sequence[$i]
                    }
                    Connect-M365Tenant @connectParams
                }

                $Script:AzureConnectedStates | Should -Be $ExpectedStates
                [System.String]$Script:MSCloudLoginConnectionProfile.Azure.SubscriptionId | Should -Be $ExpectedSubscriptionId
            }
        }
    }

    Context 'When connecting to ExchangeOnline with cmdlets to load' {
        It 'Should map ExchangeOnlineCmdlets onto CmdletsToLoad and treat omitted cmdlets as drift' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:ExoConnectedStates = [System.Collections.Generic.List[bool]]::new()
                Mock -CommandName Connect-MSCloudLoginExchangeOnline -MockWith {
                    $Script:ExoConnectedStates.Add($Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected)
                    $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CompleteConnection()
                }

                Connect-M365Tenant -Workload 'ExchangeOnline' -ApplicationId 'app-id' -TenantId 'tenant-id' -ApplicationSecret 'secret' -ExchangeOnlineCmdlets @('Get-Mailbox')
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CmdletsToLoad | Should -Be @('Get-Mailbox')

                Connect-M365Tenant -Workload 'ExchangeOnline' -ApplicationId 'app-id' -TenantId 'tenant-id' -ApplicationSecret 'secret'
                Connect-M365Tenant -Workload 'ExchangeOnline' -ApplicationId 'app-id' -TenantId 'tenant-id' -ApplicationSecret 'secret'

                $Script:ExoConnectedStates | Should -Be @($false, $false, $true)
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CmdletsToLoad | Should -BeNullOrEmpty
            }
        }
    }
}

# ---------------------------------------------------------------------------
# Compare-InputParametersForChange
# ---------------------------------------------------------------------------
Describe 'Compare-InputParametersForChange' {

    BeforeAll {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
        }
    }

    Context 'When no prior connection profile exists' {
        It 'Should return true to force a reconnect from the corrupted state' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = $null
                $params = @{ Workload = 'AdminAPI' }
                $result = Compare-InputParametersForChange -CurrentParamSet $params
                $result | Should -BeTrue
            }
        }
    }

    Context 'When comparing against the stored workload profile' {
        It 'Should return <Expected> when <Scenario>' -TestCases @(
            @{
                Scenario      = 'the authentication type changes'
                ProfileName   = 'AdminAPI'
                ProfileValues = @{ AuthenticationType = 'Credentials'; RequestedAuthenticationType = 'ServicePrincipalWithThumbprint' }
                Params        = @{ Workload = 'AdminAPI'; ApplicationId = 'app-id'; TenantId = 'tenant-id' }
                Expected      = $true
            }
            @{
                Scenario      = 'the parameters have not changed'
                ProfileName   = 'AdminAPI'
                ProfileValues = @{ AuthenticationType = 'ServicePrincipalWithThumbprint'; RequestedAuthenticationType = 'ServicePrincipalWithThumbprint'; ApplicationId = 'app-id'; TenantId = 'tenant-id'; CertificateThumbprint = 'thumb' }
                Params        = @{ Workload = 'AdminAPI'; ApplicationId = 'app-id'; TenantId = 'tenant-id'; CertificateThumbprint = 'thumb' }
                Expected      = $false
            }
            @{
                Scenario      = 'two parameter values are swapped'
                ProfileName   = 'AdminAPI'
                ProfileValues = @{ AuthenticationType = 'ServicePrincipalWithSecret'; RequestedAuthenticationType = 'ServicePrincipalWithSecret'; ApplicationId = 'value-A'; TenantId = 'value-B'; ApplicationSecret = 'secret' }
                Params        = @{ Workload = 'AdminAPI'; ApplicationId = 'value-B'; TenantId = 'value-A'; ApplicationSecret = 'secret' }
                Expected      = $true
            }
            @{
                Scenario      = 'a SubscriptionId is newly provided for Azure'
                ProfileName   = 'Azure'
                ProfileValues = @{ AuthenticationType = 'ServicePrincipalWithSecret'; RequestedAuthenticationType = 'ServicePrincipalWithSecret'; ApplicationId = 'app-id'; TenantId = 'tenant-id'; ApplicationSecret = 'secret' }
                Params        = @{ Workload = 'Azure'; ApplicationId = 'app-id'; TenantId = 'tenant-id'; ApplicationSecret = 'secret'; SubscriptionId = 'sub-B' }
                Expected      = $true
            }
            @{
                Scenario      = 'the SubscriptionId is omitted'
                ProfileName   = 'Azure'
                ProfileValues = @{ AuthenticationType = 'ServicePrincipalWithSecret'; RequestedAuthenticationType = 'ServicePrincipalWithSecret'; ApplicationId = 'app-id'; TenantId = 'tenant-id'; ApplicationSecret = 'secret'; SubscriptionId = 'sub-A' }
                Params        = @{ Workload = 'Azure'; ApplicationId = 'app-id'; TenantId = 'tenant-id'; ApplicationSecret = 'secret' }
                Expected      = $true
            }
            @{
                Scenario      = 'the cmdlets to load are omitted'
                ProfileName   = 'ExchangeOnline'
                ProfileValues = @{ AuthenticationType = 'ServicePrincipalWithSecret'; RequestedAuthenticationType = 'ServicePrincipalWithSecret'; ApplicationId = 'app-id'; TenantId = 'tenant-id'; ApplicationSecret = 'secret'; CmdletsToLoad = @('Get-Mailbox') }
                Params        = @{ Workload = 'ExchangeOnline'; ApplicationId = 'app-id'; TenantId = 'tenant-id'; ApplicationSecret = 'secret' }
                Expected      = $true
            }
            # EnableSearchOnlySession only exists on SecurityComplianceCenter. Supplying it
            # for another workload must not be reported as a change on every call.
            @{
                Scenario      = 'a session parameter that the workload does not own is supplied'
                ProfileName   = 'MicrosoftGraph'
                ProfileValues = @{ AuthenticationType = 'ServicePrincipalWithSecret'; RequestedAuthenticationType = 'ServicePrincipalWithSecret'; ApplicationId = 'app-id'; TenantId = 'tenant-id'; ApplicationSecret = 'secret' }
                Params        = @{ Workload = 'MicrosoftGraph'; ApplicationId = 'app-id'; TenantId = 'tenant-id'; ApplicationSecret = 'secret'; EnableSearchOnlySession = $true }
                Expected      = $false
            }
            @{
                Scenario      = 'only the case of an identifier changes'
                ProfileName   = 'AdminAPI'
                ProfileValues = @{ AuthenticationType = 'ServicePrincipalWithSecret'; RequestedAuthenticationType = 'ServicePrincipalWithSecret'; ApplicationId = 'app-id'; TenantId = 'Tenant-Id'; ApplicationSecret = 'secret' }
                Params        = @{ Workload = 'AdminAPI'; ApplicationId = 'APP-ID'; TenantId = 'tenant-id'; ApplicationSecret = 'secret' }
                Expected      = $false
            }
            @{
                Scenario      = 'only the case of a secret changes'
                ProfileName   = 'AdminAPI'
                ProfileValues = @{ AuthenticationType = 'ServicePrincipalWithSecret'; RequestedAuthenticationType = 'ServicePrincipalWithSecret'; ApplicationId = 'app-id'; TenantId = 'tenant-id'; ApplicationSecret = 'Secret' }
                Params        = @{ Workload = 'AdminAPI'; ApplicationId = 'app-id'; TenantId = 'tenant-id'; ApplicationSecret = 'secret' }
                Expected      = $true
            }
            @{
                Scenario      = 'an empty string on the profile meets an absent parameter'
                ProfileName   = 'AdminAPI'
                ProfileValues = @{ AuthenticationType = 'ServicePrincipalWithSecret'; RequestedAuthenticationType = 'ServicePrincipalWithSecret'; ApplicationId = 'app-id'; TenantId = ''; ApplicationSecret = 'secret' }
                Params        = @{ Workload = 'AdminAPI'; ApplicationId = 'app-id'; ApplicationSecret = 'secret' }
                Expected      = $false
            }
            @{
                Scenario      = 'the MicrosoftTeams workload is compared against the Teams profile and the token it acquired itself'
                ProfileName   = 'Teams'
                ProfileValues = @{ AuthenticationType = 'ServicePrincipalWithThumbprint'; RequestedAuthenticationType = 'ServicePrincipalWithThumbprint'; ApplicationId = 'app-id'; TenantId = 'tenant-id'; CertificateThumbprint = 'thumb'; AccessTokens = @('acquired-token') }
                Params        = @{ Workload = 'MicrosoftTeams'; ApplicationId = 'app-id'; TenantId = 'tenant-id'; CertificateThumbprint = 'thumb' }
                Expected      = $false
            }
            @{
                Scenario      = 'a Graph credential connection holds the default application id, the UPN tenant and the token it acquired itself'
                ProfileName   = 'MicrosoftGraph'
                ProfileValues = @{ AuthenticationType = 'Credentials'; RequestedAuthenticationType = 'Credentials'; Credentials = (New-Object PSCredential ('user@contoso.com', (ConvertTo-SecureString 'pwd' -AsPlainText -Force))); ApplicationId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'; TenantId = 'contoso.com'; AccessTokens = @('acquired-token') }
                Params        = @{ Workload = 'MicrosoftGraph'; Credential = (New-Object PSCredential ('user@contoso.com', (ConvertTo-SecureString 'pwd' -AsPlainText -Force))) }
                Expected      = $false
            }
            @{
                Scenario      = 'the caller passes a different access token'
                ProfileName   = 'MicrosoftGraph'
                ProfileValues = @{ AuthenticationType = 'AccessTokens'; RequestedAuthenticationType = 'AccessTokens'; TenantId = 'contoso.onmicrosoft.com'; AccessTokens = @('first-token') }
                Params        = @{ Workload = 'MicrosoftGraph'; TenantId = 'contoso.onmicrosoft.com'; AccessTokens = @('second-token') }
                Expected      = $true
            }
            @{
                Scenario      = 'the PowerPlatforms workload is compared against the PowerPlatform profile'
                ProfileName   = 'PowerPlatform'
                ProfileValues = @{ AuthenticationType = 'ServicePrincipalWithThumbprint'; RequestedAuthenticationType = 'ServicePrincipalWithThumbprint'; ApplicationId = 'app-id'; TenantId = 'tenant-id'; CertificateThumbprint = 'thumb' }
                Params        = @{ Workload = 'PowerPlatforms'; ApplicationId = 'app-id'; TenantId = 'tenant-id'; CertificateThumbprint = 'thumb' }
                Expected      = $false
            }
        ) {
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{ ProfileName = $ProfileName; ProfileValues = $ProfileValues; Params = $Params; Expected = $Expected } {
                param ($ProfileName, $ProfileValues, $Params, $Expected)
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $workloadProfile = $Script:MSCloudLoginConnectionProfile.$ProfileName
                foreach ($key in $ProfileValues.Keys)
                {
                    $workloadProfile.$key = $ProfileValues[$key]
                }
                (Compare-InputParametersForChange -CurrentParamSet $Params) | Should -Be $Expected
            }
        }
    }

    Context 'When the credential is compared' {
        It 'Should return <Expected> for <Scenario> without mutating the parameter set' -TestCases @(
            @{ Scenario = 'a changed password of the same user name'; UserName = 'user@contoso.com'; Password = 'NewPwd'; Expected = $true }
            @{ Scenario = 'an identical credential with a differently cased user name'; UserName = 'USER@contoso.com'; Password = 'SamePwd'; Expected = $false }
        ) {
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{ UserName = $UserName; Password = $Password; Expected = $Expected } {
                param ($UserName, $Password, $Expected)
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $storedCred = New-Object PSCredential ('user@contoso.com', (ConvertTo-SecureString 'SamePwd' -AsPlainText -Force))
                $newCred = New-Object PSCredential ($UserName, (ConvertTo-SecureString $Password -AsPlainText -Force))
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.AuthenticationType          = 'Credentials'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.RequestedAuthenticationType = 'Credentials'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Credentials                 = $storedCred

                $params = @{
                    Workload   = 'ExchangeOnline'
                    Credential = $newCred
                }
                (Compare-InputParametersForChange -CurrentParamSet $params) | Should -Be $Expected
                $params.Keys.Count | Should -Be 2
                $params.ContainsKey('Credential') | Should -BeTrue
                $params.ContainsKey('Workload') | Should -BeTrue
            }
        }
    }

    Context 'Secret values are never logged' {
        It 'Should log key names only when a secret changed' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.AdminAPI.AuthenticationType          = 'ServicePrincipalWithSecret'
                $Script:MSCloudLoginConnectionProfile.AdminAPI.RequestedAuthenticationType = 'ServicePrincipalWithSecret'
                $Script:MSCloudLoginConnectionProfile.AdminAPI.ApplicationId               = 'app-id'
                $Script:MSCloudLoginConnectionProfile.AdminAPI.TenantId                    = 'tenant-id'
                $Script:MSCloudLoginConnectionProfile.AdminAPI.ApplicationSecret           = 'super-secret-old'

                $params = @{
                    Workload          = 'AdminAPI'
                    ApplicationId     = 'app-id'
                    TenantId          = 'tenant-id'
                    ApplicationSecret = 'super-secret-new'
                }
                (Compare-InputParametersForChange -CurrentParamSet $params) | Should -BeTrue
                Should -Invoke Add-MSCloudLoginAssistantEvent -Exactly 1 -ParameterFilter {
                    $Message -like '*changed for workload {AdminAPI}: ApplicationSecret'
                }
                Should -Invoke Add-MSCloudLoginAssistantEvent -Exactly 0 -ParameterFilter {
                    $Message -like '*super-secret*'
                }
            }
        }
    }
}

# ---------------------------------------------------------------------------
# Reset-MSCloudLoginConnectionProfileContext
# ---------------------------------------------------------------------------
Describe 'Reset-MSCloudLoginConnectionProfileContext' {

    BeforeAll {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

            # Mock all disconnect functions
            Mock -CommandName Disconnect-MSCloudLoginAdminAPI -MockWith { }
            Mock -CommandName Disconnect-MSCloudLoginAzure -MockWith { }
            Mock -CommandName Disconnect-MSCloudLoginAzureDevOPS -MockWith { }
            Mock -CommandName Disconnect-MSCloudLoginDefenderForEndpoint -MockWith { }
            Mock -CommandName Disconnect-MSCloudLoginEngageHub -MockWith { }
            Mock -CommandName Disconnect-ExchangeOnline -MockWith { }
            Mock -CommandName Disconnect-MSCloudLoginFabric -MockWith { }
            Mock -CommandName Disconnect-MSCloudLoginLicensing -MockWith { }
            Mock -CommandName Disconnect-MSCloudLoginO365Portal -MockWith { }
            Mock -CommandName Disconnect-MSCloudLoginMicrosoftGraph -MockWith { }
            Mock -CommandName Disconnect-MSCloudLoginPnP -MockWith { }
            Mock -CommandName Disconnect-MSCloudLoginPowerPlatformREST -MockWith { }
            Mock -CommandName Disconnect-MSCloudLoginSecurityCompliance -MockWith { }
            Mock -CommandName Disconnect-MSCloudLoginSharePointOnlineREST -MockWith { }
            Mock -CommandName Disconnect-MSCloudLoginTasks -MockWith { }
            Mock -CommandName Disconnect-MSCloudLoginTeams -MockWith { }
        }
    }

    Context 'When resetting a specific workload' {
        It 'Should call Disconnect on the specified workload' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                Reset-MSCloudLoginConnectionProfileContext -Workload 'AdminAPI'
                Should -Invoke Disconnect-MSCloudLoginAdminAPI -Exactly 1
            }
        }
    }

    Context 'When resetting all workloads' {
        It 'Should recreate the connection profile' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $originalTime = $Script:MSCloudLoginConnectionProfile.CreatedTime

                # Small delay to ensure timestamp changes
                Start-Sleep -Seconds 1
                Reset-MSCloudLoginConnectionProfileContext

                $Script:MSCloudLoginConnectionProfile.CreatedTime | Should -Not -Be $originalTime
            }
        }
    }

    Context 'When a workload has no Disconnect method' {
        It 'Should log that the operation was ignored' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Get-Member -MockWith { return $null } -ParameterFilter {
                    $Name -eq 'Disconnect'
                }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                Reset-MSCloudLoginConnectionProfileContext -Workload 'AdminAPI'

                Should -Invoke Add-MSCloudLoginAssistantEvent -ParameterFilter {
                    $Message -like 'No disconnect method found for workload {AdminAPI}*'
                }
            }
        }
    }
}

# ---------------------------------------------------------------------------
# Get-MSCloudLoginConnectionProfile
# ---------------------------------------------------------------------------
Describe 'Get-MSCloudLoginConnectionProfile' {

    BeforeAll {
        InModuleScope 'MSCloudLoginAssistant' {
            $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
            $Script:MSCloudLoginConnectionProfile.AdminAPI.TenantId = 'test-tenant-profile'
        }
    }

    Context 'When requesting an existing workload profile' {
        It 'Should return a clone of the workload profile' {
            $result = Get-MSCloudLoginConnectionProfile -Workload 'AdminAPI'
            $result | Should -Not -BeNullOrEmpty
            $result.TenantId | Should -Be 'test-tenant-profile'
        }
    }
}

# ---------------------------------------------------------------------------
# Connect-MSCloudLoginSecurityCompliance
# ---------------------------------------------------------------------------
Describe 'Connect-MSCloudLoginSecurityCompliance' {

    BeforeAll {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
            Mock -CommandName Remove-MSCloudLoginProxyModule -MockWith { }
            Mock -CommandName Get-Module -MockWith { return @() }
            Mock -CommandName Get-PSSession -MockWith { return @() }
            Mock -CommandName Connect-IPPSSession -MockWith { }
        }
    }

    Context 'When the authentication type is credential based' {
        It 'Should connect for both Credentials and CredentialsWithApplicationId' -TestCases @(
            @{ AuthenticationType = 'Credentials' }
            @{ AuthenticationType = 'CredentialsWithApplicationId' }
        ) {
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{ AuthenticationType = $AuthenticationType } {
                param($AuthenticationType)

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $profile = $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter
                $profile.AuthenticationType = $AuthenticationType
                $profile.Credentials = [System.Management.Automation.PSCredential]::new(
                    'admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))
                $profile.ConnectionUrl = 'https://ps.compliance.protection.outlook.com/powershell-liveid/'
                $profile.AzureADAuthorizationEndpointUri = 'https://login.microsoftonline.com/organizations'

                { Connect-MSCloudLoginSecurityCompliance } | Should -Not -Throw
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.Connected | Should -BeTrue
                Should -Invoke Connect-IPPSSession -Exactly 1 -ParameterFilter {
                    $Credential.UserName -eq 'admin@contoso.onmicrosoft.com'
                }
            }
        }
    }

    Context 'When the authentication type is not supported' {
        It 'Should throw' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter.AuthenticationType = 'Interactive'

                { Connect-MSCloudLoginSecurityCompliance } | Should -Throw "*is not supported for workload 'SecurityComplianceCenter'*"
            }
        }
    }
}

# ---------------------------------------------------------------------------
# Connect-M365Tenant PnP URL handling
# ---------------------------------------------------------------------------
Describe 'Connect-M365Tenant PnP URL handling' {

    BeforeAll {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
            Mock -CommandName Import-Module -MockWith { }
            Mock -CommandName Get-MSCloudLoginCertificate -MockWith {
                return New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
            }
            Mock -CommandName Get-CloudEnvironmentInfo -MockWith {
                return @{ tenant_region_sub_scope = $null; token_endpoint = 'https://login.microsoftonline.com/t/oauth2/v2.0/token' }
            }
            Mock -CommandName Connect-MSCloudLoginPnP -MockWith { }
            Mock -CommandName Connect-PnPOnline -MockWith { }
            Mock -CommandName Connect-MgGraph -MockWith { }
            Mock -CommandName Get-Module -MockWith {
                if ($ListAvailable.IsPresent -and $Name -eq 'PnP.PowerShell') {
                    return @([PSCustomObject]@{ Name = 'PnP.PowerShell'; Version = [System.Version]'1.10.0'; CompatiblePSEditions = @('Desktop', 'Core') })
                }
                return @()
            }
            Mock -CommandName Get-MSCloudLoginSPOUrlFromTenantId -MockWith {
                return @{ ConnectionUrl = 'https://contoso.sharepoint.com'; AdminUrl = 'https://contoso-admin.sharepoint.com' }
            }
        }
    }

    Context 'When connecting to PnP for the first time with a URL' {
        It 'Should set the AdminUrl from the provided URL' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                Connect-M365Tenant -Workload 'PnP' `
                    -Url 'https://contoso-admin.sharepoint.com' `
                    -ApplicationId 'app-id' -TenantId 'tenant-id' `
                    -CertificateThumbprint 'thumb'

                $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl | Should -Be 'https://contoso-admin.sharepoint.com'
            }
        }
    }

    Context 'When connecting to PnP without a URL after AdminUrl is set' {
        It 'Should use AdminUrl and handle context mismatch' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Get-PnPContext -MockWith {
                    return @{ Url = 'https://contoso.sharepoint.com/sites/marketing' }
                }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl = 'https://contoso-admin.sharepoint.com'
                $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl = 'https://contoso-admin.sharepoint.com'
                $Script:MSCloudLoginConnectionProfile.PnP.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.PnP.TenantId = 'tenant-id'
                $Script:MSCloudLoginConnectionProfile.PnP.CertificateThumbprint = 'thumb'
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'ServicePrincipalWithThumbprint'
                $Script:MSCloudLoginConnectionProfile.PnP.RequestedAuthenticationType = 'ServicePrincipalWithThumbprint'
                $Script:MSCloudLoginConnectionProfile.PnP.Connected = $true
                $Script:MSCloudLoginConnectionProfile.PnP.ConnectedDateTime = [System.DateTime]::Now.ToString()

                Connect-M365Tenant -Workload 'PnP' `
                    -ApplicationId 'app-id' -TenantId 'tenant-id' `
                    -CertificateThumbprint 'thumb'

                # ConnectionUrl should be back to AdminUrl after context mismatch
                $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl | Should -Be 'https://contoso-admin.sharepoint.com'
            }
        }
    }
}

# ---------------------------------------------------------------------------
# Custom environment reload in Connect-M365Tenant
# ---------------------------------------------------------------------------
Describe 'Connect-M365Tenant custom environment reload' {

    BeforeAll {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
            Mock -CommandName Connect-MSCloudLoginAdminAPI -MockWith { }
            Mock -CommandName Import-PowerShellDataFile -MockWith {
                return @{ CustomEnvironment = $false }
            }
        }
    }

    Context 'When the custom environment file name changes' {
        It 'Should reload the custom environment configuration' {
            InModuleScope 'MSCloudLoginAssistant' {
                # Clear the module-level config to force a reload
                $Script:CustomEnvConfig = $null

                Connect-M365Tenant -Workload 'AdminAPI' `
                    -ApplicationId 'app-id' -TenantId 'tenant-id' -ApplicationSecret 'secret' `
                    -CustomEnvironmentFileName 'OtherEnvironment.psd1'

                $Script:LoadedCustomEnvFileName | Should -Be 'OtherEnvironment.psd1'
                Should -Invoke Import-PowerShellDataFile -ParameterFilter {
                    $Path -like '*OtherEnvironment.psd1'
                }
            }
        }
    }
}
