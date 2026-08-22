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
        $manifestPath = Join-Path $tempModuleBase "$graphModuleName.psd1"
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

Describe 'PnP constructor validation' {

    It 'Should refuse a certificate thumbprint combined with a certificate path' {
        InModuleScope 'MSCloudLoginAssistant' {
            # Property initializers run before the base constructor body, which lets us
            # instantiate the class with both certificate options already populated.
            # The class keyword resolves base types at parse time, so the derived class
            # must be compiled at runtime, after the module types became available.
            $derivedClassDefinition = @'
class PnPWithConflictingCertificate : PnP
{
    [string]$CertificateThumbprint = 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD'
    [string]$CertificatePath = "$env:TEMP\contoso.pfx"
}
'@
            . ([scriptblock]::Create($derivedClassDefinition))

            { [PnPWithConflictingCertificate]::new() } |
                Should -Throw '*Cannot specify both a Certificate Thumbprint and Certificate Path and Password*'
        }
    }

    It 'Should allow an instance without conflicting certificate settings' {
        InModuleScope 'MSCloudLoginAssistant' {
            { [PnP]::new() } | Should -Not -Throw
        }
    }
}

Describe 'PnP sovereign cloud environment translation' {

    BeforeEach {
        InModuleScope 'MSCloudLoginAssistant' {
            $Script:MSCloudLoginConnectionProfile = $null
            $Script:MSCloudLoginTriedGetEnvironment = $true
            $Script:CloudEnvironmentInfo = $null
        }
    }

    It 'Should translate <EnvironmentName> into the <ExpectedPnPEnvironment> PnP environment' -TestCases @(
        @{ EnvironmentName = 'AzureFranceCloud'; ExpectedPnPEnvironment = 'BleuCloud'; Discovery = '{ "tenant_region_scope": "FG", "token_endpoint": "https://login.sovcloud-identity.fr/t/oauth2/v2.0/token" }' }
        @{ EnvironmentName = 'AzureGermanyCloud'; ExpectedPnPEnvironment = 'DelosCloud'; Discovery = '{ "tenant_region_scope": "GG2", "token_endpoint": "https://login.sovcloud-identity.de/t/oauth2/v2.0/token" }' }
    ) {
        param ($EnvironmentName, $ExpectedPnPEnvironment, $Discovery)
        InModuleScope 'MSCloudLoginAssistant' -Parameters @{
            EnvironmentName        = $EnvironmentName
            ExpectedPnPEnvironment = $ExpectedPnPEnvironment
            Discovery              = $Discovery
        } {
            param ($EnvironmentName, $ExpectedPnPEnvironment, $Discovery)

            Mock -CommandName Import-Module -MockWith { }
            Mock -CommandName Connect-PnPOnline -MockWith { }
            Mock -CommandName Get-PnPContext -MockWith { throw 'no context yet' }

            $Script:CloudEnvironmentInfo = ConvertFrom-Json $Discovery

            Connect-M365Tenant -Workload 'PnP' `
                -Url 'https://contoso-admin.sharepoint.com' `
                -ApplicationId '11111111-1111-1111-1111-111111111111' `
                -TenantId 'contoso.onmicrosoft.com' `
                -CertificateThumbprint 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD'

            $connection = Get-MSCloudLoginConnectionProfile -Workload 'PnP'
            $connection.EnvironmentName | Should -Be $EnvironmentName
            $connection.PnPAzureEnvironment | Should -Be $ExpectedPnPEnvironment
            Should -Invoke Connect-PnPOnline -Exactly 1 -ParameterFilter {
                $AzureEnvironment -eq $ExpectedPnPEnvironment
            }
        }
    }
}

Describe 'MicrosoftTeams sovereign cloud environment registration' {

    BeforeEach {
        InModuleScope 'MSCloudLoginAssistant' {
            $Script:MSCloudLoginConnectionProfile = $null
            $Script:MSCloudLoginTriedGetEnvironment = $true
            $Script:CloudEnvironmentInfo = $null
        }
    }

    AfterEach {
        InModuleScope 'MSCloudLoginAssistant' -Parameters @{ ModuleRoot = $script:moduleRoot } {
            param ($ModuleRoot)

            # Teams.Connect() writes into the shared custom environment configuration.
            $Script:CustomEnvConfig = Import-PowerShellDataFile -Path (Join-Path $ModuleRoot 'CustomEnvironment.psd1')
            $Script:LoadedCustomEnvFileName = 'CustomEnvironment.psd1'
        }
    }

    It 'Should register the German sovereign endpoints and refuse the connection outside of Windows PowerShell 5' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Get-CsTeamsCallingPolicy -MockWith { throw 'no session' }
            Mock -CommandName Connect-MicrosoftTeams -MockWith { }
            Mock -CommandName Set-TeamsEnvironmentConfig -MockWith { }

            $Script:CloudEnvironmentInfo = ConvertFrom-Json '{ "tenant_region_scope": "GG2", "token_endpoint": "https://login.sovcloud-identity.de/t/oauth2/v2.0/token" }'

            { Connect-M365Tenant -Workload 'MicrosoftTeams' `
                -ApplicationId '11111111-1111-1111-1111-111111111111' `
                -TenantId 'contoso.onsovcloud.de' `
                -CertificateThumbprint 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD' } |
                Should -Throw '*only supported in PowerShell 5*'

            $connection = Get-MSCloudLoginConnectionProfile -Workload 'MicrosoftTeams'
            $connection.EnvironmentName | Should -Be 'AzureGermanyCloud'
            $connection.AuthorizationUrl | Should -Be 'https://login.sovcloud-identity.de/'
            $Script:CustomEnvConfig.CustomTeamsEndpoints.ActiveDirectory | Should -Be 'https://login.sovcloud-identity.de/'
            $Script:CustomEnvConfig.CustomTeamsEndpoints.MsGraphEndpointResourceId | Should -Be 'https://graph.svc.sovcloud.de'
            $Script:CustomEnvConfig.CustomTeamsEndpoints.TeamsConfigApiEndPoint | Should -Be 'https://config.teams.sovcloud.de'
            Should -Invoke Set-TeamsEnvironmentConfig -Exactly 0
        }
    }
}

Describe 'SharePointOnlineREST connection logic' {

    BeforeEach {
        InModuleScope 'MSCloudLoginAssistant' {
            $Script:MSCloudLoginTriedGetEnvironment = $true
            $Script:CloudEnvironmentInfo = $null
            $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile

            # Configure the stored profile so that Connect() can discover the admin URL,
            # then invoke Connect() on a separate instance to observe how the discovered
            # URL is adopted from the profile.
            $storedProfile = $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST
            $storedProfile.AuthenticationType = 'ServicePrincipalWithSecret'
            $storedProfile.RequestedAuthenticationType = 'ServicePrincipalWithSecret'
            $storedProfile.ApplicationId = '11111111-1111-1111-1111-111111111111'
            $storedProfile.ApplicationSecret = 'super-secret'
            $storedProfile.TenantId = 'contoso.onmicrosoft.com'
        }
    }

    AfterEach {
        InModuleScope 'MSCloudLoginAssistant' -Parameters @{ ModuleRoot = $script:moduleRoot } {
            param ($ModuleRoot)
            $Script:CustomEnvConfig = Import-PowerShellDataFile -Path (Join-Path $ModuleRoot 'CustomEnvironment.psd1')
            $Script:LoadedCustomEnvFileName = 'CustomEnvironment.psd1'
        }
    }

    It 'Should adopt the admin URL that was discovered through the tenant id' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Connect-MSCloudLoginSharePointOnlineREST -MockWith { }

            $workload = [SharePointOnlineREST]::new()
            $workload.RequestedAuthenticationType = 'ServicePrincipalWithSecret'
            $workload.Connect()

            $workload.AdminUrl | Should -Be 'https://contoso-admin.sharepoint.com'
            $workload.HostUrl | Should -Be 'https://contoso-admin.sharepoint.com'
            $workload.Scope | Should -Be 'https://contoso-admin.sharepoint.com/.default'
            $workload.AuthorizationUrl | Should -Be 'https://login.microsoftonline.com'
            Should -Invoke Connect-MSCloudLoginSharePointOnlineREST -Exactly 1
        }
    }

    It 'Should derive the scope from the custom host URL' {
        InModuleScope 'MSCloudLoginAssistant' {
            $Script:CustomEnvConfig.CustomEnvironment = $true
            Mock -CommandName Connect-MSCloudLoginSharePointOnlineREST -MockWith { }

            $workload = [SharePointOnlineREST]::new()
            $workload.RequestedAuthenticationType = 'ServicePrincipalWithSecret'
            $workload.Connect()

            $workload.EnvironmentName | Should -Be 'Custom'
            $workload.AdminUrl | Should -Be 'https://contoso-admin.sharepoint.com'
            $workload.HostUrl | Should -Be 'https://customdomain.sharepoint.com'
            $workload.AuthorizationUrl | Should -Be 'https://login.microsoftonline.com'
            # The custom configuration defines no dedicated scope key, so the class
            # derives the scope from the resolved host URL.
            $workload.Scope | Should -Be 'https://customdomain.sharepoint.com/.default'
            Should -Invoke Connect-MSCloudLoginSharePointOnlineREST -Exactly 1
        }
    }
}
