#Requires -Modules Pester

Describe 'Connect-MSCloudLoginPowerPlatformREST' {
    BeforeAll {
        Import-Module ./Modules/MSCloudLoginAssistant/MSCloudLoginAssistant.psd1 -Force
    }

    Context 'When connecting to PowerPlatformREST' {
        It 'Should run <ExpectedProbes> token probe(s), reset a rejected token and connect when <Scenario>' -TestCases @(
            @{ Scenario = 'no token is cached'; Connected = $false; AccessToken = $null; ProbeFails = $false; ExpectedProbes = 0; ExpectedConnected = $false; ExpectedAccessToken = $null }
            @{ Scenario = 'the token probe succeeds'; Connected = $true; AccessToken = 'Bearer token123'; ProbeFails = $false; ExpectedProbes = 1; ExpectedConnected = $true; ExpectedAccessToken = 'Bearer token123' }
            @{ Scenario = 'the token probe fails'; Connected = $true; AccessToken = 'Bearer token123'; ProbeFails = $true; ExpectedProbes = 1; ExpectedConnected = $false; ExpectedAccessToken = $null }
        ) {
            param ($Scenario, $Connected, $AccessToken, $ProbeFails, $ExpectedProbes, $ExpectedConnected, $ExpectedAccessToken)
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{
                Connected           = $Connected
                AccessToken         = $AccessToken
                ProbeFails          = $ProbeFails
                ExpectedProbes      = $ExpectedProbes
                ExpectedConnected   = $ExpectedConnected
                ExpectedAccessToken = $ExpectedAccessToken
            } {
                param ($Connected, $AccessToken, $ProbeFails, $ExpectedProbes, $ExpectedConnected, $ExpectedAccessToken)

                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Connect-MSCloudLoginRESTWorkload -MockWith { return @{} }
                if ($ProbeFails)
                {
                    Mock -CommandName Invoke-WebRequest -MockWith { throw '401 Unauthorized' }
                }
                else
                {
                    Mock -CommandName Invoke-WebRequest -MockWith { return @{ StatusCode = 200 } }
                }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AuthenticationType = 'ServicePrincipalWithThumbprint'
                $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.BapEndpoint = 'bap.endpoint.com'
                $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AccessToken = $AccessToken
                $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.Connected = $Connected

                Connect-MSCloudLoginPowerPlatformREST

                $Script:MSCloudLoginConnectionProfile.PowerPlatformREST.Connected | Should -Be $ExpectedConnected
                "$($Script:MSCloudLoginConnectionProfile.PowerPlatformREST.AccessToken)" | Should -Be "$ExpectedAccessToken"
                Should -Invoke Invoke-WebRequest -Exactly $ExpectedProbes -ParameterFilter {
                    $Uri -eq 'https://bap.endpoint.com/providers/Microsoft.BusinessAppPlatform/scopes/admin/environments?api-version=2024-05-01' -and
                    $Headers.Authorization -eq 'Bearer token123'
                }
                Should -Invoke Connect-MSCloudLoginRESTWorkload -Exactly 1 -ParameterFilter {
                    $WorkloadName -eq 'PowerPlatformREST'
                }
            }
        }
    }
}

AfterAll {
    Remove-Module MSCloudLoginAssistant
}
