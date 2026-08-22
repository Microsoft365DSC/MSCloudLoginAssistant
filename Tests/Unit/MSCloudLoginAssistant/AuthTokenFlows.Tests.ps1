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

Describe 'Get-AuthToken token endpoint selection' {

    It 'Should use the v1.0 endpoint and a resource when a resource is requested' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Invoke-RestMethod -MockWith { return @{ access_token = 'token' } }

            Get-AuthToken -AuthorizationUrl 'https://login.microsoftonline.com' `
                -TenantId 'contoso.onmicrosoft.com' -ClientId 'client' -ClientSecret 'secret' `
                -Resource 'https://admin.microsoft.com'

            Should -Invoke Invoke-RestMethod -Exactly 1 -ParameterFilter {
                $Uri -eq 'https://login.microsoftonline.com/contoso.onmicrosoft.com/oauth2/token' -and
                $Body.resource -eq 'https://admin.microsoft.com' -and
                $null -eq $Body.scope
            }
        }
    }

    It 'Should prefer an explicitly supplied token endpoint over the derived one' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Invoke-RestMethod -MockWith { return @{ access_token = 'token' } }

            Get-AuthToken -AuthorizationUrl 'https://login.microsoftonline.com' `
                -TenantId 'contoso.onmicrosoft.com' -ClientId 'client' -ClientSecret 'secret' `
                -Scope 'https://graph.microsoft.com/.default' `
                -TokenEndpoint 'https://login.contoso.local/custom/oauth2/v2.0/token'

            Should -Invoke Invoke-RestMethod -Exactly 1 -ParameterFilter {
                $Uri -eq 'https://login.contoso.local/custom/oauth2/v2.0/token'
            }
        }
    }
}

Describe 'Get-AuthToken refresh token flow' {

    It 'Should exchange the refresh token for a resource scoped token' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Invoke-RestMethod -MockWith { return @{ access_token = 'refreshed' } }

            $result = Get-AuthToken -AuthorizationUrl 'https://login.microsoftonline.com' `
                -TenantId 'contoso.onmicrosoft.com' -ClientId 'client' `
                -RefreshToken 'the-refresh-token' -Resource 'https://admin.microsoft.com'

            $result.access_token | Should -Be 'refreshed'
            Should -Invoke Invoke-RestMethod -Exactly 1 -ParameterFilter {
                $Body.grant_type -eq 'refresh_token' -and
                $Body.refresh_token -eq 'the-refresh-token' -and
                $Body.resource -eq 'https://admin.microsoft.com'
            }
        }
    }
}

Describe 'Get-AuthToken password grant' {

    It 'Should send the user name and password of the credential for a resource scoped token' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Invoke-RestMethod -MockWith { return @{ access_token = 'password-grant' } }

            $credential = New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))
            Get-AuthToken -AuthorizationUrl 'https://login.microsoftonline.com' `
                -TenantId 'contoso.onmicrosoft.com' -ClientId 'client' `
                -Credentials $credential -Resource 'https://admin.microsoft.com'

            Should -Invoke Invoke-RestMethod -Exactly 1 -ParameterFilter {
                $Body.grant_type -eq 'password' -and
                $Body.username -eq 'admin@contoso.onmicrosoft.com' -and
                $Body.password -eq 'p@ssw0rd' -and
                $Body.resource -eq 'https://admin.microsoft.com'
            }
        }
    }
}

Describe 'Get-AuthToken device code flow' {

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

    It 'Should surface a terminal OAuth error immediately' {
        InModuleScope 'MSCloudLoginAssistant' {
            $script:deviceCodeCalls = 0
            Mock -CommandName Write-Verbose -MockWith { }
            Mock -CommandName Invoke-RestMethod -MockWith {
                $script:deviceCodeCalls++
                if ($script:deviceCodeCalls -eq 1)
                {
                    return @{ device_code = 'device-code'; interval = 0; message = 'Sign in please' }
                }
                $errorRecord = [System.Management.Automation.ErrorRecord]::new(
                    [System.Exception]::new('the user declined the sign-in'), 'AccessDenied', 'NotSpecified', $null)
                $errorRecord.ErrorDetails = [System.Management.Automation.ErrorDetails]::new('{"error":"access_denied"}')
                throw $errorRecord
            }

            { Get-AuthToken -AuthorizationUrl 'https://login.microsoftonline.com' `
                -TenantId 'contoso.onmicrosoft.com' -ClientId 'client' -DeviceCode `
                -Scope 'https://graph.microsoft.com/.default' } | Should -Throw '*declined the sign-in*'

            $script:deviceCodeCalls | Should -Be 2
        }
    }

    It 'Should treat a network failure during polling as terminal' {
        InModuleScope 'MSCloudLoginAssistant' {
            $script:deviceCodeCalls = 0
            Mock -CommandName Write-Verbose -MockWith { }
            Mock -CommandName Invoke-RestMethod -MockWith {
                $script:deviceCodeCalls++
                if ($script:deviceCodeCalls -eq 1)
                {
                    return @{ device_code = 'device-code'; interval = 0; message = 'Sign in please' }
                }
                $errorRecord = [System.Management.Automation.ErrorRecord]::new(
                    [System.Exception]::new('the remote name could not be resolved'), 'ConnectionFailure', 'NotSpecified', $null)
                $errorRecord.ErrorDetails = [System.Management.Automation.ErrorDetails]::new('the remote name could not be resolved')
                throw $errorRecord
            }

            { Get-AuthToken -AuthorizationUrl 'https://login.microsoftonline.com' `
                -TenantId 'contoso.onmicrosoft.com' -ClientId 'client' -DeviceCode `
                -Scope 'https://graph.microsoft.com/.default' } | Should -Throw '*could not be resolved*'
        }
    }
}

Describe 'Get-AuthToken managed identity in Azure Automation' {

    It 'Should send the identity header of the automation account' {
        InModuleScope 'MSCloudLoginAssistant' {
            $savedAzurePsHost = $env:AZUREPS_HOST_ENVIRONMENT
            $savedIdentityEndpoint = $env:IDENTITY_ENDPOINT
            $savedIdentityHeader = $env:IDENTITY_HEADER
            $env:AZUREPS_HOST_ENVIRONMENT = 'AzureAutomation/8.0'
            $env:IDENTITY_ENDPOINT = 'http://localhost:9999/metadata/identity/oauth2/token'
            $env:IDENTITY_HEADER = 'the-identity-header'

            Mock -CommandName Invoke-RestMethod -MockWith { return @{ access_token = 'automation-token' } }

            try
            {
                $token = Get-AuthToken -Identity -Resource 'https://graph.microsoft.com'
            }
            finally
            {
                $env:AZUREPS_HOST_ENVIRONMENT = $savedAzurePsHost
                $env:IDENTITY_ENDPOINT = $savedIdentityEndpoint
                $env:IDENTITY_HEADER = $savedIdentityHeader
            }

            $token | Should -Be 'automation-token'
            Should -Invoke Invoke-RestMethod -Exactly 1 -ParameterFilter {
                $Headers.'X-IDENTITY-HEADER' -eq 'the-identity-header' -and
                $Body.resource -eq 'https://graph.microsoft.com'
            }
        }
    }
}

Describe 'Connect-MSCloudLoginRESTWorkload client id fallback' {

    It 'Should use the client id of the caller when the profile has no application id' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
            Mock -CommandName Get-AuthToken -MockWith { return @{ token_type = 'Bearer'; access_token = 'token' } }

            $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
            $workloadProfile = $Script:MSCloudLoginConnectionProfile.EngageHub
            $workloadProfile.AuthenticationType = 'Credentials'
            $workloadProfile.RequestedAuthenticationType = 'Credentials'
            $workloadProfile.Credentials = New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))

            Connect-MSCloudLoginRESTWorkload -WorkloadName 'EngageHub' `
                -AuthorizationUrl 'https://login.microsoftonline.com' `
                -Scope 'https://engagehub.microsoft.com/.default' `
                -ClientId 'fallback-client-id'

            Should -Invoke Get-AuthToken -Exactly 1 -ParameterFilter { $ClientId -eq 'fallback-client-id' }
        }
    }

    It 'Should renew an expired token based connection instead of reusing it' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
            Mock -CommandName Get-AuthToken -MockWith { return @{ token_type = 'Bearer'; access_token = 'renewed-token' } }

            $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
            $workloadProfile = $Script:MSCloudLoginConnectionProfile.AdminAPI
            $workloadProfile.AuthenticationType = 'ServicePrincipalWithSecret'
            $workloadProfile.RequestedAuthenticationType = 'ServicePrincipalWithSecret'
            $workloadProfile.ApplicationId = 'app-id'
            $workloadProfile.ApplicationSecret = 'secret'
            $workloadProfile.TenantId = 'contoso.onmicrosoft.com'
            $workloadProfile.CompleteConnection()
            $workloadProfile.ConnectedDateTime = [System.DateTime]::Now.AddMinutes(-90).ToString()

            Connect-MSCloudLoginRESTWorkload -WorkloadName 'AdminAPI' `
                -AuthorizationUrl 'https://login.microsoftonline.com' -Scope 'scope/.default' -ClientId 'client'

            $workloadProfile.AccessToken | Should -Be 'Bearer renewed-token'
            Should -Invoke Get-AuthToken -Exactly 1
        }
    }
}
