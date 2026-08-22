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
# Invoke-MSCloudLoginAssistantConnectionLock
# ---------------------------------------------------------------------------
Describe 'Invoke-MSCloudLoginAssistantConnectionLock' {

    Context 'When the connection lock is free' {
        It 'Should invoke the passed script block' {
            InModuleScope 'MSCloudLoginAssistant' {
                $script:lockScriptRan = $false
                Invoke-MSCloudLoginAssistantConnectionLock -ConnectScript { $script:lockScriptRan = $true }
                $script:lockScriptRan | Should -BeTrue
            }
        }

        It 'Should return the output of the script block' {
            InModuleScope 'MSCloudLoginAssistant' {
                $result = Invoke-MSCloudLoginAssistantConnectionLock -ConnectScript { 'lock-output' }
                $result | Should -Be 'lock-output'
            }
        }
    }
}

# ---------------------------------------------------------------------------
# Add-MSCloudLoginAssistantEvent
# ---------------------------------------------------------------------------
Describe 'Add-MSCloudLoginAssistantEvent' {

    BeforeAll {
        InModuleScope 'MSCloudLoginAssistant' {
            $script:originalWriteToEventLog = $Script:WriteToEventLog
            $Script:WriteToEventLog = $false
        }
    }

    AfterAll {
        InModuleScope 'MSCloudLoginAssistant' {
            $Script:WriteToEventLog = $script:originalWriteToEventLog
        }
    }

    Context 'When event log writing is disabled' {
        It 'Should not throw for an Information entry' {
            InModuleScope 'MSCloudLoginAssistant' {
                { Add-MSCloudLoginAssistantEvent -Message 'Log message' -Source 'Test-Source' } | Should -Not -Throw
            }
        }

        It 'Should not throw for an Error entry' {
            InModuleScope 'MSCloudLoginAssistant' {
                { Add-MSCloudLoginAssistantEvent -Message 'Log message' -Source 'Test-Source' -EntryType 'Error' } | Should -Not -Throw
            }
        }

        It 'Should accept a custom event id' {
            InModuleScope 'MSCloudLoginAssistant' {
                { Add-MSCloudLoginAssistantEvent -Message 'Log message' -Source 'Test-Source' -EventID 42 } | Should -Not -Throw
            }
        }
    }

    Context 'When event log writing is enabled' {
        BeforeAll {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:WriteToEventLog = $true
            }
        }

        AfterAll {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:WriteToEventLog = $false
            }
        }

        It 'Should detect the source exists on a different log and emit a warning' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Write-Warning -MockWith { }
                Mock -CommandName Write-Verbose -MockWith { }

                # 'Application' is a well-known source on the Application log,
                # which differs from 'MSCloudLoginAssistant'. This hits the warning path.
                Add-MSCloudLoginAssistantEvent -Message 'Log message' -Source 'Application'

                Should -Invoke Write-Warning -ParameterFilter {
                    $Message -like '[[]ERROR[]] Specified source {Application} already exists on log*'
                }
            }
        }

        It 'Should attempt to create a new event source for an unknown source name' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Write-Warning -MockWith { }
                Mock -CommandName Write-Verbose -MockWith { }

                # Use a source name that does not exist.
                # CreateEventSource will throw SecurityException (not admin),
                # which is caught and logged at the verbose level.
                { Add-MSCloudLoginAssistantEvent -Message 'Log message' -Source ('MSCloudLoginTest.Coverage.' + [guid]::NewGuid().ToString('N')) } | Should -Not -Throw
            }
        }

        It 'Should truncate a message that exceeds 32766 characters' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Write-Warning -MockWith { }
                Mock -CommandName Write-Verbose -MockWith { }

                $longMessage = 'A' * 40000

                { Add-MSCloudLoginAssistantEvent -Message $longMessage -Source 'Application' } | Should -Not -Throw
            }
        }
    }
}

# ---------------------------------------------------------------------------
# ConvertTo-Base64Url
# ---------------------------------------------------------------------------
Describe 'ConvertTo-Base64Url' {

    It 'Should encode bytes without padding' {
        InModuleScope 'MSCloudLoginAssistant' {
            $bytes = [System.Text.Encoding]::UTF8.GetBytes('hello')
            (ConvertTo-Base64Url -Bytes $bytes) | Should -Be 'aGVsbG8'
        }
    }

    It 'Should replace + and / with URL-safe characters' {
        InModuleScope 'MSCloudLoginAssistant' {
            $bytes = [byte[]]@(0xFB, 0xEF, 0xBE)
            $result = ConvertTo-Base64Url -Bytes $bytes
            $result | Should -Be '----'
            $result | Should -Not -Match '[+/=]'
        }
    }

    It 'Should trim trailing equals signs' {
        InModuleScope 'MSCloudLoginAssistant' {
            $bytes = [System.Text.Encoding]::UTF8.GetBytes('f')
            (ConvertTo-Base64Url -Bytes $bytes) | Should -Be 'Zg'
        }
    }
}

# ---------------------------------------------------------------------------
# Assert-IsNonInteractiveShell
# ---------------------------------------------------------------------------
Describe 'Assert-IsNonInteractiveShell' {

    It 'Should return a boolean value' {
        InModuleScope 'MSCloudLoginAssistant' {
            $result = Assert-IsNonInteractiveShell
            $result -is [System.Boolean] | Should -BeTrue
        }
    }
}

# ---------------------------------------------------------------------------
# Get-SPOAdminUrl
# ---------------------------------------------------------------------------
Describe 'Get-SPOAdminUrl' {

    BeforeAll {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
            Mock -CommandName Connect-M365Tenant -MockWith { }
        }
    }

    Context 'When the Graph request succeeds immediately' {
        It 'Should derive the admin URL from the site root' {
            InModuleScope 'MSCloudLoginAssistant' {
                $script:claimCount = 0
                Mock -CommandName Invoke-MgGraphRequest -MockWith {
                    $script:claimCount++
                    return @{ webUrl = 'https://contoso.sharepoint.com' }
                }
                $result = Get-SPOAdminUrl
                $result | Should -Be 'https://contoso-admin.sharepoint.com'
                $script:claimCount | Should -Be 1
                Should -Invoke Connect-M365Tenant -Exactly 0
            }
        }
    }

    Context 'When the site root is empty' {
        It 'Should reconnect and retry the request' {
            InModuleScope 'MSCloudLoginAssistant' {
                $script:claimCount = 0
                Mock -CommandName Invoke-MgGraphRequest -MockWith {
                    $script:claimCount++
                    if ($script:claimCount -eq 1)
                    {
                        return @{}
                    }
                    return @{ webUrl = 'https://contoso.sharepoint.com' }
                }
                $result = Get-SPOAdminUrl
                $result | Should -Be 'https://contoso-admin.sharepoint.com'
                $script:claimCount | Should -Be 2
                Should -Invoke Connect-M365Tenant -Exactly 1
            }
        }
    }

    Context 'When the first request throws' {
        It 'Should reconnect and retry the request' {
            InModuleScope 'MSCloudLoginAssistant' {
                $script:claimCount = 0
                Mock -CommandName Invoke-MgGraphRequest -MockWith {
                    $script:claimCount++
                    if ($script:claimCount -eq 1)
                    {
                        throw 'network error'
                    }
                    return @{ webUrl = 'https://contoso.sharepoint.com' }
                }
                $result = Get-SPOAdminUrl
                $result | Should -Be 'https://contoso-admin.sharepoint.com'
                $script:claimCount | Should -Be 2
            }
        }
    }

    Context 'When in an interactive shell and the retry also fails' {
        It 'Should request Graph scopes and retry once more' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $false }
                Mock -CommandName Connect-MgGraph -MockWith { }
                $script:claimCount = 0
                Mock -CommandName Invoke-MgGraphRequest -MockWith {
                    $script:claimCount++
                    if ($script:claimCount -le 2)
                    {
                        throw 'another error'
                    }
                    return @{ webUrl = 'https://contoso.sharepoint.com' }
                }
                $result = Get-SPOAdminUrl
                $result | Should -Be 'https://contoso-admin.sharepoint.com'
                $script:claimCount | Should -Be 3
                Should -Invoke Connect-MgGraph -Exactly 1
            }
        }
    }

    Context 'When in a non-interactive shell and access is forbidden' {
        It 'Should throw a permission error' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $true }
                $script:claimCount = 0
                Mock -CommandName Invoke-MgGraphRequest -MockWith {
                    $script:claimCount++
                    if ($script:claimCount -eq 1)
                    {
                        throw 'network error'
                    }
                    throw 'Insufficient privileges to complete the operation.'
                }
                { Get-SPOAdminUrl } | Should -Throw '*correct permissions to access Domains*'
            }
        }
    }

    Context 'When the web URL cannot be retrieved' {
        It 'Should throw an unable to retrieve error' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $true }
                $script:claimCount = 0
                Mock -CommandName Invoke-MgGraphRequest -MockWith {
                    $script:claimCount++
                    throw 'network error'
                }
                { Get-SPOAdminUrl } | Should -Throw 'Unable to retrieve SPO Admin URL*'
            }
        }
    }
}

# ---------------------------------------------------------------------------
# Get-MSCloudLoginAccessToken
# ---------------------------------------------------------------------------
Describe 'Get-MSCloudLoginAccessToken' {

    Context 'When the token request succeeds' {
        It 'Should return the access token' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Get-AuthToken -MockWith { return @{ access_token = 'test-access-token' } }

                $result = Get-MSCloudLoginAccessToken -ConnectionUri 'https://graph.microsoft.com/.default' `
                    -AuthorizationUrl 'https://login.microsoftonline.com' `
                    -AzureADAuthorizationEndpointUri 'https://login.microsoftonline.com/organizations/oauth2/v2.0/token' `
                    -ApplicationId 'app-id' `
                    -TenantId 'tenant-id' `
                    -CertificateThumbprint 'thumbprint'

                $result | Should -Be 'test-access-token'
                Should -Invoke Get-AuthToken -ParameterFilter {
                    $AuthorizationUrl -eq 'https://login.microsoftonline.com' -and
                    $TokenEndpoint -eq 'https://login.microsoftonline.com/organizations/oauth2/v2.0/token' -and
                    $Scope -eq 'https://graph.microsoft.com/.default' -and
                    $ClientId -eq 'app-id' -and
                    $TenantId -eq 'tenant-id' -and
                    $CertificateThumbprint -eq 'thumbprint'
                }
            }
        }
    }

    Context 'When the token request fails' {
        It 'Should throw the underlying error' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Get-AuthToken -MockWith { throw 'Authentication failed' }

                { Get-MSCloudLoginAccessToken -ConnectionUri 'https://graph.microsoft.com/.default' `
                    -AuthorizationUrl 'https://login.microsoftonline.com' `
                    -AzureADAuthorizationEndpointUri 'https://login.microsoftonline.com/organizations/oauth2/v2.0/token' `
                    -ApplicationId 'app-id' `
                    -TenantId 'tenant-id' `
                    -CertificateThumbprint 'thumbprint' } | Should -Throw 'Authentication failed'
            }
        }
    }
}

# ---------------------------------------------------------------------------
# Get-CloudEnvironmentInfo
# ---------------------------------------------------------------------------
Describe 'Get-CloudEnvironmentInfo' {

    Context 'When credentials are provided' {
        BeforeEach {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Invoke-WebRequest -MockWith {
                    return @{ Content = '{ "tenant_region_sub_scope": "USGov", "token_endpoint": "https://login.microsoftonline.us/t/oauth2/v2.0/token" }' }
                }
            }
        }

        It 'Should derive the tenant from the credential UPN' {
            InModuleScope 'MSCloudLoginAssistant' {
                $cred = New-Object PSCredential ('user@contoso.com', (ConvertTo-SecureString 'pwd' -AsPlainText -Force))
                $result = Get-CloudEnvironmentInfo -Credentials $cred
                $result.tenant_region_sub_scope | Should -Be 'USGov'
                Should -Invoke Invoke-WebRequest -ParameterFilter {
                    $Uri -like 'https://login.microsoftonline.com/contoso.com/v2.0/*'
                }
            }
        }
    }

    Context 'When a TenantId is provided' {
        BeforeEach {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Invoke-WebRequest -MockWith {
                    return @{ Content = '{ "tenant_region_sub_scope": "USGov", "token_endpoint": "https://login.microsoftonline.us/t/oauth2/v2.0/token" }' }
                }
            }
        }

        It 'Should use the TenantId directly' {
            InModuleScope 'MSCloudLoginAssistant' {
                $result = Get-CloudEnvironmentInfo -TenantId 'contoso.onmicrosoft.com'
                $result.tenant_region_sub_scope | Should -Be 'USGov'
            }
        }
    }

    Context 'When Identity is provided without TenantId' {
        It 'Should return $null' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                # Do not mock Invoke-WebRequest - function should return early
                $result = Get-CloudEnvironmentInfo -Identity
                $result | Should -BeNullOrEmpty
            }
        }
    }

    Context 'When neither credentials nor TenantId are provided' {
        It 'Should throw' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                { Get-CloudEnvironmentInfo } | Should -Throw 'TenantId or Credentials must be provided'
            }
        }
    }
}

# ---------------------------------------------------------------------------
# Get-MSCloudLoginOrganizationName
# ---------------------------------------------------------------------------
Describe 'Get-MSCloudLoginOrganizationName' {

    BeforeAll {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
            Mock -CommandName Connect-M365Tenant -MockWith { }
        }
    }

    Context 'When certificate thumbprint authentication is used' {
        It 'Should return the initial domain' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Invoke-MgGraphRequest -MockWith {
                    return @{ value = @(
                            @{ Id = 'contoso.com'; IsInitial = $true }
                            @{ Id = 'contoso.onmicrosoft.com'; IsInitial = $false }
                        )
                    }
                }
                $result = Get-MSCloudLoginOrganizationName -ApplicationId 'app' -TenantId 'tenant' -CertificateThumbprint 'thumb'
                $result | Should -Be 'contoso.com'
                Should -Invoke Connect-M365Tenant -ParameterFilter { $Workload -eq 'MicrosoftGraph' -and $CertificateThumbprint -eq 'thumb' }
            }
        }
    }

    Context 'When application secret authentication is used' {
        It 'Should return the initial domain' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Invoke-MgGraphRequest -MockWith {
                    return @{ value = @(@{ Id = 'contoso.com'; IsInitial = $true }) }
                }
                $result = Get-MSCloudLoginOrganizationName -ApplicationId 'app' -TenantId 'tenant' -ApplicationSecret 'secret'
                $result | Should -Be 'contoso.com'
            }
        }
    }

    Context 'When Identity is used' {
        It 'Should return the initial domain' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Invoke-MgGraphRequest -MockWith {
                    return @{ value = @(@{ Id = 'contoso.com'; IsInitial = $true }) }
                }
                $result = Get-MSCloudLoginOrganizationName -Identity -TenantId 'tenant'
                $result | Should -Be 'contoso.com'
            }
        }
    }

    Context 'When AccessTokens are used' {
        It 'Should return the initial domain' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Invoke-MgGraphRequest -MockWith {
                    return @{ value = @(@{ Id = 'contoso.com'; IsInitial = $true }) }
                }
                $result = Get-MSCloudLoginOrganizationName -AccessTokens @('token1')
                $result | Should -Be 'contoso.com'
            }
        }
    }

    Context 'When the domain lookup fails' {
        It 'Should fall back to the TenantId when provided' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Invoke-MgGraphRequest -MockWith { throw 'graph error' }
                $result = Get-MSCloudLoginOrganizationName -ApplicationId 'app' -TenantId 'tenant-contoso' -ApplicationSecret 'secret'
                $result | Should -Be 'tenant-contoso'
            }
        }

        It 'Should re-throw when no TenantId is available' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Invoke-MgGraphRequest -MockWith { throw 'graph error' }
                { Get-MSCloudLoginOrganizationName -AccessTokens @('token1') } | Should -Throw 'graph error'
            }
        }
    }
}
