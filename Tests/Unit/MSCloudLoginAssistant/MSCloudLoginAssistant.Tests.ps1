#Requires -Modules Pester

BeforeAll {
    # Import stubs so that external cmdlets are available as no-ops
    $stubModule = Join-Path $PSScriptRoot '..\Stubs\Stubs.psm1'
    Import-Module $stubModule -Force -WarningAction SilentlyContinue

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

    Context 'When ServicePrincipalWithThumbprint parameters are provided' {
        It 'Should return ServicePrincipalWithThumbprint' {
            InModuleScope 'MSCloudLoginAssistant' {
                $params = @{
                    ApplicationId        = 'app-id'
                    TenantId             = 'tenant-id'
                    CertificateThumbprint = 'thumb'
                }
                $result = Get-AuthenticationTypeFromParameters -AuthenticationObject $params
                $result | Should -Be 'ServicePrincipalWithThumbprint'
            }
        }
    }

    Context 'When ServicePrincipalWithSecret parameters are provided' {
        It 'Should return ServicePrincipalWithSecret' {
            InModuleScope 'MSCloudLoginAssistant' {
                $params = @{
                    ApplicationId     = 'app-id'
                    TenantId          = 'tenant-id'
                    ApplicationSecret = 'secret'
                }
                $result = Get-AuthenticationTypeFromParameters -AuthenticationObject $params
                $result | Should -Be 'ServicePrincipalWithSecret'
            }
        }
    }

    Context 'When ServicePrincipalWithPath parameters are provided' {
        It 'Should return ServicePrincipalWithPath' {
            InModuleScope 'MSCloudLoginAssistant' {
                $secPwd = ConvertTo-SecureString 'pass' -AsPlainText -Force
                $params = @{
                    ApplicationId       = 'app-id'
                    TenantId            = 'tenant-id'
                    CertificatePath     = 'C:\cert.pfx'
                    CertificatePassword = $secPwd
                }
                $result = Get-AuthenticationTypeFromParameters -AuthenticationObject $params
                $result | Should -Be 'ServicePrincipalWithPath'
            }
        }
    }

    Context 'When CredentialsWithApplicationId parameters are provided' {
        It 'Should return CredentialsWithApplicationId' {
            InModuleScope 'MSCloudLoginAssistant' {
                $secPwd = ConvertTo-SecureString 'pass' -AsPlainText -Force
                $cred = New-Object PSCredential ('user@contoso.com', $secPwd)
                $params = @{
                    Credentials   = $cred
                    ApplicationId = 'app-id'
                }
                $result = Get-AuthenticationTypeFromParameters -AuthenticationObject $params
                $result | Should -Be 'CredentialsWithApplicationId'
            }
        }
    }

    Context 'When CredentialsWithTenantId parameters are provided' {
        It 'Should return CredentialsWithTenantId' {
            InModuleScope 'MSCloudLoginAssistant' {
                $secPwd = ConvertTo-SecureString 'pass' -AsPlainText -Force
                $cred = New-Object PSCredential ('user@contoso.com', $secPwd)
                $params = @{
                    Credentials = $cred
                    TenantId    = 'tenant-id'
                }
                $result = Get-AuthenticationTypeFromParameters -AuthenticationObject $params
                $result | Should -Be 'CredentialsWithTenantId'
            }
        }
    }

    Context 'When only Credentials are provided' {
        It 'Should return Credentials' {
            InModuleScope 'MSCloudLoginAssistant' {
                $secPwd = ConvertTo-SecureString 'pass' -AsPlainText -Force
                $cred = New-Object PSCredential ('user@contoso.com', $secPwd)
                $params = @{ Credentials = $cred }
                $result = Get-AuthenticationTypeFromParameters -AuthenticationObject $params
                $result | Should -Be 'Credentials'
            }
        }
    }

    Context 'When Identity is provided' {
        It 'Should return Identity' {
            InModuleScope 'MSCloudLoginAssistant' {
                $params = @{ Identity = $true }
                $result = Get-AuthenticationTypeFromParameters -AuthenticationObject $params
                $result | Should -Be 'Identity'
            }
        }
    }

    Context 'When AccessTokens are provided' {
        It 'Should return AccessTokens' {
            InModuleScope 'MSCloudLoginAssistant' {
                $params = @{
                    AccessTokens = @('token1', 'token2')
                    TenantId     = 'tenant-id'
                }
                $result = Get-AuthenticationTypeFromParameters -AuthenticationObject $params
                $result | Should -Be 'AccessTokens'
            }
        }
    }

    Context 'When no recognised parameters are provided' {
        It 'Should return Interactive' {
            InModuleScope 'MSCloudLoginAssistant' {
                $params = @{}
                $result = Get-AuthenticationTypeFromParameters -AuthenticationObject $params
                $result | Should -Be 'Interactive'
            }
        }
    }
}

# ---------------------------------------------------------------------------
# MSCloudLoginConnectionProfile class
# ---------------------------------------------------------------------------
Describe 'MSCloudLoginConnectionProfile' {

    Context 'Constructor defaults' {
        It 'Should initialise all workload objects' {
            InModuleScope 'MSCloudLoginAssistant' {
                $cloudProfile = New-Object MSCloudLoginConnectionProfile

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

        It 'Should set CreatedTime' {
            InModuleScope 'MSCloudLoginAssistant' {
                $cloudProfile = New-Object MSCloudLoginConnectionProfile
                $cloudProfile.CreatedTime | Should -Not -BeNullOrEmpty
            }
        }
    }

    Context 'Workload default ApplicationIds' {
        It 'Should set correct default ApplicationId for AdminAPI' {
            InModuleScope 'MSCloudLoginAssistant' {
                $instance = New-Object AdminAPI
                $instance.ApplicationId | Should -Be '1950a258-227b-4e31-a9cf-717495945fc2'
            }
        }

        It 'Should set correct default ApplicationId for Fabric' {
            InModuleScope 'MSCloudLoginAssistant' {
                $instance = New-Object Fabric
                $instance.ApplicationId | Should -Be '23d8f6bd-1eb0-4cc2-a08c-7bf525c67bcd'
            }
        }

        It 'Should set correct default ApplicationId for Tasks' {
            InModuleScope 'MSCloudLoginAssistant' {
                $instance = New-Object Tasks
                $instance.ApplicationId | Should -Be '9ac8c0b3-2c30-497c-b4bc-cadfe9bd6eed'
            }
        }

        It 'Should set correct default ApplicationId for SharePointOnlineREST' {
            InModuleScope 'MSCloudLoginAssistant' {
                $instance = New-Object SharePointOnlineREST
                $instance.ApplicationId | Should -Be '31359c7f-bd7e-475c-86db-fdb8c937548e'
            }
        }
    }

    Context 'Workload CompleteConnection' {
        It 'Should mark the workload as connected' {
            InModuleScope 'MSCloudLoginAssistant' {
                $instance = New-Object AdminAPI
                $instance.Connected | Should -BeFalse
                $instance.CompleteConnection()
                $instance.Connected | Should -BeTrue
                $instance.ConnectedDateTime | Should -Not -BeNullOrEmpty
            }
        }

        It 'Should track MFA usage when specified' {
            InModuleScope 'MSCloudLoginAssistant' {
                $instance = New-Object AdminAPI
                $instance.CompleteConnection($true)
                $instance.MultiFactorAuthentication | Should -BeTrue
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

    Context 'When connecting to AdminAPI' {
        It 'Should invoke the AdminAPI connect function' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                Connect-M365Tenant -Workload 'AdminAPI' -ApplicationId 'app-id' -TenantId 'tenant-id' -ApplicationSecret 'secret'
                Should -Invoke Connect-MSCloudLoginAdminAPI -Exactly 1
            }
        }
    }

    Context 'When connecting to ExchangeOnline' {
        It 'Should invoke the ExchangeOnline connect function' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                Connect-M365Tenant -Workload 'ExchangeOnline' -ApplicationId 'app-id' -TenantId 'tenant-id' -CertificateThumbprint 'thumb'
                Should -Invoke Connect-MSCloudLoginExchangeOnline -Exactly 1
            }
        }
    }

    Context 'When connecting to MicrosoftGraph' {
        It 'Should invoke the MicrosoftGraph connect function' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                Connect-M365Tenant -Workload 'MicrosoftGraph' -ApplicationId 'app-id' -TenantId 'tenant-id' -ApplicationSecret 'secret'
                Should -Invoke Connect-MSCloudLoginMicrosoftGraph -Exactly 1
            }
        }
    }

    Context 'When connecting to MicrosoftTeams' {
        It 'Should map to Teams and invoke the Teams connect function' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                Connect-M365Tenant -Workload 'MicrosoftTeams' -ApplicationId 'app-id' -TenantId 'tenant-id' -CertificateThumbprint 'thumb'
                Should -Invoke Connect-MSCloudLoginTeams -Exactly 1
            }
        }
    }

    Context 'When connecting to PowerPlatforms' {
        It 'Should map to PowerPlatform and invoke the PowerPlatform connect function' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                Connect-M365Tenant -Workload 'PowerPlatforms' -ApplicationId 'app-id' -TenantId 'tenant-id' -CertificateThumbprint 'thumb'
                Should -Invoke Connect-MSCloudLoginPowerPlatform -Exactly 1
            }
        }
    }

    Context 'When connecting to SecurityComplianceCenter' {
        It 'Should invoke the SecurityCompliance connect function' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                Connect-M365Tenant -Workload 'SecurityComplianceCenter' -ApplicationId 'app-id' -TenantId 'tenant-id' -CertificateThumbprint 'thumb'
                Should -Invoke Connect-MSCloudLoginSecurityCompliance -Exactly 1
            }
        }
    }

    Context 'When connecting to Tasks' {
        It 'Should invoke the Tasks connect function' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                Connect-M365Tenant -Workload 'Tasks' -ApplicationId 'app-id' -TenantId 'tenant-id' -ApplicationSecret 'secret'
                Should -Invoke Connect-MSCloudLoginTasks -Exactly 1
            }
        }
    }

    Context 'When setting authentication parameters' {
        It 'Should set the authentication type on the workload profile' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                Connect-M365Tenant -Workload 'AdminAPI' -ApplicationId 'app-id' -TenantId 'tenant-id' -CertificateThumbprint 'thumb'

                $Script:MSCloudLoginConnectionProfile.AdminAPI.ApplicationId        | Should -Be 'app-id'
                $Script:MSCloudLoginConnectionProfile.AdminAPI.TenantId             | Should -Be 'tenant-id'
                $Script:MSCloudLoginConnectionProfile.AdminAPI.CertificateThumbprint | Should -Be 'thumb'
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
        It 'Should return false' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = $null
                $params = @{ Workload = 'AdminAPI' }
                $result = Compare-InputParametersForChange -CurrentParamSet $params
                $result | Should -BeFalse
            }
        }
    }

    Context 'When authentication type changes' {
        It 'Should return true' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.AdminAPI.AuthenticationType          = 'Credentials'
                $Script:MSCloudLoginConnectionProfile.AdminAPI.RequestedAuthenticationType = 'ServicePrincipalWithThumbprint'

                $params = @{
                    Workload      = 'AdminAPI'
                    ApplicationId = 'app-id'
                    TenantId      = 'tenant-id'
                }
                $result = Compare-InputParametersForChange -CurrentParamSet $params
                $result | Should -BeTrue
            }
        }
    }

    Context 'When parameters have not changed' {
        It 'Should return false' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.AdminAPI.AuthenticationType          = 'ServicePrincipalWithThumbprint'
                $Script:MSCloudLoginConnectionProfile.AdminAPI.RequestedAuthenticationType = 'ServicePrincipalWithThumbprint'
                $Script:MSCloudLoginConnectionProfile.AdminAPI.ApplicationId               = 'app-id'
                $Script:MSCloudLoginConnectionProfile.AdminAPI.TenantId                    = 'tenant-id'
                $Script:MSCloudLoginConnectionProfile.AdminAPI.CertificateThumbprint       = 'thumb'

                $params = @{
                    Workload              = 'AdminAPI'
                    ApplicationId         = 'app-id'
                    TenantId              = 'tenant-id'
                    CertificateThumbprint = 'thumb'
                }
                $result = Compare-InputParametersForChange -CurrentParamSet $params
                $result | Should -BeFalse
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
                Start-Sleep -Milliseconds 50
                Reset-MSCloudLoginConnectionProfileContext

                $Script:MSCloudLoginConnectionProfile.CreatedTime | Should -Not -Be $originalTime
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
