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

    Context 'When connecting to Azure with a SubscriptionId' {
        It 'Should persist the SubscriptionId on the Azure connection profile' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                Connect-M365Tenant -Workload 'Azure' -ApplicationId 'app-id' -TenantId 'tenant-id' -ApplicationSecret 'secret' -SubscriptionId 'sub-A'

                $Script:MSCloudLoginConnectionProfile.Azure.SubscriptionId | Should -Be 'sub-A'
            }
        }

        It 'Should keep the session valid when the same SubscriptionId is provided again' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:AzureConnectedStates = [System.Collections.Generic.List[bool]]::new()
                Mock -CommandName Connect-MSCloudLoginAzure -MockWith {
                    $Script:AzureConnectedStates.Add($Script:MSCloudLoginConnectionProfile.Azure.Connected)
                    $Script:MSCloudLoginConnectionProfile.Azure.CompleteConnection()
                }

                Connect-M365Tenant -Workload 'Azure' -ApplicationId 'app-id' -TenantId 'tenant-id' -ApplicationSecret 'secret' -SubscriptionId 'sub-A'
                Connect-M365Tenant -Workload 'Azure' -ApplicationId 'app-id' -TenantId 'tenant-id' -ApplicationSecret 'secret' -SubscriptionId 'sub-A'

                $Script:AzureConnectedStates | Should -Be @($false, $true)
                $Script:MSCloudLoginConnectionProfile.Azure.SubscriptionId | Should -Be 'sub-A'
            }
        }

        It 'Should invalidate the session when a different SubscriptionId is provided' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:AzureConnectedStates = [System.Collections.Generic.List[bool]]::new()
                Mock -CommandName Connect-MSCloudLoginAzure -MockWith {
                    $Script:AzureConnectedStates.Add($Script:MSCloudLoginConnectionProfile.Azure.Connected)
                    $Script:MSCloudLoginConnectionProfile.Azure.CompleteConnection()
                }

                Connect-M365Tenant -Workload 'Azure' -ApplicationId 'app-id' -TenantId 'tenant-id' -ApplicationSecret 'secret' -SubscriptionId 'sub-A'
                Connect-M365Tenant -Workload 'Azure' -ApplicationId 'app-id' -TenantId 'tenant-id' -ApplicationSecret 'secret' -SubscriptionId 'sub-B'

                $Script:AzureConnectedStates | Should -Be @($false, $false)
                $Script:MSCloudLoginConnectionProfile.Azure.SubscriptionId | Should -Be 'sub-B'
            }
        }

        It 'Should treat an omitted SubscriptionId as drift and clear it on the profile' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:AzureConnectedStates = [System.Collections.Generic.List[bool]]::new()
                Mock -CommandName Connect-MSCloudLoginAzure -MockWith {
                    $Script:AzureConnectedStates.Add($Script:MSCloudLoginConnectionProfile.Azure.Connected)
                    $Script:MSCloudLoginConnectionProfile.Azure.CompleteConnection()
                }

                Connect-M365Tenant -Workload 'Azure' -ApplicationId 'app-id' -TenantId 'tenant-id' -ApplicationSecret 'secret' -SubscriptionId 'sub-A'
                Connect-M365Tenant -Workload 'Azure' -ApplicationId 'app-id' -TenantId 'tenant-id' -ApplicationSecret 'secret'
                Connect-M365Tenant -Workload 'Azure' -ApplicationId 'app-id' -TenantId 'tenant-id' -ApplicationSecret 'secret'

                # Only the first call that drops the SubscriptionId reconnects, afterwards the
                # profile is back to the unset value and further calls converge again.
                $Script:AzureConnectedStates | Should -Be @($false, $false, $true)
                $Script:MSCloudLoginConnectionProfile.Azure.SubscriptionId | Should -BeNullOrEmpty
            }
        }

        It 'Should invalidate the session when the SubscriptionId is set again after being omitted' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:AzureConnectedStates = [System.Collections.Generic.List[bool]]::new()
                Mock -CommandName Connect-MSCloudLoginAzure -MockWith {
                    $Script:AzureConnectedStates.Add($Script:MSCloudLoginConnectionProfile.Azure.Connected)
                    $Script:MSCloudLoginConnectionProfile.Azure.CompleteConnection()
                }

                Connect-M365Tenant -Workload 'Azure' -ApplicationId 'app-id' -TenantId 'tenant-id' -ApplicationSecret 'secret' -SubscriptionId 'sub-A'
                Connect-M365Tenant -Workload 'Azure' -ApplicationId 'app-id' -TenantId 'tenant-id' -ApplicationSecret 'secret'
                Connect-M365Tenant -Workload 'Azure' -ApplicationId 'app-id' -TenantId 'tenant-id' -ApplicationSecret 'secret' -SubscriptionId 'sub-B'

                $Script:AzureConnectedStates | Should -Be @($false, $false, $false)
                $Script:MSCloudLoginConnectionProfile.Azure.SubscriptionId | Should -Be 'sub-B'
            }
        }

        It 'Should detect a SubscriptionId change even when no identity parameter is repeated' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:AzureConnectedStates = [System.Collections.Generic.List[bool]]::new()
                Mock -CommandName Connect-MSCloudLoginAzure -MockWith {
                    $Script:AzureConnectedStates.Add($Script:MSCloudLoginConnectionProfile.Azure.Connected)
                    $Script:MSCloudLoginConnectionProfile.Azure.CompleteConnection()
                }

                Connect-M365Tenant -Workload 'Azure' -ApplicationId 'app-id' -TenantId 'tenant-id' -ApplicationSecret 'secret' -SubscriptionId 'sub-A'
                # Same subscription, no identity parameters: the session must be reused.
                Connect-M365Tenant -Workload 'Azure' -SubscriptionId 'sub-A'
                # Different subscription, still no identity parameters: this is drift.
                Connect-M365Tenant -Workload 'Azure' -SubscriptionId 'sub-B'

                $Script:AzureConnectedStates | Should -Be @($false, $true, $false)
                $Script:MSCloudLoginConnectionProfile.Azure.SubscriptionId | Should -Be 'sub-B'
            }
        }
    }

    Context 'When connecting to ExchangeOnline with cmdlets to load' {
        It 'Should map ExchangeOnlineCmdlets onto the CmdletsToLoad property' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                Connect-M365Tenant -Workload 'ExchangeOnline' -ApplicationId 'app-id' -TenantId 'tenant-id' -ApplicationSecret 'secret' -ExchangeOnlineCmdlets @('Get-Mailbox')

                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CmdletsToLoad | Should -Be @('Get-Mailbox')
            }
        }

        It 'Should treat omitted cmdlets as drift and clear them on the profile' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:ExoConnectedStates = [System.Collections.Generic.List[bool]]::new()
                Mock -CommandName Connect-MSCloudLoginExchangeOnline -MockWith {
                    $Script:ExoConnectedStates.Add($Script:MSCloudLoginConnectionProfile.ExchangeOnline.Connected)
                    $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CompleteConnection()
                }

                Connect-M365Tenant -Workload 'ExchangeOnline' -ApplicationId 'app-id' -TenantId 'tenant-id' -ApplicationSecret 'secret' -ExchangeOnlineCmdlets @('Get-Mailbox')
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

    Context 'When two parameter values are swapped' {
        It 'Should detect the change' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.AdminAPI.AuthenticationType          = 'ServicePrincipalWithSecret'
                $Script:MSCloudLoginConnectionProfile.AdminAPI.RequestedAuthenticationType = 'ServicePrincipalWithSecret'
                $Script:MSCloudLoginConnectionProfile.AdminAPI.ApplicationId               = 'value-A'
                $Script:MSCloudLoginConnectionProfile.AdminAPI.TenantId                    = 'value-B'
                $Script:MSCloudLoginConnectionProfile.AdminAPI.ApplicationSecret           = 'secret'

                $params = @{
                    Workload          = 'AdminAPI'
                    ApplicationId     = 'value-B'
                    TenantId          = 'value-A'
                    ApplicationSecret = 'secret'
                }
                (Compare-InputParametersForChange -CurrentParamSet $params) | Should -BeTrue
            }
        }
    }

    Context 'When the credential password changes' {
        It 'Should detect the change for the same user name' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $oldCred = New-Object PSCredential ('user@contoso.com', (ConvertTo-SecureString 'OldPwd' -AsPlainText -Force))
                $newCred = New-Object PSCredential ('user@contoso.com', (ConvertTo-SecureString 'NewPwd' -AsPlainText -Force))
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.AuthenticationType          = 'Credentials'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.RequestedAuthenticationType = 'Credentials'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Credentials                 = $oldCred

                $params = @{
                    Workload   = 'ExchangeOnline'
                    Credential = $newCred
                }
                (Compare-InputParametersForChange -CurrentParamSet $params) | Should -BeTrue
            }
        }

        It 'Should report no change for an identical credential' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $cred1 = New-Object PSCredential ('user@contoso.com', (ConvertTo-SecureString 'SamePwd' -AsPlainText -Force))
                $cred2 = New-Object PSCredential ('USER@contoso.com', (ConvertTo-SecureString 'SamePwd' -AsPlainText -Force))
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.AuthenticationType          = 'Credentials'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.RequestedAuthenticationType = 'Credentials'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.Credentials                 = $cred1

                $params = @{
                    Workload   = 'ExchangeOnline'
                    Credential = $cred2
                }
                (Compare-InputParametersForChange -CurrentParamSet $params) | Should -BeFalse
            }
        }
    }

    Context 'When a parameter exists on only one side' {
        It 'Should detect a newly provided SubscriptionId for Azure' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Azure.AuthenticationType          = 'ServicePrincipalWithSecret'
                $Script:MSCloudLoginConnectionProfile.Azure.RequestedAuthenticationType = 'ServicePrincipalWithSecret'
                $Script:MSCloudLoginConnectionProfile.Azure.ApplicationId               = 'app-id'
                $Script:MSCloudLoginConnectionProfile.Azure.TenantId                    = 'tenant-id'
                $Script:MSCloudLoginConnectionProfile.Azure.ApplicationSecret           = 'secret'

                $params = @{
                    Workload          = 'Azure'
                    ApplicationId     = 'app-id'
                    TenantId          = 'tenant-id'
                    ApplicationSecret = 'secret'
                    SubscriptionId    = 'sub-B'
                }
                (Compare-InputParametersForChange -CurrentParamSet $params) | Should -BeTrue
            }
        }

        It 'Should detect an omitted SubscriptionId as a change' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.Azure.AuthenticationType          = 'ServicePrincipalWithSecret'
                $Script:MSCloudLoginConnectionProfile.Azure.RequestedAuthenticationType = 'ServicePrincipalWithSecret'
                $Script:MSCloudLoginConnectionProfile.Azure.ApplicationId               = 'app-id'
                $Script:MSCloudLoginConnectionProfile.Azure.TenantId                    = 'tenant-id'
                $Script:MSCloudLoginConnectionProfile.Azure.ApplicationSecret           = 'secret'
                $Script:MSCloudLoginConnectionProfile.Azure.SubscriptionId              = 'sub-A'

                $params = @{
                    Workload          = 'Azure'
                    ApplicationId     = 'app-id'
                    TenantId          = 'tenant-id'
                    ApplicationSecret = 'secret'
                }
                (Compare-InputParametersForChange -CurrentParamSet $params) | Should -BeTrue
            }
        }

        It 'Should detect omitted cmdlets to load as a change' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.AuthenticationType          = 'ServicePrincipalWithSecret'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.RequestedAuthenticationType = 'ServicePrincipalWithSecret'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ApplicationId               = 'app-id'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.TenantId                    = 'tenant-id'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.ApplicationSecret           = 'secret'
                $Script:MSCloudLoginConnectionProfile.ExchangeOnline.CmdletsToLoad               = @('Get-Mailbox')

                $params = @{
                    Workload          = 'ExchangeOnline'
                    ApplicationId     = 'app-id'
                    TenantId          = 'tenant-id'
                    ApplicationSecret = 'secret'
                }
                (Compare-InputParametersForChange -CurrentParamSet $params) | Should -BeTrue
            }
        }

        It 'Should ignore a session parameter that the workload does not own' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AuthenticationType          = 'ServicePrincipalWithSecret'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.RequestedAuthenticationType = 'ServicePrincipalWithSecret'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ApplicationId               = 'app-id'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.TenantId                    = 'tenant-id'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ApplicationSecret           = 'secret'

                # EnableSearchOnlySession only exists on SecurityComplianceCenter. Supplying it
                # for another workload must not be reported as a change on every call.
                $params = @{
                    Workload                = 'MicrosoftGraph'
                    ApplicationId           = 'app-id'
                    TenantId                = 'tenant-id'
                    ApplicationSecret       = 'secret'
                    EnableSearchOnlySession = $true
                }
                (Compare-InputParametersForChange -CurrentParamSet $params) | Should -BeFalse
            }
        }
    }

    Context 'String comparison semantics' {
        It 'Should ignore a case-only change of an identifier' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.AdminAPI.AuthenticationType          = 'ServicePrincipalWithSecret'
                $Script:MSCloudLoginConnectionProfile.AdminAPI.RequestedAuthenticationType = 'ServicePrincipalWithSecret'
                $Script:MSCloudLoginConnectionProfile.AdminAPI.ApplicationId               = 'app-id'
                $Script:MSCloudLoginConnectionProfile.AdminAPI.TenantId                    = 'Tenant-Id'
                $Script:MSCloudLoginConnectionProfile.AdminAPI.ApplicationSecret           = 'secret'

                $params = @{
                    Workload          = 'AdminAPI'
                    ApplicationId     = 'APP-ID'
                    TenantId          = 'tenant-id'
                    ApplicationSecret = 'secret'
                }
                (Compare-InputParametersForChange -CurrentParamSet $params) | Should -BeFalse
            }
        }

        It 'Should detect a case-only change of a secret' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.AdminAPI.AuthenticationType          = 'ServicePrincipalWithSecret'
                $Script:MSCloudLoginConnectionProfile.AdminAPI.RequestedAuthenticationType = 'ServicePrincipalWithSecret'
                $Script:MSCloudLoginConnectionProfile.AdminAPI.ApplicationId               = 'app-id'
                $Script:MSCloudLoginConnectionProfile.AdminAPI.TenantId                    = 'tenant-id'
                $Script:MSCloudLoginConnectionProfile.AdminAPI.ApplicationSecret           = 'Secret'

                $params = @{
                    Workload          = 'AdminAPI'
                    ApplicationId     = 'app-id'
                    TenantId          = 'tenant-id'
                    ApplicationSecret = 'secret'
                }
                (Compare-InputParametersForChange -CurrentParamSet $params) | Should -BeTrue
            }
        }
    }

    Context 'Null and empty handling' {
        It 'Should treat empty string on the profile and absent parameter as equal' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.AdminAPI.AuthenticationType          = 'ServicePrincipalWithSecret'
                $Script:MSCloudLoginConnectionProfile.AdminAPI.RequestedAuthenticationType = 'ServicePrincipalWithSecret'
                $Script:MSCloudLoginConnectionProfile.AdminAPI.ApplicationId               = 'app-id'
                $Script:MSCloudLoginConnectionProfile.AdminAPI.TenantId                    = ''
                $Script:MSCloudLoginConnectionProfile.AdminAPI.ApplicationSecret           = 'secret'

                $params = @{
                    Workload          = 'AdminAPI'
                    ApplicationId     = 'app-id'
                    ApplicationSecret = 'secret'
                }
                (Compare-InputParametersForChange -CurrentParamSet $params) | Should -BeFalse
            }
        }
    }

    Context 'Input hashtable integrity' {
        It 'Should not mutate the provided parameter set' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $cred = New-Object PSCredential ('user@contoso.com', (ConvertTo-SecureString 'pwd' -AsPlainText -Force))
                $params = @{
                    Workload   = 'ExchangeOnline'
                    Credential = $cred
                }
                $null = Compare-InputParametersForChange -CurrentParamSet $params
                $params.ContainsKey('Credential') | Should -BeTrue
                $params.ContainsKey('Workload') | Should -BeTrue
                $params.ContainsKey('UserName') | Should -BeFalse
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
                Should -Invoke Add-MSCloudLoginAssistantEvent -ParameterFilter {
                    $Message -notlike '*super-secret*'
                }
            }
        }
    }

    Context 'Microsoft Graph special cases' {
        It 'Should not report a change for the auto-injected default Graph application id' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $cred = New-Object PSCredential ('user@contoso.com', (ConvertTo-SecureString 'pwd' -AsPlainText -Force))
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.AuthenticationType          = 'Credentials'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.RequestedAuthenticationType = 'Credentials'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.Credentials                 = $cred
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.ApplicationId               = '14d82eec-204b-4c2f-b7e8-296a70dab67e'
                $Script:MSCloudLoginConnectionProfile.MicrosoftGraph.TenantId                    = 'contoso.com'

                $params = @{
                    Workload   = 'MicrosoftGraph'
                    Credential = $cred
                }
                (Compare-InputParametersForChange -CurrentParamSet $params) | Should -BeFalse
            }
        }
    }
}

# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------
Describe 'Test-MSCloudLoginMFARequiredError' {

    Context 'Known MFA error codes' {
        It 'Should detect AADSTS50076 in the exception message' {
            InModuleScope 'MSCloudLoginAssistant' {
                $err = $null
                try { throw 'AADSTS50076: Due to a configuration change made by your administrator...' } catch { $err = $_ }
                (Test-MSCloudLoginMFARequiredError -ErrorRecord $err) | Should -BeTrue
            }
        }

        It 'Should detect the plain multi-factor wording' {
            InModuleScope 'MSCloudLoginAssistant' {
                $err = $null
                try { throw 'you must use multi-factor authentication to access this resource' } catch { $err = $_ }
                (Test-MSCloudLoginMFARequiredError -ErrorRecord $err) | Should -BeTrue
            }
        }

        It 'Should honor additional patterns' {
            InModuleScope 'MSCloudLoginAssistant' {
                $err = $null
                try { throw 'WAM Error 12345' } catch { $err = $_ }
                (Test-MSCloudLoginMFARequiredError -ErrorRecord $err) | Should -BeFalse
                (Test-MSCloudLoginMFARequiredError -ErrorRecord $err -AdditionalPatterns @('*WAM Error*')) | Should -BeTrue
            }
        }

        It 'Should not match unrelated errors' {
            InModuleScope 'MSCloudLoginAssistant' {
                $err = $null
                try { throw 'The sign-in name or password is incorrect' } catch { $err = $_ }
                (Test-MSCloudLoginMFARequiredError -ErrorRecord $err) | Should -BeFalse
            }
        }
    }
}

Describe 'Get-MSCloudLoginSPOUrlFromTenantId' {

    Context 'URL derivation per environment' {
        It 'Should derive commercial URLs' {
            InModuleScope 'MSCloudLoginAssistant' {
                $result = Get-MSCloudLoginSPOUrlFromTenantId -TenantId 'contoso.onmicrosoft.com' -EnvironmentName 'AzureCloud'
                $result.AdminUrl      | Should -Be 'https://contoso-admin.sharepoint.com'
                $result.ConnectionUrl | Should -Be 'https://contoso.sharepoint.com'
            }
        }

        It 'Should derive GCC High URLs with the .us suffix' {
            InModuleScope 'MSCloudLoginAssistant' {
                $result = Get-MSCloudLoginSPOUrlFromTenantId -TenantId 'contoso.onmicrosoft.com' -EnvironmentName 'AzureUSGovernment'
                $result.AdminUrl | Should -Be 'https://contoso-admin.sharepoint.us'
            }
        }

        It 'Should derive DoD URLs on the sharepoint-mil domain' {
            InModuleScope 'MSCloudLoginAssistant' {
                $result = Get-MSCloudLoginSPOUrlFromTenantId -TenantId 'contoso.onmicrosoft.com' -EnvironmentName 'AzureDOD'
                $result.AdminUrl | Should -Be 'https://contoso-admin.sharepoint-mil.us'
            }
        }

        It 'Should derive China URLs' {
            InModuleScope 'MSCloudLoginAssistant' {
                $result = Get-MSCloudLoginSPOUrlFromTenantId -TenantId 'contoso.partner.onmschina.cn' -EnvironmentName 'AzureChinaCloud'
                $result.AdminUrl | Should -Be 'https://contoso-admin.sharepoint.cn'
            }
        }

        It 'Should derive .onms. URLs on the .spo. domain' {
            InModuleScope 'MSCloudLoginAssistant' {
                $result = Get-MSCloudLoginSPOUrlFromTenantId -TenantId 'contoso.onms.fr' -EnvironmentName 'AzureFranceCloud'
                $result.AdminUrl | Should -Be 'https://contoso-admin.spo.fr'
            }
        }

        It 'Should throw for an unrecognized tenant format' {
            InModuleScope 'MSCloudLoginAssistant' {
                { Get-MSCloudLoginSPOUrlFromTenantId -TenantId 'contoso.com' -EnvironmentName 'AzureCloud' } | Should -Throw
            }
        }
    }
}

Describe 'Get-MSCloudLoginAccessTokenValue' {

    Context 'Token representations' {
        It 'Should pass through a plain string' {
            InModuleScope 'MSCloudLoginAssistant' {
                (Get-MSCloudLoginAccessTokenValue -Token 'plain-token') | Should -Be 'plain-token'
            }
        }

        It 'Should decrypt a SecureString' {
            InModuleScope 'MSCloudLoginAssistant' {
                $secure = ConvertTo-SecureString 'secure-token' -AsPlainText -Force
                (Get-MSCloudLoginAccessTokenValue -Token $secure) | Should -Be 'secure-token'
            }
        }

        It 'Should extract the password from a PSCredential' {
            InModuleScope 'MSCloudLoginAssistant' {
                $cred = New-Object PSCredential ('token', (ConvertTo-SecureString 'cred-token' -AsPlainText -Force))
                (Get-MSCloudLoginAccessTokenValue -Token $cred) | Should -Be 'cred-token'
            }
        }
    }
}

Describe 'Get-MSCloudLoginTenantDomainFromCredentials' {

    It 'Should return the domain part of a UPN' {
        InModuleScope 'MSCloudLoginAssistant' {
            $cred = New-Object PSCredential ('user@contoso.com', (ConvertTo-SecureString 'pwd' -AsPlainText -Force))
            (Get-MSCloudLoginTenantDomainFromCredentials -Credentials $cred) | Should -Be 'contoso.com'
        }
    }

    It 'Should throw when the user name is not a UPN' {
        InModuleScope 'MSCloudLoginAssistant' {
            $cred = New-Object PSCredential ('CONTOSO\user', (ConvertTo-SecureString 'pwd' -AsPlainText -Force))
            { Get-MSCloudLoginTenantDomainFromCredentials -Credentials $cred } | Should -Throw
        }
    }
}

Describe 'Test-MSCloudLoginConnectionReusable' {

    BeforeAll {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
        }
    }

    It 'Should return false when not connected' {
        InModuleScope 'MSCloudLoginAssistant' {
            $workloadProfile = New-Object AdminAPI
            (Test-MSCloudLoginConnectionReusable -WorkloadProfile $workloadProfile -Source 'Test') | Should -BeFalse
        }
    }

    It 'Should reset a connection without a timestamp' {
        InModuleScope 'MSCloudLoginAssistant' {
            $workloadProfile = New-Object AdminAPI
            $workloadProfile.Connected = $true
            (Test-MSCloudLoginConnectionReusable -WorkloadProfile $workloadProfile -Source 'Test') | Should -BeFalse
            $workloadProfile.Connected | Should -BeFalse
        }
    }

    It 'Should reset an expired token-based connection' {
        InModuleScope 'MSCloudLoginAssistant' {
            $workloadProfile = New-Object AdminAPI
            $workloadProfile.AuthenticationType = 'ServicePrincipalWithSecret'
            $workloadProfile.CompleteConnection()
            $workloadProfile.ConnectedDateTime = [System.DateTime]::Now.AddMinutes(-60).ToString()
            (Test-MSCloudLoginConnectionReusable -WorkloadProfile $workloadProfile -Source 'Test') | Should -BeFalse
            $workloadProfile.Connected | Should -BeFalse
        }
    }

    It 'Should reuse a fresh token-based connection' {
        InModuleScope 'MSCloudLoginAssistant' {
            $workloadProfile = New-Object AdminAPI
            $workloadProfile.AuthenticationType = 'ServicePrincipalWithSecret'
            $workloadProfile.CompleteConnection()
            (Test-MSCloudLoginConnectionReusable -WorkloadProfile $workloadProfile -Source 'Test') | Should -BeTrue
        }
    }

    It 'Should reset the connection when the probe returns null' {
        InModuleScope 'MSCloudLoginAssistant' {
            $workloadProfile = New-Object AdminAPI
            $workloadProfile.AuthenticationType = 'ServicePrincipalWithThumbprint'
            $workloadProfile.CompleteConnection()
            (Test-MSCloudLoginConnectionReusable -WorkloadProfile $workloadProfile -ProbeScript { $null } -Source 'Test') | Should -BeFalse
            $workloadProfile.Connected | Should -BeFalse
        }
    }
}

# ---------------------------------------------------------------------------
# Endpoint data table
# ---------------------------------------------------------------------------
Describe 'Get-MSCloudLoginEndpointInfo' {

    Context 'Endpoint snapshots' {
        It 'Should resolve <Workload>/<Environment> endpoints' -TestCases @(
            @{ Workload = 'AdminAPI'; Environment = 'AzureCloud'; Property = 'AuthorizationUrl'; Expected = 'https://login.microsoftonline.com' }
            @{ Workload = 'AdminAPI'; Environment = 'AzureDOD'; Property = 'AuthorizationUrl'; Expected = 'https://login.microsoftonline.us' }
            @{ Workload = 'AzureDevOPS'; Environment = 'AzureDOD'; Property = 'HostUrl'; Expected = 'https://dev.azure.us' }
            @{ Workload = 'DefenderForEndpoint'; Environment = 'AzureUSGovernment'; Property = 'HostUrl'; Expected = 'https://api-gcc.securitycenter.microsoft.us' }
            @{ Workload = 'Fabric'; Environment = 'AzureCloud'; Property = 'Scope'; Expected = 'https://api.fabric.microsoft.com/.default' }
            @{ Workload = 'Licensing'; Environment = 'AzureCloud'; Property = 'HostUrl'; Expected = 'https://licensing.m365.microsoft.com' }
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

        It 'Should replace placeholders' {
            InModuleScope 'MSCloudLoginAssistant' {
                $result = Get-MSCloudLoginEndpointInfo -Workload 'AdminAPI' -EnvironmentName 'AzureCloud' -Replacements @{ Resource = 'my-resource' }
                $result.Scope | Should -Be 'my-resource/.default'
            }
        }

        It 'Should fall back to the default entry for unknown environments' {
            InModuleScope 'MSCloudLoginAssistant' {
                $result = Get-MSCloudLoginEndpointInfo -Workload 'Fabric' -EnvironmentName 'SomethingElse'
                $result.AuthorizationUrl | Should -Be 'https://login.microsoftonline.com'
            }
        }

        It 'Should throw for an unknown workload' {
            InModuleScope 'MSCloudLoginAssistant' {
                { Get-MSCloudLoginEndpointInfo -Workload 'DoesNotExist' -EnvironmentName 'AzureCloud' } | Should -Throw
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

AfterAll {
    Remove-Module MSCloudLoginAssistant
}
