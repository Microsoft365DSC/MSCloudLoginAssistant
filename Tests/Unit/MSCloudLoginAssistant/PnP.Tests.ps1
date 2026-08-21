#Requires -Modules Pester

Describe 'Connect-MSCloudLoginPnP' {
    BeforeAll {
        # Plain function stubs keep command resolution off the real SDK modules,
        # which would otherwise be discovered and imported on first use.
        Import-Module ./Tests/Unit/Stubs/Stubs.psm1 -Force -Global -WarningAction SilentlyContinue
        Import-Module ./Modules/MSCloudLoginAssistant/MSCloudLoginAssistant.psd1 -Force

        # Compile and instantiate the workload classes once here so that the cost
        # does not show up inside the first test of this file.
        InModuleScope 'MSCloudLoginAssistant' {
            $null = New-Object MSCloudLoginConnectionProfile
        }
    }

    Context 'When connecting with AccessTokens' {
        It 'Should call Connect-PnPOnline with AccessToken' {
            InModuleScope 'MSCloudLoginAssistant' {
                Import-Module ./Tests/Unit/Stubs/Stubs.psm1 -Force
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-Module -MockWith {
                    param ($Name, $ListAvailable)
                    if ($ListAvailable -and $Name -eq 'PnP.PowerShell') {
                        return @(
                            [pscustomobject]@{
                                Name = 'PnP.PowerShell'
                                Version = '1.10.0'
                                CompatiblePSEditions = @('Desktop', 'Core')
                            }
                        )
                    }
                    if ($Name -eq 'PnP.PowerShell') { return $null }
                    return @()
                }
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Get-MSCloudLoginSPOUrlFromTenantId -MockWith { return @{ ConnectionUrl = 'https://contoso.sharepoint.com'; AdminUrl = 'https://contoso-admin.sharepoint.com' } }

                $secPwd = ConvertTo-SecureString 'password' -AsPlainText -Force

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl = 'https://contoso-admin.sharepoint.com'
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'AccessTokens'
                $Script:MSCloudLoginConnectionProfile.PnP.AccessTokens = @( $secPwd )
                $Script:MSCloudLoginConnectionProfile.PnP.TenantId = 'tenant-id'
                $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment = 'Production'
                $Script:MSCloudLoginConnectionProfile.PnP.EnvironmentName = 'AzureCloud'
                $Script:MSCloudLoginConnectionProfile.PnP.Connected = $false

                Connect-MSCloudLoginPnP

                Should -Invoke Connect-PnPOnline -ParameterFilter {
                    $Url -eq 'https://contoso-admin.sharepoint.com' -and
                    $AccessToken -eq $secPwd -and
                    $AzureEnvironment -eq 'Production'
                }
            }
        }
    }

    Context 'When the connection is still reusable' {
        It 'Should return early without connecting again' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $true }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.Connected = $true

                { Connect-MSCloudLoginPnP } | Should -Not -Throw

                Should -Invoke Add-MSCloudLoginAssistantEvent -ParameterFilter {
                    $Message -like '*Already connected to PnP*'
                }
                Should -Invoke Connect-PnPOnline -Exactly 0
            }
        }
    }

    Context 'When importing the Graph module as a workaround' {
        It 'Should log a warning when the import fails' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-Module -MockWith {
                    param ($Name)
                    if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                    if ($Name -eq 'PnP.PowerShell') { return [pscustomobject]@{ Name = 'PnP.PowerShell' } }
                    return $null
                }
                Mock -CommandName Import-Module -MockWith { throw 'the module could not be loaded' }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'AccessTokens'
                $Script:MSCloudLoginConnectionProfile.PnP.AccessTokens = @('raw-token')
                $Script:MSCloudLoginConnectionProfile.PnP.TenantId = 'contoso.onmicrosoft.com'
                $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl = 'https://contoso-admin.sharepoint.com'

                { Connect-MSCloudLoginPnP } | Should -Not -Throw

                Should -Invoke Add-MSCloudLoginAssistantEvent -ParameterFilter {
                    $Message -like '*Failed to import Microsoft.Graph.Authentication*' -and $EntryType -eq 'Warning'
                }
            }
        }
    }

    Context 'When loading the PnP.PowerShell module on PowerShell Core' {
        It 'Should load the Desktop edition through Windows PowerShell when only v1 is available' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-Module -MockWith {
                    param ($Name, $ListAvailable)
                    if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                    if ($Name -ne 'PnP.PowerShell') { return $null }
                    if (-not $ListAvailable) { return $null }
                    return @(
                        [pscustomobject]@{
                            Name                 = 'PnP.PowerShell'
                            Version              = [version]'1.10.0'
                            CompatiblePSEditions = @('Core', 'Desktop')
                        }
                    )
                }
                Mock -CommandName Import-Module -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'AccessTokens'
                $Script:MSCloudLoginConnectionProfile.PnP.AccessTokens = @('raw-token')
                $Script:MSCloudLoginConnectionProfile.PnP.TenantId = 'contoso.onmicrosoft.com'
                $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl = 'https://contoso-admin.sharepoint.com'

                { Connect-MSCloudLoginPnP } | Should -Not -Throw

                Should -Invoke Import-Module -Exactly 1 -ParameterFilter {
                    $Name -eq 'PnP.PowerShell' -and
                    $RequiredVersion -eq [version]'1.10.0' -and
                    $UseWindowsPowerShell.IsPresent
                }
            }
        }

        It 'Should explain that the Windows PowerShell installation is missing when no Desktop edition exists' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-Module -MockWith {
                    param ($Name, $ListAvailable)
                    if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                    if ($Name -ne 'PnP.PowerShell') { return $null }
                    if (-not $ListAvailable) { return $null }
                    return @(
                        [pscustomobject]@{
                            Name                 = 'PnP.PowerShell'
                            Version              = [version]'1.10.0'
                            CompatiblePSEditions = @('Core')
                        }
                    )
                }
                Mock -CommandName Import-Module -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'AccessTokens'
                $Script:MSCloudLoginConnectionProfile.PnP.AccessTokens = @('raw-token')
                $Script:MSCloudLoginConnectionProfile.PnP.TenantId = 'contoso.onmicrosoft.com'
                $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl = 'https://contoso-admin.sharepoint.com'

                { Connect-MSCloudLoginPnP } |
                    Should -Throw '*Powershell 7+ was detected*-UseWindowsPowerShell*not installed for Windows PowerShell*'
            }
        }

        It 'Should not reload the module when PnP.PowerShell is already loaded' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-Module -MockWith {
                    param ($Name, $ListAvailable)
                    if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                    if ($Name -eq 'PnP.PowerShell') { return [pscustomobject]@{ Name = 'PnP.PowerShell' } }
                    return $null
                }
                Mock -CommandName Import-Module -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'AccessTokens'
                $Script:MSCloudLoginConnectionProfile.PnP.AccessTokens = @('raw-token')
                $Script:MSCloudLoginConnectionProfile.PnP.TenantId = 'contoso.onmicrosoft.com'
                $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl = 'https://contoso-admin.sharepoint.com'

                { Connect-MSCloudLoginPnP } | Should -Not -Throw

                Should -Invoke Import-Module -Exactly 0 -ParameterFilter {
                    $Name -eq 'PnP.PowerShell'
                }
            }
        }
    }

    Context 'When resolving the connection URL' {
        It 'Should adopt the admin URL as connection URL when only the admin URL is set' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-Module -MockWith {
                    param ($Name)
                    if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                    if ($Name -eq 'PnP.PowerShell') { return [pscustomobject]@{ Name = 'PnP.PowerShell' } }
                    return $null
                }
                Mock -CommandName Import-Module -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'AccessTokens'
                $Script:MSCloudLoginConnectionProfile.PnP.AccessTokens = @('raw-token')
                $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl = 'https://contoso-admin.sharepoint.com'

                { Connect-MSCloudLoginPnP } | Should -Not -Throw

                $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl | Should -Be 'https://contoso-admin.sharepoint.com'
                Should -Invoke Connect-PnPOnline -Exactly 1
            }
        }

        It 'Should discover the admin URL through Graph when connecting with credentials' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-Module -MockWith {
                    param ($Name)
                    if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                    if ($Name -eq 'PnP.PowerShell') { return [pscustomobject]@{ Name = 'PnP.PowerShell' } }
                    return $null
                }
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Get-SPOAdminUrl -MockWith { return 'https://contoso-admin.sharepoint.com' }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'Credentials'
                $Script:MSCloudLoginConnectionProfile.PnP.Credentials =
                    New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))

                { Connect-MSCloudLoginPnP } | Should -Not -Throw

                $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl | Should -Be 'https://contoso-admin.sharepoint.com'
                $Script:MSCloudLoginConnectionProfile.PnP.Connected | Should -BeTrue
                Should -Invoke Get-SPOAdminUrl -Exactly 1
                Should -Invoke Connect-PnPOnline -Exactly 1 -ParameterFilter {
                    $Url -eq 'https://contoso-admin.sharepoint.com' -and
                    $ClientId -eq '9bc3ab49-b65d-410a-85ad-de819febfddc'
                }
            }
        }

        It 'Should fail with a clear message when the admin URL cannot be resolved' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-Module -MockWith {
                    param ($Name)
                    if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                    if ($Name -eq 'PnP.PowerShell') { return [pscustomobject]@{ Name = 'PnP.PowerShell' } }
                    return $null
                }
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Get-SPOAdminUrl -MockWith { return '' }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'Credentials'
                $Script:MSCloudLoginConnectionProfile.PnP.Credentials =
                    New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))

                { Connect-MSCloudLoginPnP } | Should -Throw '*Unable to retrieve SharePoint Admin Url*'
            }
        }

        It 'Should derive both URLs from the tenant id for non credential authentication' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-Module -MockWith {
                    param ($Name)
                    if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                    if ($Name -eq 'PnP.PowerShell') { return [pscustomobject]@{ Name = 'PnP.PowerShell' } }
                    return $null
                }
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Get-MSCloudLoginSPOUrlFromTenantId -MockWith {
                    return @{
                        AdminUrl      = 'https://contoso-admin.sharepoint.com'
                        ConnectionUrl = 'https://contoso.sharepoint.com'
                    }
                }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'ServicePrincipalWithSecret'
                $Script:MSCloudLoginConnectionProfile.PnP.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.PnP.ApplicationSecret = 'secret'
                $Script:MSCloudLoginConnectionProfile.PnP.TenantId = 'contoso.onmicrosoft.com'
                $Script:MSCloudLoginConnectionProfile.PnP.EnvironmentName = 'AzureCloud'
                $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment = 'Production'

                { Connect-MSCloudLoginPnP } | Should -Not -Throw

                $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl | Should -Be 'https://contoso-admin.sharepoint.com'
                Should -Invoke Get-MSCloudLoginSPOUrlFromTenantId -Exactly 1 -ParameterFilter {
                    $TenantId -eq 'contoso.onmicrosoft.com' -and $EnvironmentName -eq 'AzureCloud'
                }
                Should -Invoke Connect-PnPOnline -Exactly 1 -ParameterFilter {
                    $ClientSecret -eq 'secret' -and $null -ne $ClientId
                }
            }
        }

        It 'Should backfill the admin URL from an explicit connection URL' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-Module -MockWith {
                    param ($Name)
                    if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                    if ($Name -eq 'PnP.PowerShell') { return [pscustomobject]@{ Name = 'PnP.PowerShell' } }
                    return $null
                }
                Mock -CommandName Import-Module -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'AccessTokens'
                $Script:MSCloudLoginConnectionProfile.PnP.AccessTokens = @('raw-token')
                $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl = 'https://contoso-admin.sharepoint.com'

                { Connect-MSCloudLoginPnP } | Should -Not -Throw

                $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl | Should -Be 'https://contoso-admin.sharepoint.com'
            }
        }
    }

    Context 'When connecting with ServicePrincipalWithThumbprint' {
        It 'Should acquire a local token when scope and token URL are configured' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-Module -MockWith {
                    param ($Name)
                    if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                    if ($Name -eq 'PnP.PowerShell') { return [pscustomobject]@{ Name = 'PnP.PowerShell' } }
                    return $null
                }
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Get-MSCloudLoginAccessToken -MockWith { return 'locally-issued-token' }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'ServicePrincipalWithThumbprint'
                $Script:MSCloudLoginConnectionProfile.PnP.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.PnP.TenantId = 'contoso.onmicrosoft.com'
                $Script:MSCloudLoginConnectionProfile.PnP.CertificateThumbprint = 'thumbprint'
                $Script:MSCloudLoginConnectionProfile.PnP.AuthorizationUrl = 'https://login.contoso.local'
                $Script:MSCloudLoginConnectionProfile.PnP.Scope = 'https://contoso.sharepoint.com/.default'
                $Script:MSCloudLoginConnectionProfile.PnP.TokenUrl = 'https://login.contoso.local/oauth2/v2.0/token'
                $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl = 'https://contoso-admin.sharepoint.com'
                $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment = 'Production'

                { Connect-MSCloudLoginPnP } | Should -Not -Throw

                $Script:MSCloudLoginConnectionProfile.PnP.Connected | Should -BeTrue
                Should -Invoke Get-MSCloudLoginAccessToken -Exactly 1 -ParameterFilter {
                    $ApplicationId -eq 'app-id' -and $CertificateThumbprint -eq 'thumbprint'
                }
                Should -Invoke Connect-PnPOnline -Exactly 1 -ParameterFilter {
                    $AccessToken -eq 'locally-issued-token'
                }
            }
        }

        It 'Should connect by thumbprint and environment name' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-Module -MockWith {
                    param ($Name)
                    if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                    if ($Name -eq 'PnP.PowerShell') { return [pscustomobject]@{ Name = 'PnP.PowerShell' } }
                    return $null
                }
                Mock -CommandName Import-Module -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'ServicePrincipalWithThumbprint'
                $Script:MSCloudLoginConnectionProfile.PnP.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.PnP.TenantId = 'contoso.onmicrosoft.com'
                $Script:MSCloudLoginConnectionProfile.PnP.CertificateThumbprint = 'thumbprint'
                $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl = 'https://contoso-admin.sharepoint.com'
                $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment = 'Production'

                { Connect-MSCloudLoginPnP } | Should -Not -Throw

                $Script:MSCloudLoginConnectionProfile.PnP.Connected | Should -BeTrue
                Should -Invoke Connect-PnPOnline -Exactly 1 -ParameterFilter {
                    $ClientId -eq 'app-id' -and
                    $Tenant -eq 'contoso.onmicrosoft.com' -and
                    $Thumbprint -eq 'thumbprint' -and
                    $AzureEnvironment -eq 'Production'
                }
            }
        }

        It 'Should pass the custom endpoints when the environment is custom' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-Module -MockWith {
                    param ($Name)
                    if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                    if ($Name -eq 'PnP.PowerShell') { return [pscustomobject]@{ Name = 'PnP.PowerShell' } }
                    return $null
                }
                Mock -CommandName Import-Module -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'ServicePrincipalWithThumbprint'
                $Script:MSCloudLoginConnectionProfile.PnP.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.PnP.TenantId = 'contoso.local'
                $Script:MSCloudLoginConnectionProfile.PnP.CertificateThumbprint = 'thumbprint'
                $Script:MSCloudLoginConnectionProfile.PnP.EndPoints = @{
                    AzureADLoginEndPoint   = 'https://login.contoso.local'
                    MicrosoftGraphEndPoint = 'https://graph.contoso.local'
                }
                $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl = 'https://contoso-admin.sharepoint.contoso.local'
                $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment = 'Custom'

                { Connect-MSCloudLoginPnP } | Should -Not -Throw

                Should -Invoke Connect-PnPOnline -Exactly 1 -ParameterFilter {
                    $AzureEnvironment -eq 'Custom' -and
                    $AzureADLoginEndPoint -eq 'https://login.contoso.local' -and
                    $MicrosoftGraphEndPoint -eq 'https://graph.contoso.local'
                }
            }
        }

        It 'Should pass the custom endpoints when connecting through the admin URL in a custom environment' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-Module -MockWith {
                    param ($Name)
                    if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                    if ($Name -eq 'PnP.PowerShell') { return [pscustomobject]@{ Name = 'PnP.PowerShell' } }
                    return $null
                }
                Mock -CommandName Import-Module -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'ServicePrincipalWithThumbprint'
                $Script:MSCloudLoginConnectionProfile.PnP.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.PnP.TenantId = 'contoso.local'
                $Script:MSCloudLoginConnectionProfile.PnP.CertificateThumbprint = 'thumbprint'
                $Script:MSCloudLoginConnectionProfile.PnP.EndPoints = @{
                    AzureADLoginEndPoint   = 'https://login.contoso.local'
                    MicrosoftGraphEndPoint = 'https://graph.contoso.local'
                }
                # Leave ConnectionUrl empty so that it is adopted from AdminUrl below.
                $Script:MSCloudLoginConnectionProfile.PnP.AdminUrl = 'https://contoso-admin.sharepoint.contoso.local'
                $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment = 'Custom'

                { Connect-MSCloudLoginPnP } | Should -Not -Throw

                Should -Invoke Connect-PnPOnline -Exactly 1 -ParameterFilter {
                    $Url -eq 'https://contoso-admin.sharepoint.contoso.local' -and
                    $AzureADLoginEndPoint -eq 'https://login.contoso.local'
                }
            }
        }
    }

    Context 'When connecting with ServicePrincipalWithPath' {
        It 'Should pass the certificate path and password to Connect-PnPOnline' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-Module -MockWith {
                    param ($Name)
                    if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                    if ($Name -eq 'PnP.PowerShell') { return [pscustomobject]@{ Name = 'PnP.PowerShell' } }
                    return $null
                }
                Mock -CommandName Import-Module -MockWith { }

                $certificatePassword = ConvertTo-SecureString 'cert-password' -AsPlainText -Force

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'ServicePrincipalWithPath'
                $Script:MSCloudLoginConnectionProfile.PnP.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.PnP.TenantId = 'contoso.onmicrosoft.com'
                $Script:MSCloudLoginConnectionProfile.PnP.CertificatePath = 'C:\certs\contoso.pfx'
                $Script:MSCloudLoginConnectionProfile.PnP.CertificatePassword = $certificatePassword
                $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl = 'https://contoso-admin.sharepoint.com'
                $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment = 'Production'

                { Connect-MSCloudLoginPnP } | Should -Not -Throw

                $Script:MSCloudLoginConnectionProfile.PnP.Connected | Should -BeTrue
                Should -Invoke Connect-PnPOnline -Exactly 1 -ParameterFilter {
                    $ClientId -eq 'app-id' -and
                    $CertificatePath -eq 'C:\certs\contoso.pfx' -and
                    $CertificatePassword -is [System.Security.SecureString]
                }
            }
        }
    }

    Context 'When connecting with ServicePrincipalWithSecret' {
        It 'Should connect with the client secret' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-Module -MockWith {
                    param ($Name)
                    if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                    if ($Name -eq 'PnP.PowerShell') { return [pscustomobject]@{ Name = 'PnP.PowerShell' } }
                    return $null
                }
                Mock -CommandName Import-Module -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'ServicePrincipalWithSecret'
                $Script:MSCloudLoginConnectionProfile.PnP.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.PnP.ApplicationSecret = 'secret'
                $Script:MSCloudLoginConnectionProfile.PnP.TenantId = 'contoso.onmicrosoft.com'
                $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl = 'https://contoso-admin.sharepoint.com'
                $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment = 'Production'

                { Connect-MSCloudLoginPnP } | Should -Not -Throw

                $Script:MSCloudLoginConnectionProfile.PnP.Connected | Should -BeTrue
                Should -Invoke Connect-PnPOnline -Exactly 1 -ParameterFilter {
                    $ClientSecret -eq 'secret' -and $Url -eq 'https://contoso-admin.sharepoint.com'
                }
            }
        }

        It 'Should honour a forced refresh even without a connection URL' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-Module -MockWith {
                    param ($Name)
                    if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                    if ($Name -eq 'PnP.PowerShell') { return [pscustomobject]@{ Name = 'PnP.PowerShell' } }
                    return $null
                }
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Get-MSCloudLoginSPOUrlFromTenantId -MockWith {
                    return @{
                        AdminUrl      = 'https://contoso-admin.sharepoint.com'
                        ConnectionUrl = 'https://contoso.sharepoint.com'
                    }
                }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'ServicePrincipalWithSecret'
                $Script:MSCloudLoginConnectionProfile.PnP.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.PnP.ApplicationSecret = 'secret'
                $Script:MSCloudLoginConnectionProfile.PnP.TenantId = 'contoso.onmicrosoft.com'
                $Script:MSCloudLoginConnectionProfile.PnP.EnvironmentName = 'AzureCloud'
                $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment = 'Production'

                { Connect-MSCloudLoginPnP -ForceRefreshConnection:$true } | Should -Not -Throw

                $Script:MSCloudLoginConnectionProfile.PnP.Connected | Should -BeTrue
            }
        }
    }

    Context 'When connecting with CredentialsWithTenantId' {
        It 'Should refuse the combination of credentials and tenant id' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-Module -MockWith {
                    param ($Name)
                    if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                    if ($Name -eq 'PnP.PowerShell') { return [pscustomobject]@{ Name = 'PnP.PowerShell' } }
                    return $null
                }
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Get-MSCloudLoginSPOUrlFromTenantId -MockWith {
                    return @{
                        AdminUrl      = 'https://contoso-admin.sharepoint.com'
                        ConnectionUrl = 'https://contoso.sharepoint.com'
                    }
                }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'CredentialsWithTenantId'
                $Script:MSCloudLoginConnectionProfile.PnP.Credentials =
                    New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))
                $Script:MSCloudLoginConnectionProfile.PnP.TenantId = 'contoso.onmicrosoft.com'
                $Script:MSCloudLoginConnectionProfile.PnP.EnvironmentName = 'AzureCloud'

                { Connect-MSCloudLoginPnP } |
                    Should -Throw '*You cannot specify TenantId with Credentials when connecting to PnP*'
            }
        }
    }

    Context 'When connecting with CredentialsWithApplicationId' {
        It 'Should pass the credential together with the application id' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-Module -MockWith {
                    param ($Name)
                    if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                    if ($Name -eq 'PnP.PowerShell') { return [pscustomobject]@{ Name = 'PnP.PowerShell' } }
                    return $null
                }
                Mock -CommandName Import-Module -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'CredentialsWithApplicationId'
                $Script:MSCloudLoginConnectionProfile.PnP.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.PnP.Credentials =
                    New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))
                $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl = 'https://contoso-admin.sharepoint.com'
                $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment = 'Production'

                { Connect-MSCloudLoginPnP } | Should -Not -Throw

                $Script:MSCloudLoginConnectionProfile.PnP.Connected | Should -BeTrue
                Should -Invoke Connect-PnPOnline -Exactly 1 -ParameterFilter {
                    $ClientId -eq 'app-id' -and $null -ne $Credentials
                }
            }
        }
    }

    Context 'When connecting with Credentials' {
        It 'Should connect with the management shell client id' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-Module -MockWith {
                    param ($Name)
                    if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                    if ($Name -eq 'PnP.PowerShell') { return [pscustomobject]@{ Name = 'PnP.PowerShell' } }
                    return $null
                }
                Mock -CommandName Import-Module -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'Credentials'
                $Script:MSCloudLoginConnectionProfile.PnP.Credentials =
                    New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))
                $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl = 'https://contoso-admin.sharepoint.com'
                $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment = 'Production'

                { Connect-MSCloudLoginPnP } | Should -Not -Throw

                $Script:MSCloudLoginConnectionProfile.PnP.Connected | Should -BeTrue
                Should -Invoke Connect-PnPOnline -Exactly 1 -ParameterFilter {
                    $null -ne $Credentials -and $ClientId -eq '9bc3ab49-b65d-410a-85ad-de819febfddc'
                }
            }
        }
    }

    Context 'When connecting with Identity' {
        It 'Should request a managed identity token for the connection URL' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-Module -MockWith {
                    param ($Name)
                    if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                    if ($Name -eq 'PnP.PowerShell') { return [pscustomobject]@{ Name = 'PnP.PowerShell' } }
                    return $null
                }
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Get-AuthToken -MockWith { return 'managed-identity-token' }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'Identity'
                $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl = 'https://contoso-admin.sharepoint.com'
                $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment = 'Production'

                { Connect-MSCloudLoginPnP } | Should -Not -Throw

                $Script:MSCloudLoginConnectionProfile.PnP.Connected | Should -BeTrue
                Should -Invoke Get-AuthToken -Exactly 1 -ParameterFilter {
                    $Resource -eq 'https://contoso-admin.sharepoint.com' -and $Identity.IsPresent
                }
                Should -Invoke Connect-PnPOnline -Exactly 1 -ParameterFilter {
                    $AccessToken -eq 'managed-identity-token'
                }
            }
        }
    }

    Context 'When the authentication type is not supported' {
        It 'Should throw an error naming the unsupported type' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-Module -MockWith {
                    param ($Name)
                    if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                    if ($Name -eq 'PnP.PowerShell') { return [pscustomobject]@{ Name = 'PnP.PowerShell' } }
                    return $null
                }
                Mock -CommandName Import-Module -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'Interactive'
                $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl = 'https://contoso-admin.sharepoint.com'

                { Connect-MSCloudLoginPnP } |
                    Should -Throw "*Authentication type 'Interactive' is not supported for workload 'PnP'*"
            }
        }
    }

    Context 'When the sign-in requires MFA' {
        It 'Should retry interactively and mark the connection as multi-factor authenticated' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:pnpConnectCalls = 0
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $false }
                Mock -CommandName Get-Module -MockWith {
                    param ($Name)
                    if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                    if ($Name -eq 'PnP.PowerShell') { return [pscustomobject]@{ Name = 'PnP.PowerShell' } }
                    return $null
                }
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Connect-PnPOnline -MockWith {
                    $Script:pnpConnectCalls++
                    if ($Script:pnpConnectCalls -eq 1)
                    {
                        throw 'AADSTS50076: multi-factor authentication is required'
                    }
                }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'AccessTokens'
                $Script:MSCloudLoginConnectionProfile.PnP.AccessTokens = @('raw-token')
                $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl = 'https://contoso-admin.sharepoint.com'
                $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment = 'Production'

                { Connect-MSCloudLoginPnP } | Should -Not -Throw

                $Script:MSCloudLoginConnectionProfile.PnP.Connected | Should -BeTrue
                $Script:MSCloudLoginConnectionProfile.PnP.MultiFactorAuthentication | Should -BeTrue
                $Script:pnpConnectCalls | Should -Be 2
            }
        }

        It 'Should fall back to the web login when the interactive attempt fails' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:pnpConnectCalls = 0
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $false }
                Mock -CommandName Get-Module -MockWith {
                    param ($Name)
                    if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                    if ($Name -eq 'PnP.PowerShell') { return [pscustomobject]@{ Name = 'PnP.PowerShell' } }
                    return $null
                }
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Connect-PnPOnline -MockWith {
                    $Script:pnpConnectCalls++
                    switch ($Script:pnpConnectCalls)
                    {
                        1 { throw 'AADSTS50076: multi-factor authentication is required' }
                        2 { throw 'the interactive window was dismissed' }
                        default { }
                    }
                }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'AccessTokens'
                $Script:MSCloudLoginConnectionProfile.PnP.AccessTokens = @('raw-token')
                $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl = 'https://contoso-admin.sharepoint.com'

                { Connect-MSCloudLoginPnP } | Should -Not -Throw

                $Script:MSCloudLoginConnectionProfile.PnP.Connected | Should -BeTrue
                $Script:pnpConnectCalls | Should -Be 3
                Should -Invoke Connect-PnPOnline -Exactly 1 -ParameterFilter {
                    $UseWebLogin.IsPresent
                }
            }
        }

        It 'Should surface the failure when every fallback fails' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $false }
                Mock -CommandName Get-Module -MockWith {
                    param ($Name)
                    if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                    if ($Name -eq 'PnP.PowerShell') { return [pscustomobject]@{ Name = 'PnP.PowerShell' } }
                    return $null
                }
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Connect-PnPOnline -MockWith { throw 'AADSTS50076: multi-factor authentication is required' }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $pnpProfile = $Script:MSCloudLoginConnectionProfile.PnP
                $pnpProfile.AuthenticationType = 'AccessTokens'
                $pnpProfile.AccessTokens = @('raw-token')
                $pnpProfile.ConnectionUrl = 'https://contoso-admin.sharepoint.com'

                { Connect-MSCloudLoginPnP } | Should -Throw '*multi-factor authentication is required*'

                $pnpProfile.Connected | Should -BeFalse
                Should -Invoke Add-MSCloudLoginAssistantEvent -ParameterFilter {
                    $Message -like '*Failed to connect to PnP interactively after MFA-required error*' -and $EntryType -eq 'Error'
                }
            }
        }
    }

    Context 'When the credentials are rejected because of MFA' {
        It 'Should retry interactively on the password mismatch error' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:pnpConnectCalls = 0
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $false }
                Mock -CommandName Get-Module -MockWith {
                    param ($Name)
                    if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                    if ($Name -eq 'PnP.PowerShell') { return [pscustomobject]@{ Name = 'PnP.PowerShell' } }
                    return $null
                }
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Connect-PnPOnline -MockWith {
                    $Script:pnpConnectCalls++
                    if ($Script:pnpConnectCalls -eq 1)
                    {
                        throw 'The sign-in name or password does not match one in the Microsoft account system.'
                    }
                }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'Credentials'
                $Script:MSCloudLoginConnectionProfile.PnP.Credentials =
                    New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))
                $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl = 'https://contoso-admin.sharepoint.com'
                $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment = 'Production'

                { Connect-MSCloudLoginPnP } | Should -Not -Throw

                $Script:MSCloudLoginConnectionProfile.PnP.Connected | Should -BeTrue
                $Script:MSCloudLoginConnectionProfile.PnP.MultiFactorAuthentication | Should -BeTrue
            }
        }

        It 'Should surface the failure when the interactive retry also fails' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $false }
                Mock -CommandName Get-Module -MockWith {
                    param ($Name)
                    if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                    if ($Name -eq 'PnP.PowerShell') { return [pscustomobject]@{ Name = 'PnP.PowerShell' } }
                    return $null
                }
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Connect-PnPOnline -MockWith {
                    throw 'The sign-in name or password does not match one in the Microsoft account system.'
                }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $pnpProfile = $Script:MSCloudLoginConnectionProfile.PnP
                $pnpProfile.AuthenticationType = 'Credentials'
                $pnpProfile.Credentials =
                    New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))
                $pnpProfile.ConnectionUrl = 'https://contoso-admin.sharepoint.com'
                $pnpProfile.PnPAzureEnvironment = 'Production'

                { Connect-MSCloudLoginPnP } | Should -Throw '*sign-in name or password does not match*'

                $pnpProfile.Connected | Should -BeFalse
                Should -Invoke Add-MSCloudLoginAssistantEvent -ParameterFilter {
                    $Message -like '*Failed to connect to PnP interactively:*' -and $EntryType -eq 'Error'
                }
            }
        }
    }

    Context 'When the application has not been consented' {
        It 'Should register the management shell access and reconnect via web login' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:pnpConnectCalls = 0
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $false }
                Mock -CommandName Get-Module -MockWith {
                    param ($Name)
                    if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                    if ($Name -eq 'PnP.PowerShell') { return [pscustomobject]@{ Name = 'PnP.PowerShell' } }
                    return $null
                }
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Register-PnPManagementShellAccess -MockWith { }
                Mock -CommandName Connect-PnPOnline -MockWith {
                    $Script:pnpConnectCalls++
                    if ($Script:pnpConnectCalls -eq 1)
                    {
                        throw 'AADSTS65001: The user or administrator has not consented to use the application with ID abc'
                    }
                }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'AccessTokens'
                $Script:MSCloudLoginConnectionProfile.PnP.AccessTokens = @('raw-token')
                $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl = 'https://contoso-admin.sharepoint.com'

                { Connect-MSCloudLoginPnP } | Should -Not -Throw

                $Script:MSCloudLoginConnectionProfile.PnP.Connected | Should -BeTrue
                Should -Invoke Register-PnPManagementShellAccess -Exactly 1
                Should -Invoke Connect-PnPOnline -Exactly 1 -ParameterFilter {
                    $UseWebLogin.IsPresent
                }
            }
        }

        It 'Should wrap the error when the consent flow fails' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $false }
                Mock -CommandName Get-Module -MockWith {
                    param ($Name)
                    if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                    if ($Name -eq 'PnP.PowerShell') { return [pscustomobject]@{ Name = 'PnP.PowerShell' } }
                    return $null
                }
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Register-PnPManagementShellAccess -MockWith { throw 'access denied while registering' }
                Mock -CommandName Connect-PnPOnline -MockWith {
                    throw 'AADSTS65001: The user or administrator has not consented to use the application with ID abc'
                }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'AccessTokens'
                $Script:MSCloudLoginConnectionProfile.PnP.AccessTokens = @('raw-token')
                $Script:MSCloudLoginConnectionProfile.PnP.ConnectionUrl = 'https://contoso-admin.sharepoint.com'

                { Connect-MSCloudLoginPnP } |
                    Should -Throw "*has not been granted access for this tenant*Register-PnPManagementShellAccess*"
            }
        }
    }

    Context 'When the connection fails for any other reason' {
        It 'Should surface the underlying error and leave the workload disconnected' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-PnPOnline -MockWith { throw 'the site collection is unavailable' }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-Module -MockWith {
                    param ($Name)
                    if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                    if ($Name -eq 'PnP.PowerShell') { return [pscustomobject]@{ Name = 'PnP.PowerShell' } }
                    return $null
                }
                Mock -CommandName Import-Module -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $pnpProfile = $Script:MSCloudLoginConnectionProfile.PnP
                $pnpProfile.AuthenticationType = 'AccessTokens'
                $pnpProfile.AccessTokens = @('raw-token')
                $pnpProfile.ConnectionUrl = 'https://contoso-admin.sharepoint.com'
                $pnpProfile.PnPAzureEnvironment = 'Production'

                { Connect-MSCloudLoginPnP } | Should -Throw '*site collection is unavailable*'

                $pnpProfile.Connected | Should -BeFalse
                Should -Invoke Add-MSCloudLoginAssistantEvent -ParameterFilter {
                    $Message -like '*Failed to connect to PnP: *' -and $EntryType -eq 'Error'
                }
            }
        }
    }

    Context 'When only the admin URL is available after URL resolution' {
        It 'Should use the tenant GUID instead of the tenant name for AzureChinaCloud' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-Module -MockWith {
                    param ($Name)
                    if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                    if ($Name -eq 'PnP.PowerShell') { return [pscustomobject]@{ Name = 'PnP.PowerShell' } }
                    return $null
                }
                Mock -CommandName Import-Module -MockWith { }
                # An empty ConnectionUrl in the resolved result leaves the profile
                # without a connection URL, so the AdminUrl based path applies.
                Mock -CommandName Get-MSCloudLoginSPOUrlFromTenantId -MockWith {
                    return @{
                        AdminUrl      = 'https://contoso-admin.sharepoint.cn'
                        ConnectionUrl = ''
                    }
                }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'ServicePrincipalWithThumbprint'
                $Script:MSCloudLoginConnectionProfile.PnP.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.PnP.TenantId = 'contoso.partner.onmschina.cn'
                $Script:MSCloudLoginConnectionProfile.PnP.TenantGUID = '22222222-2222-2222-2222-222222222222'
                $Script:MSCloudLoginConnectionProfile.PnP.CertificateThumbprint = 'thumbprint'
                $Script:MSCloudLoginConnectionProfile.PnP.EnvironmentName = 'AzureChinaCloud'
                $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment = 'China'

                { Connect-MSCloudLoginPnP } | Should -Not -Throw

                Should -Invoke Connect-PnPOnline -Exactly 1 -ParameterFilter {
                    $Tenant -eq '22222222-2222-2222-2222-222222222222' -and
                    $Url -eq 'https://contoso-admin.sharepoint.cn' -and
                    $AzureEnvironment -eq 'China'
                }
            }
        }

        It 'Should pass the custom endpoints when connecting with a thumbprint through the admin URL' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-Module -MockWith {
                    param ($Name)
                    if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                    if ($Name -eq 'PnP.PowerShell') { return [pscustomobject]@{ Name = 'PnP.PowerShell' } }
                    return $null
                }
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Get-MSCloudLoginSPOUrlFromTenantId -MockWith {
                    return @{
                        AdminUrl      = 'https://contoso-admin.sharepoint.contoso.local'
                        ConnectionUrl = ''
                    }
                }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'ServicePrincipalWithThumbprint'
                $Script:MSCloudLoginConnectionProfile.PnP.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.PnP.TenantId = 'contoso.local'
                $Script:MSCloudLoginConnectionProfile.PnP.CertificateThumbprint = 'thumbprint'
                $Script:MSCloudLoginConnectionProfile.PnP.EndPoints = @{
                    AzureADLoginEndPoint   = 'https://login.contoso.local'
                    MicrosoftGraphEndPoint = 'https://graph.contoso.local'
                }
                $Script:MSCloudLoginConnectionProfile.PnP.EnvironmentName = 'AzureCloud'
                $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment = 'Custom'

                { Connect-MSCloudLoginPnP } | Should -Not -Throw

                Should -Invoke Connect-PnPOnline -Exactly 1 -ParameterFilter {
                    $Url -eq 'https://contoso-admin.sharepoint.contoso.local' -and
                    $AzureADLoginEndPoint -eq 'https://login.contoso.local' -and
                    $MicrosoftGraphEndPoint -eq 'https://graph.contoso.local'
                }
            }
        }

        It 'Should connect by certificate path through the admin URL' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-Module -MockWith {
                    param ($Name)
                    if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                    if ($Name -eq 'PnP.PowerShell') { return [pscustomobject]@{ Name = 'PnP.PowerShell' } }
                    return $null
                }
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Get-MSCloudLoginSPOUrlFromTenantId -MockWith {
                    return @{
                        AdminUrl      = 'https://contoso-admin.sharepoint.com'
                        ConnectionUrl = ''
                    }
                }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'ServicePrincipalWithPath'
                $Script:MSCloudLoginConnectionProfile.PnP.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.PnP.TenantId = 'contoso.onmicrosoft.com'
                $Script:MSCloudLoginConnectionProfile.PnP.CertificatePath = 'C:\certs\contoso.pfx'
                $Script:MSCloudLoginConnectionProfile.PnP.CertificatePassword =
                    ConvertTo-SecureString 'cert-password' -AsPlainText -Force
                $Script:MSCloudLoginConnectionProfile.PnP.EnvironmentName = 'AzureCloud'
                $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment = 'Production'

                { Connect-MSCloudLoginPnP } | Should -Not -Throw

                Should -Invoke Connect-PnPOnline -Exactly 1 -ParameterFilter {
                    $Url -eq 'https://contoso-admin.sharepoint.com' -and
                    $CertificatePath -eq 'C:\certs\contoso.pfx'
                }
            }
        }

        It 'Should connect with the client secret through the admin URL' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-Module -MockWith {
                    param ($Name)
                    if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                    if ($Name -eq 'PnP.PowerShell') { return [pscustomobject]@{ Name = 'PnP.PowerShell' } }
                    return $null
                }
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Get-MSCloudLoginSPOUrlFromTenantId -MockWith {
                    return @{
                        AdminUrl      = 'https://contoso-admin.sharepoint.com'
                        ConnectionUrl = ''
                    }
                }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'ServicePrincipalWithSecret'
                $Script:MSCloudLoginConnectionProfile.PnP.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.PnP.ApplicationSecret = 'secret'
                $Script:MSCloudLoginConnectionProfile.PnP.TenantId = 'contoso.onmicrosoft.com'
                $Script:MSCloudLoginConnectionProfile.PnP.EnvironmentName = 'AzureCloud'
                $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment = 'Production'

                { Connect-MSCloudLoginPnP } | Should -Not -Throw

                Should -Invoke Add-MSCloudLoginAssistantEvent -ParameterFilter {
                    $Message -like '*AdminUrl: https://contoso-admin.sharepoint.com*'
                }
                Should -Invoke Connect-PnPOnline -Exactly 1 -ParameterFilter {
                    $ClientSecret -eq 'secret' -and $Url -eq 'https://contoso-admin.sharepoint.com'
                }
            }
        }

        It 'Should connect with credentials and application id through the admin URL' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-Module -MockWith {
                    param ($Name)
                    if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                    if ($Name -eq 'PnP.PowerShell') { return [pscustomobject]@{ Name = 'PnP.PowerShell' } }
                    return $null
                }
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Get-MSCloudLoginSPOUrlFromTenantId -MockWith {
                    return @{
                        AdminUrl      = 'https://contoso-admin.sharepoint.com'
                        ConnectionUrl = ''
                    }
                }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'CredentialsWithApplicationId'
                $Script:MSCloudLoginConnectionProfile.PnP.ApplicationId = 'app-id'
                $Script:MSCloudLoginConnectionProfile.PnP.Credentials =
                    New-Object PSCredential ('admin@contoso.onmicrosoft.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))
                $Script:MSCloudLoginConnectionProfile.PnP.TenantId = 'contoso.onmicrosoft.com'
                $Script:MSCloudLoginConnectionProfile.PnP.EnvironmentName = 'AzureCloud'
                $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment = 'Production'

                { Connect-MSCloudLoginPnP } | Should -Not -Throw

                Should -Invoke Add-MSCloudLoginAssistantEvent -ParameterFilter {
                    $Message -like '*AdminUrl: https://contoso-admin.sharepoint.com*'
                }
                Should -Invoke Connect-PnPOnline -Exactly 1 -ParameterFilter {
                    $ClientId -eq 'app-id' -and $null -ne $Credentials -and
                    $Url -eq 'https://contoso-admin.sharepoint.com'
                }
            }
        }

        It 'Should connect with credentials through the admin URL when the type resolves late' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-Module -MockWith {
                    param ($Name)
                    if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                    if ($Name -eq 'PnP.PowerShell') { return [pscustomobject]@{ Name = 'PnP.PowerShell' } }
                    return $null
                }
                Mock -CommandName Import-Module -MockWith { }
                # The authentication type flips to Credentials after the URLs have been
                # resolved, so the connection URL stays empty and the admin URL based
                # credentials path applies.
                Mock -CommandName Get-MSCloudLoginSPOUrlFromTenantId -MockWith {
                    $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'Credentials'
                    return @{
                        AdminUrl      = 'https://contoso-admin.sharepoint.com'
                        ConnectionUrl = ''
                    }
                }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'AccessTokens'
                $Script:MSCloudLoginConnectionProfile.PnP.AccessTokens = @('raw-token')
                $Script:MSCloudLoginConnectionProfile.PnP.TenantId = 'contoso.onmicrosoft.com'
                $Script:MSCloudLoginConnectionProfile.PnP.EnvironmentName = 'AzureCloud'
                $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment = 'Production'

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'AccessTokens'
                $Script:MSCloudLoginConnectionProfile.PnP.AccessTokens = @('raw-token')
                $Script:MSCloudLoginConnectionProfile.PnP.TenantId = 'contoso.onmicrosoft.com'
                $Script:MSCloudLoginConnectionProfile.PnP.EnvironmentName = 'AzureCloud'
                $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment = 'Production'

                { Connect-MSCloudLoginPnP } | Should -Not -Throw

                $Script:MSCloudLoginConnectionProfile.PnP.Connected | Should -BeTrue
                Should -Invoke Add-MSCloudLoginAssistantEvent -ParameterFilter {
                    $Message -like '*using SPOManagementShell and AdminUrl*'
                }
                Should -Invoke Connect-PnPOnline -Exactly 1 -ParameterFilter {
                    $ClientId -eq '9bc3ab49-b65d-410a-85ad-de819febfddc' -and
                    $Url -eq 'https://contoso-admin.sharepoint.com'
                }
            }
        }

        It 'Should request the managed identity token for the admin URL' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-Module -MockWith {
                    param ($Name)
                    if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                    if ($Name -eq 'PnP.PowerShell') { return [pscustomobject]@{ Name = 'PnP.PowerShell' } }
                    return $null
                }
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Get-MSCloudLoginSPOUrlFromTenantId -MockWith {
                    return @{
                        AdminUrl      = 'https://contoso-admin.sharepoint.com'
                        ConnectionUrl = ''
                    }
                }
                Mock -CommandName Get-AuthToken -MockWith { return 'managed-identity-token' }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'Identity'
                $Script:MSCloudLoginConnectionProfile.PnP.TenantId = 'contoso.onmicrosoft.com'
                $Script:MSCloudLoginConnectionProfile.PnP.EnvironmentName = 'AzureCloud'
                $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment = 'Production'

                { Connect-MSCloudLoginPnP } | Should -Not -Throw

                Should -Invoke Get-AuthToken -Exactly 1 -ParameterFilter {
                    $Resource -eq 'https://contoso-admin.sharepoint.com' -and $Identity.IsPresent
                }
                Should -Invoke Connect-PnPOnline -Exactly 1 -ParameterFilter {
                    $AccessToken -eq 'managed-identity-token'
                }
            }
        }

        It 'Should connect with an access token through the admin URL' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Get-Module -MockWith {
                    param ($Name)
                    if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                    if ($Name -eq 'PnP.PowerShell') { return [pscustomobject]@{ Name = 'PnP.PowerShell' } }
                    return $null
                }
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Get-MSCloudLoginSPOUrlFromTenantId -MockWith {
                    return @{
                        AdminUrl      = 'https://contoso-admin.sharepoint.com'
                        ConnectionUrl = ''
                    }
                }
                Mock -CommandName Get-MSCloudLoginAccessTokenValue -MockWith { return 'resolved-token' }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'AccessTokens'
                $Script:MSCloudLoginConnectionProfile.PnP.AccessTokens = @('raw-token')
                $Script:MSCloudLoginConnectionProfile.PnP.TenantId = 'contoso.onmicrosoft.com'
                $Script:MSCloudLoginConnectionProfile.PnP.EnvironmentName = 'AzureCloud'
                $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment = 'Production'

                { Connect-MSCloudLoginPnP } | Should -Not -Throw

                Should -Invoke Connect-PnPOnline -Exactly 1 -ParameterFilter {
                    $AccessToken -eq 'resolved-token' -and
                    $Url -eq 'https://contoso-admin.sharepoint.com'
                }
            }
        }

        It 'Should retry interactively against the admin URL on the password mismatch error' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:pnpConnectCalls = 0
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
                Mock -CommandName Test-MSCloudLoginConnectionReusable -MockWith { return $false }
                Mock -CommandName Assert-IsNonInteractiveShell -MockWith { return $false }
                Mock -CommandName Get-Module -MockWith {
                    param ($Name)
                    if ($Name -eq 'Microsoft.Graph.Authentication') { return $null }
                    if ($Name -eq 'PnP.PowerShell') { return [pscustomobject]@{ Name = 'PnP.PowerShell' } }
                    return $null
                }
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Get-MSCloudLoginSPOUrlFromTenantId -MockWith {
                    return @{
                        AdminUrl      = 'https://contoso-admin.sharepoint.com'
                        ConnectionUrl = ''
                    }
                }
                Mock -CommandName Get-MSCloudLoginAccessTokenValue -MockWith { return 'resolved-token' }
                Mock -CommandName Connect-PnPOnline -MockWith {
                    $Script:pnpConnectCalls++
                    if ($Script:pnpConnectCalls -eq 1)
                    {
                        throw 'The sign-in name or password does not match one in the Microsoft account system.'
                    }
                }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.AuthenticationType = 'AccessTokens'
                $Script:MSCloudLoginConnectionProfile.PnP.AccessTokens = @('raw-token')
                $Script:MSCloudLoginConnectionProfile.PnP.TenantId = 'contoso.onmicrosoft.com'
                $Script:MSCloudLoginConnectionProfile.PnP.EnvironmentName = 'AzureCloud'
                $Script:MSCloudLoginConnectionProfile.PnP.PnPAzureEnvironment = 'Production'

                { Connect-MSCloudLoginPnP } | Should -Not -Throw

                $Script:MSCloudLoginConnectionProfile.PnP.Connected | Should -BeTrue
                $Script:MSCloudLoginConnectionProfile.PnP.MultiFactorAuthentication | Should -BeTrue
                $Script:pnpConnectCalls | Should -Be 2
                Should -Invoke Connect-PnPOnline -Exactly 1 -ParameterFilter {
                    $Interactive.IsPresent -and $Url -eq 'https://contoso-admin.sharepoint.com'
                }
            }
        }
    }
}

Describe 'Disconnect-MSCloudLoginPnP' {
    Context 'When PnP is connected' {
        It 'Should call Disconnect-PnPOnline' {
            InModuleScope 'MSCloudLoginAssistant' {
                Import-Module ./Tests/Unit/Stubs/Stubs.psm1 -Force
                Mock -CommandName Disconnect-PnPOnline -MockWith { }
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.Connected = $true

                Disconnect-MSCloudLoginPnP

                Should -Invoke Disconnect-PnPOnline
            }
        }
    }

    Context 'When PnP is not connected' {
        It 'Should not throw and log message' {
            InModuleScope 'MSCloudLoginAssistant' {
                Import-Module ./Tests/Unit/Stubs/Stubs.psm1 -Force
                Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }

                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $Script:MSCloudLoginConnectionProfile.PnP.Connected = $false

                { Disconnect-MSCloudLoginPnP } | Should -Not -Throw
            }
        }
    }
}

AfterAll {
    Remove-Module MSCloudLoginAssistant
}
