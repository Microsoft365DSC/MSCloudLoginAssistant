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

Describe 'Connect-M365Tenant URL routing' {

    BeforeEach {
        InModuleScope 'MSCloudLoginAssistant' {
            $Script:MSCloudLoginConnectionProfile = $null
            $Script:MSCloudLoginTriedGetEnvironment = $true
            $Script:CloudEnvironmentInfo = $null
        }
    }

    Context 'When PnP is asked for the URL it is already configured for' {
        It 'Should connect <ExpectedConnections> time(s) when the live PnP context <Description>' -TestCases @(
            @{ Description = 'points somewhere else'; ContextUrl = 'https://contoso.sharepoint.com/sites/marketing'; ExpectedConnections = 2 }
            @{ Description = 'matches'; ContextUrl = 'https://contoso-admin.sharepoint.com'; ExpectedConnections = 1 }
            @{ Description = 'cannot be read'; ContextUrl = $null; ExpectedConnections = 1 }
        ) {
            param ($ContextUrl, $ExpectedConnections)
            InModuleScope 'MSCloudLoginAssistant' -Parameters @{ ContextUrl = $ContextUrl; ExpectedConnections = $ExpectedConnections } {
                param ($ContextUrl, $ExpectedConnections)

                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Get-PnPContext -MockWith {
                    if ($null -eq $ContextUrl)
                    {
                        throw 'no context available'
                    }
                    return @{ Url = $ContextUrl }
                }.GetNewClosure()

                $parameters = @{
                    Workload              = 'PnP'
                    Url                   = 'https://contoso-admin.sharepoint.com'
                    ApplicationId         = '11111111-1111-1111-1111-111111111111'
                    TenantId              = 'contoso.onmicrosoft.com'
                    CertificateThumbprint = 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD'
                }
                Connect-M365Tenant @parameters
                Connect-M365Tenant @parameters

                Should -Invoke Connect-PnPOnline -Exactly $ExpectedConnections
                (Get-MSCloudLoginConnectionProfile -Workload 'PnP').ConnectionUrl | Should -Be 'https://contoso-admin.sharepoint.com'
            }
        }
    }

    Context 'When PnP is asked for a different URL' {
        It 'Should force a refresh of the connection' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Import-Module -MockWith { }
                Mock -CommandName Connect-PnPOnline -MockWith { }
                Mock -CommandName Get-PnPContext -MockWith { return @{ Url = 'https://contoso-admin.sharepoint.com' } }

                Connect-M365Tenant -Workload 'PnP' -Url 'https://contoso-admin.sharepoint.com' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -CertificateThumbprint 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD'
                Connect-M365Tenant -Workload 'PnP' -Url 'https://contoso.sharepoint.com/sites/marketing' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -CertificateThumbprint 'AA11BB22CC33DD44EE55FF6677889900AABBCCDD'

                $connection = Get-MSCloudLoginConnectionProfile -Workload 'PnP'
                $connection.ConnectionUrl | Should -Be 'https://contoso.sharepoint.com/sites/marketing'
                $connection.AdminUrl | Should -Be 'https://contoso-admin.sharepoint.com'
                Should -Invoke Connect-PnPOnline -Exactly 2
            }
        }
    }

    Context 'When SharePoint Online REST is given an explicit URL' {
        It 'Should adopt that URL as the PnP admin URL' {
            InModuleScope 'MSCloudLoginAssistant' {
                Mock -CommandName Invoke-WebRequest -MockWith {
                    return @{ Content = '{ "token_endpoint": "https://login.microsoftonline.com/t/oauth2/v2.0/token" }' }
                }
                Mock -CommandName Invoke-RestMethod -MockWith {
                    return @{ token_type = 'Bearer'; access_token = 'token' }
                }

                Connect-M365Tenant -Workload 'SharePointOnlineREST' `
                    -Url 'https://contoso-admin.sharepoint.com' `
                    -ApplicationId '11111111-1111-1111-1111-111111111111' `
                    -TenantId 'contoso.onmicrosoft.com' `
                    -ApplicationSecret 'super-secret'

                (Get-MSCloudLoginConnectionProfile -Workload 'PnP').AdminUrl | Should -Be 'https://contoso-admin.sharepoint.com'
            }
        }
    }
}

Describe 'Reset-MSCloudLoginConnectionProfileContext failure handling' {

    It 'Should keep resetting the remaining workloads when one disconnect fails' {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
            Mock -CommandName Disconnect-MSCloudLoginFabric -MockWith { throw 'the session is already gone' }
            Mock -CommandName Disconnect-MSCloudLoginTasks -MockWith { }

            $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile

            { Reset-MSCloudLoginConnectionProfileContext -Workload 'Fabric', 'Tasks' } | Should -Not -Throw

            Should -Invoke Disconnect-MSCloudLoginTasks -Exactly 1
            Should -Invoke Add-MSCloudLoginAssistantEvent -ParameterFilter {
                $Message -like 'Failed to disconnect workload {Fabric}*' -and $EntryType -eq 'Error'
            }
        }
    }
}

Describe 'Compare-InputParametersForChange session parameters' {

    BeforeAll {
        InModuleScope 'MSCloudLoginAssistant' {
            Mock -CommandName Add-MSCloudLoginAssistantEvent -MockWith { }
        }
    }

    Context 'Exchange Online cmdlets' {
        It 'Should report a change only when another set of cmdlets is requested' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $workloadProfile = $Script:MSCloudLoginConnectionProfile.ExchangeOnline
                $workloadProfile.AuthenticationType = 'ServicePrincipalWithSecret'
                $workloadProfile.RequestedAuthenticationType = 'ServicePrincipalWithSecret'
                $workloadProfile.ApplicationId = 'app-id'
                $workloadProfile.TenantId = 'contoso.onmicrosoft.com'
                $workloadProfile.ApplicationSecret = 'secret'
                $workloadProfile.CmdletsToLoad = @('Get-Mailbox', 'Set-Mailbox')

                $parameters = @{
                    Workload              = 'ExchangeOnline'
                    ApplicationId         = 'app-id'
                    TenantId              = 'contoso.onmicrosoft.com'
                    ApplicationSecret     = 'secret'
                    ExchangeOnlineCmdlets = @('Set-Mailbox', 'Get-Mailbox')
                }
                (Compare-InputParametersForChange -CurrentParamSet $parameters) | Should -BeFalse

                $parameters['ExchangeOnlineCmdlets'] = @('Get-Mailbox', 'Get-User')
                (Compare-InputParametersForChange -CurrentParamSet $parameters) | Should -BeTrue
            }
        }
    }

    Context 'Connection URLs' {
        It 'Should compare the URL of a PnP connection against the active connection URL' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $workloadProfile = $Script:MSCloudLoginConnectionProfile.PnP
                $workloadProfile.AuthenticationType = 'ServicePrincipalWithThumbprint'
                $workloadProfile.RequestedAuthenticationType = 'ServicePrincipalWithThumbprint'
                $workloadProfile.ApplicationId = 'app-id'
                $workloadProfile.TenantId = 'contoso.onmicrosoft.com'
                $workloadProfile.CertificateThumbprint = 'thumbprint'
                $workloadProfile.ConnectionUrl = 'https://contoso-admin.sharepoint.com'

                $unchanged = @{
                    Workload              = 'PnP'
                    ApplicationId         = 'app-id'
                    TenantId              = 'contoso.onmicrosoft.com'
                    CertificateThumbprint = 'thumbprint'
                    Url                   = 'https://contoso-admin.sharepoint.com'
                }
                (Compare-InputParametersForChange -CurrentParamSet $unchanged) | Should -BeFalse

                $changed = $unchanged.Clone()
                $changed['Url'] = 'https://contoso.sharepoint.com/sites/marketing'
                (Compare-InputParametersForChange -CurrentParamSet $changed) | Should -BeTrue
            }
        }

        It 'Should compare only the session keys when no identity parameter is repeated' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $workloadProfile = $Script:MSCloudLoginConnectionProfile.SharePointOnlineREST
                $workloadProfile.AuthenticationType = 'ServicePrincipalWithSecret'
                $workloadProfile.RequestedAuthenticationType = 'ServicePrincipalWithSecret'
                $workloadProfile.ApplicationId = 'app-id'
                $workloadProfile.TenantId = 'contoso.onmicrosoft.com'
                $workloadProfile.ApplicationSecret = 'secret'
                $workloadProfile.ConnectionUrl = 'https://contoso-admin.sharepoint.com'

                $parameters = @{
                    Workload = 'SharePointOnlineREST'
                    Url      = 'https://contoso-admin.sharepoint.com'
                }
                (Compare-InputParametersForChange -CurrentParamSet $parameters -LimitToKeys @('ConnectionUrl')) | Should -BeFalse

                $parameters['Url'] = 'https://contoso.sharepoint.com'
                (Compare-InputParametersForChange -CurrentParamSet $parameters -LimitToKeys @('ConnectionUrl')) | Should -BeTrue
            }
        }
    }

    Context 'Search only sessions' {
        It 'Should report a change only when the search only session was turned on' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $workloadProfile = $Script:MSCloudLoginConnectionProfile.SecurityComplianceCenter
                $workloadProfile.AuthenticationType = 'ServicePrincipalWithThumbprint'
                $workloadProfile.RequestedAuthenticationType = 'ServicePrincipalWithThumbprint'
                $workloadProfile.ApplicationId = 'app-id'
                $workloadProfile.TenantId = 'contoso.onmicrosoft.com'
                $workloadProfile.CertificateThumbprint = 'thumbprint'
                $workloadProfile.EnableSearchOnlySession = $false

                $parameters = @{
                    Workload                = 'SecurityComplianceCenter'
                    ApplicationId           = 'app-id'
                    TenantId                = 'contoso.onmicrosoft.com'
                    CertificateThumbprint   = 'thumbprint'
                    EnableSearchOnlySession = $true
                }
                (Compare-InputParametersForChange -CurrentParamSet $parameters) | Should -BeTrue

                $workloadProfile.EnableSearchOnlySession = $true
                (Compare-InputParametersForChange -CurrentParamSet $parameters) | Should -BeFalse
            }
        }
    }

    Context 'Microsoft Graph tenant inference' {
        It 'Should ignore only the tenant that was inferred from the credential UPN' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                $credential = New-Object PSCredential ('admin@contoso.com', (ConvertTo-SecureString 'p@ssw0rd' -AsPlainText -Force))
                $workloadProfile = $Script:MSCloudLoginConnectionProfile.MicrosoftGraph
                $workloadProfile.AuthenticationType = 'Credentials'
                $workloadProfile.RequestedAuthenticationType = 'Credentials'
                $workloadProfile.Credentials = $credential
                $workloadProfile.TenantId = 'contoso.com'

                (Compare-InputParametersForChange -CurrentParamSet @{ Workload = 'MicrosoftGraph'; Credential = $credential }) | Should -BeFalse

                $workloadProfile.TenantId = 'fabrikam.com'
                (Compare-InputParametersForChange -CurrentParamSet @{ Workload = 'MicrosoftGraph'; Credential = $credential }) | Should -BeTrue
            }
        }
    }

    Context 'Unknown workloads' {
        It 'Should report a change when the workload is unknown or no parameter set is supplied at all' {
            InModuleScope 'MSCloudLoginAssistant' {
                $Script:MSCloudLoginConnectionProfile = New-Object MSCloudLoginConnectionProfile
                (Compare-InputParametersForChange -CurrentParamSet @{ Workload = 'DoesNotExist' }) | Should -BeTrue
                (Compare-InputParametersForChange) | Should -BeTrue
            }
        }
    }
}
