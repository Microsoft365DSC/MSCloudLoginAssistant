function New-MSCloudLoginAssistantUnitTestHelper
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [String]
        $StubModule,

        [Parameter(Mandatory = $true, ParameterSetName = 'SubModule')]
        [String]
        $SubModulePath,

        [Parameter()]
        [Switch]
        $ExcludeInvokeHelper,

        [Parameter()]
        [Switch]
        $IncludeDistributedCacheStubs
    )

    $repoRoot = Join-Path -Path $PSScriptRoot -ChildPath "..\..\" -Resolve
    $moduleRoot = Join-Path -Path $repoRoot -ChildPath "Modules\MSCloudLoginAssistant"

    $mainModule = Join-Path -Path $moduleRoot -ChildPath "MSCloudLoginAssistant.psd1"
    Import-Module -Name $mainModule -Global

    if ($PSBoundParameters.ContainsKey("SubModulePath") -eq $true)
    {
        $describeHeader = "Sub-module '$SubModulePath'"
        $moduleToLoad = Join-Path -Path $moduleRoot -ChildPath $SubModulePath
        $moduleName = (Get-Item -Path $moduleToLoad).BaseName
    }

    Import-Module -Name $moduleToLoad -Global

    $initScript = @"
            Import-Module -Name "$StubModule" -WarningAction SilentlyContinue
            Import-Module -Name "$moduleToLoad"

"@

    return @{
        DescribeHeader        = $describeHeader
        ModuleName            = $moduleName
        CurrentStubModulePath = $StubModule
        InitializeScript      = [ScriptBlock]::Create($initScript)
        RepoRoot              = $repoRoot
        CleanupScript         = [ScriptBlock]::Create(@"

            `$global:DSCMachineStatus = 0

"@)
    }
}
