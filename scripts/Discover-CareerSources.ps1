[CmdletBinding()]
param(
    [string]$EmployersPath = (Join-Path $PSScriptRoot '..\config\employers.json'),
    [string]$SettingsPath = (Join-Path $PSScriptRoot '..\config\settings.json'),
    [string]$OutputPath = '',
    [int]$ShardIndex = 0,
    [int]$ShardCount = 1,
    [switch]$RefreshExisting
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$utilitiesModule = Join-Path $PSScriptRoot '..\modules\Utilities.psm1'
$jobMatchingModule = Join-Path $PSScriptRoot '..\modules\JobMatching.psm1'
$careerSourcesModule = Join-Path $PSScriptRoot '..\modules\CareerSources.psm1'

if (-not (Test-Path -LiteralPath $utilitiesModule)) {
    throw "Required module not found: $utilitiesModule"
}

if (-not (Test-Path -LiteralPath $jobMatchingModule)) {
    throw "Required module not found: $jobMatchingModule"
}

if (-not (Test-Path -LiteralPath $careerSourcesModule)) {
    throw "Required module not found: $careerSourcesModule"
}

#
# Import dependency modules first.
#
# CareerSources imports Utilities and JobMatching internally. Importing
# Utilities LAST into the caller scope guarantees that commands such as
# Get-JsonFile, Set-JsonFile, Test-HttpUrl, and Write-PipelineLog remain
# directly available to this script after CareerSources has loaded.
#
Import-Module $jobMatchingModule -Force -Global -ErrorAction Stop
Import-Module $careerSourcesModule -Force -Global -ErrorAction Stop
Import-Module $utilitiesModule -Force -Global -ErrorAction Stop

#
# Fail immediately if required exported commands are not visible.
#
$requiredCommands = @(
    'Get-JsonFile',
    'Set-JsonFile',
    'Test-HttpUrl',
    'Write-PipelineLog',
    'Find-CareerSource'
)

foreach ($commandName in $requiredCommands) {
    if (-not (Get-Command $commandName -ErrorAction SilentlyContinue)) {
        throw "Required command '$commandName' is not available after module import."
    }
}

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = $EmployersPath
}

$registry = Get-JsonFile `
    -Path $EmployersPath `
    -Default $null

$settings = Get-JsonFile `
    -Path $SettingsPath `
    -Default $null

if (-not $registry) {
    throw "Employer registry is empty: $EmployersPath"
}

if (
    $null -eq $registry.PSObject.Properties['employers'] -or
    @($registry.employers).Count -eq 0
) {
    throw "Employer registry contains no employers: $EmployersPath"
}

if (-not $settings) {
    throw "Settings file is empty: $SettingsPath"
}

if ($ShardCount -lt 1) {
    throw 'ShardCount must be at least 1.'
}

if ($ShardIndex -lt 0 -or $ShardIndex -ge $ShardCount) {
    throw 'ShardIndex must be between 0 and ShardCount - 1.'
}

$employers = @($registry.employers)

Write-PipelineLog (
    "Career source discovery starting for shard {0} of {1}. Total employers: {2}" -f `
        ($ShardIndex + 1),
        $ShardCount,
        $employers.Count
)

for ($index = 0; $index -lt $employers.Count; $index++) {

    if (($index % $ShardCount) -ne $ShardIndex) {
        continue
    }

    $employer = $employers[$index]

    $isActive = $true

    if ($null -ne $employer.PSObject.Properties['active']) {
        $isActive = [bool]$employer.active
    }

    if (-not $isActive) {
        continue
    }

    $existingCareersUrl = ''

    if ($null -ne $employer.PSObject.Properties['careersUrl']) {
        $existingCareersUrl = [string]$employer.careersUrl
    }

    if (
        -not $RefreshExisting -and
        (Test-HttpUrl $existingCareersUrl)
    ) {
        continue
    }

    Write-PipelineLog (
        "Discovering career source [{0}/{1}] {2}" -f `
            ($index + 1),
            $employers.Count,
            [string]$employer.name
    )

    try {
        $source = Find-CareerSource `
            -Employer $employer `
            -Settings $settings

        if ($source) {

            if ($null -eq $employer.PSObject.Properties['careersUrl']) {
                $employer |
                    Add-Member `
                        -MemberType NoteProperty `
                        -Name careersUrl `
                        -Value '' `
                        -Force
            }

            if ($null -eq $employer.PSObject.Properties['ats']) {
                $employer |
                    Add-Member `
                        -MemberType NoteProperty `
                        -Name ats `
                        -Value '' `
                        -Force
            }

            if ($null -eq $employer.PSObject.Properties['lastDiscoveredAt']) {
                $employer |
                    Add-Member `
                        -MemberType NoteProperty `
                        -Name lastDiscoveredAt `
                        -Value '' `
                        -Force
            }

            $employer.careersUrl = [string]$source.careersUrl
            $employer.ats = [string]$source.ats
            $employer.lastDiscoveredAt = (
                Get-Date
            ).ToUniversalTime().ToString('o')

            Write-PipelineLog (
                "Career source: {0}" -f [string]$source.careersUrl
            ) -Level Success
        }
        else {
            Write-PipelineLog (
                "No official career source found for {0} in this pass." -f `
                    [string]$employer.name
            ) -Level Warn
        }
    }
    catch {
        Write-PipelineLog (
            "Career source discovery failed for {0}: {1}" -f `
                [string]$employer.name,
                $_.Exception.Message
        ) -Level Warn
    }

    $delayMilliseconds = 0

    if (
        $null -ne $settings.PSObject.Properties['request'] -and
        $null -ne $settings.request -and
        $null -ne $settings.request.PSObject.Properties['delayMilliseconds']
    ) {
        $delayMilliseconds = [int]$settings.request.delayMilliseconds
    }

    if ($delayMilliseconds -gt 0) {
        Start-Sleep -Milliseconds $delayMilliseconds
    }
}

if ($null -eq $registry.PSObject.Properties['generatedAt']) {
    $registry |
        Add-Member `
            -MemberType NoteProperty `
            -Name generatedAt `
            -Value '' `
            -Force
}

$registry.generatedAt = (
    Get-Date
).ToUniversalTime().ToString('o')

Set-JsonFile `
    -Path $OutputPath `
    -Value $registry `
    -Depth 12

Write-PipelineLog (
    "Career source discovery completed for shard {0} of {1}." -f `
        ($ShardIndex + 1),
        $ShardCount
) -Level Success

$registry
