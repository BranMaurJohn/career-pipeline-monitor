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

foreach ($modulePath in @($utilitiesModule, $jobMatchingModule, $careerSourcesModule)) {
    if (-not (Test-Path -LiteralPath $modulePath)) {
        throw "Required module not found: $modulePath"
    }
}

Import-Module $jobMatchingModule -Force -ErrorAction Stop
Import-Module $careerSourcesModule -Force -ErrorAction Stop
Import-Module $utilitiesModule -Force -ErrorAction Stop

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = $EmployersPath
}

$registry = Utilities\Get-JsonFile -Path $EmployersPath -Default $null
$settings = Utilities\Get-JsonFile -Path $SettingsPath -Default $null

if (-not $registry) { throw "Employer registry is empty: $EmployersPath" }
if ($null -eq $registry.PSObject.Properties['employers'] -or @($registry.employers).Count -eq 0) {
    throw "Employer registry contains no employers: $EmployersPath"
}
if (-not $settings) { throw "Settings file is empty: $SettingsPath" }
if ($ShardCount -lt 1) { throw 'ShardCount must be at least 1.' }
if ($ShardIndex -lt 0 -or $ShardIndex -ge $ShardCount) {
    throw 'ShardIndex must be between 0 and ShardCount - 1.'
}

$employers = @($registry.employers)

Utilities\Write-PipelineLog (
    "Career source discovery starting for shard {0} of {1}. Total employers: {2}" -f `
        ($ShardIndex + 1), $ShardCount, $employers.Count
)

for ($index = 0; $index -lt $employers.Count; $index++) {
    if (($index % $ShardCount) -ne $ShardIndex) { continue }

    $employer = $employers[$index]
    $isActive = $true
    if ($null -ne $employer.PSObject.Properties['active']) {
        $isActive = [bool]$employer.active
    }
    if (-not $isActive) { continue }

    foreach ($propertyName in @('careersUrl','ats','lastDiscoveredAt')) {
        if ($null -eq $employer.PSObject.Properties[$propertyName]) {
            $employer | Add-Member -MemberType NoteProperty -Name $propertyName -Value '' -Force
        }
    }

    $existingCareersUrl = [string]$employer.careersUrl
    if (-not $RefreshExisting -and (Utilities\Test-HttpUrl $existingCareersUrl)) {
        continue
    }

    Utilities\Write-PipelineLog (
        "Discovering career source [{0}/{1}] {2}" -f `
            ($index + 1), $employers.Count, [string]$employer.name
    )

    try {
        $source = CareerSources\Find-CareerSource `
            -Employer $employer `
            -Settings $settings `
            -IgnoreRegistry:$RefreshExisting

        if ($source) {
            $employer.careersUrl = [string]$source.careersUrl
            $employer.ats = [string]$source.ats
            $employer.lastDiscoveredAt = (Get-Date).ToUniversalTime().ToString('o')

            Utilities\Write-PipelineLog (
                "Career source: {0}" -f [string]$source.careersUrl
            ) -Level Success
        }
        else {
            # Refresh mode is authoritative. If a previously cached source can
            # no longer pass current official-source validation, remove it from
            # the registry so the hourly monitor cannot consume stale or false
            # source data. ClickUp tasks are never deleted by this action.
            if ($RefreshExisting) {
                $employer.careersUrl = ''
                $employer.ats = ''
                $employer.lastDiscoveredAt = ''
            }

            Utilities\Write-PipelineLog (
                "No verified official career source found for {0} in this pass." -f [string]$employer.name
            ) -Level Warn
        }
    }
    catch {
        Utilities\Write-PipelineLog (
            "Career source discovery failed for {0}: {1}" -f `
                [string]$employer.name, $_.Exception.Message
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
    $registry | Add-Member -MemberType NoteProperty -Name generatedAt -Value '' -Force
}
$registry.generatedAt = (Get-Date).ToUniversalTime().ToString('o')

Utilities\Set-JsonFile -Path $OutputPath -Value $registry -Depth 12

Utilities\Write-PipelineLog (
    "Career source discovery completed for shard {0} of {1}." -f `
        ($ShardIndex + 1), $ShardCount
) -Level Success

$registry
