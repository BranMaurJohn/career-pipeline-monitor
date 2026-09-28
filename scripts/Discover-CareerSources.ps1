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
Import-Module (Join-Path $PSScriptRoot '..\modules\Utilities.psm1') -Force
Import-Module (Join-Path $PSScriptRoot '..\modules\CareerSources.psm1') -Force

if (-not $OutputPath) { $OutputPath = $EmployersPath }
$registry = Get-JsonFile -Path $EmployersPath -Default $null
$settings = Get-JsonFile -Path $SettingsPath -Default $null
if (-not $registry -or -not $registry.employers) { throw "Employer registry is empty: $EmployersPath" }
if (-not $settings) { throw "Settings file is empty: $SettingsPath" }
if ($ShardCount -lt 1) { throw 'ShardCount must be at least 1.' }
if ($ShardIndex -lt 0 -or $ShardIndex -ge $ShardCount) { throw 'ShardIndex must be between 0 and ShardCount - 1.' }

$employers = @($registry.employers)
for ($index = 0; $index -lt $employers.Count; $index++) {
    if (($index % $ShardCount) -ne $ShardIndex) { continue }
    $employer = $employers[$index]
    if (-not [bool]$employer.active) { continue }
    if (-not $RefreshExisting -and (Test-HttpUrl ([string]$employer.careersUrl))) { continue }

    Write-PipelineLog ("Discovering career source [{0}/{1}] {2}" -f ($index + 1), $employers.Count, [string]$employer.name)
    $source = Find-CareerSource -Employer $employer -Settings $settings
    if ($source) {
        $employer.careersUrl = [string]$source.careersUrl
        $employer.ats = [string]$source.ats
        $employer.lastDiscoveredAt = (Get-Date).ToUniversalTime().ToString('o')
        Write-PipelineLog ("Career source: {0}" -f [string]$source.careersUrl) -Level Success
    } else {
        Write-PipelineLog 'No official career source found in this pass.' -Level Warn
    }
    Start-Sleep -Milliseconds ([int]$settings.request.delayMilliseconds)
}

$registry.generatedAt = (Get-Date).ToUniversalTime().ToString('o')
Set-JsonFile -Path $OutputPath -Value $registry -Depth 12
$registry
