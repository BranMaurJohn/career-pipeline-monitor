[CmdletBinding()]
param(
    [string]$EmployersPath = (Join-Path $PSScriptRoot '..\config\employers.json'),
    [string]$JobsPath = (Join-Path $PSScriptRoot '..\data\current\jobs.json'),
    [string]$SettingsPath = (Join-Path $PSScriptRoot '..\config\settings.json'),
    [switch]$DryRun
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot '..\modules\Utilities.psm1') -Force
Import-Module (Join-Path $PSScriptRoot '..\modules\ClickUp.psm1') -Force

$settings = Get-JsonFile -Path $SettingsPath -Default $null
if (-not $settings) { throw "Settings file is empty: $SettingsPath" }

$token = [string]$env:CLICKUP_API_TOKEN
$listId = [string]$env:CLICKUP_LIST_ID
if ([string]::IsNullOrWhiteSpace($listId) -and $settings.clickup.listId) { $listId = [string]$settings.clickup.listId }
if ([string]::IsNullOrWhiteSpace($token)) { throw 'CLICKUP_API_TOKEN is not set.' }
if ([string]::IsNullOrWhiteSpace($listId)) { throw 'CLICKUP_LIST_ID is not set and config/settings.json does not contain clickup.listId.' }

$registry = Get-JsonFile -Path $EmployersPath -Default $null
$jobsPayload = Get-JsonFile -Path $JobsPath -Default ([pscustomobject]@{ jobs = @() })
if (-not $registry -or -not $registry.employers) { throw "Employer registry is empty: $EmployersPath" }
$actions = @(Sync-CareerPipelineToClickUp `
    -Employers @($registry.employers) `
    -Jobs @($jobsPayload.jobs) `
    -ListId $listId `
    -Token $token `
    -TargetCompanyStatus ([string]$settings.clickup.targetCompanyStatus) `
    -RoleIdentifiedStatus ([string]$settings.clickup.roleIdentifiedStatus) `
    -WriteDelayMilliseconds ([int]$settings.clickup.writeDelayMilliseconds) `
    -DryRun:$DryRun)

$actions | Format-Table action, taskName, status -AutoSize | Out-String | Write-Host
Write-PipelineLog ("ClickUp sync evaluated {0} action(s). DryRun={1}" -f $actions.Count, [bool]$DryRun) -Level Success
$actions
