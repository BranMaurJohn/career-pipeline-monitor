[CmdletBinding()]
param(
    [string]$CurrentPath = (Join-Path $PSScriptRoot '..\data\current\jobs.json'),
    [string]$PreviousPath = (Join-Path $PSScriptRoot '..\data\previous\jobs.json'),
    [string]$OutputPath = (Join-Path $PSScriptRoot '..\data\changes\latest.json')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot '..\modules\Utilities.psm1') -Force

function Get-JobKey {
    param($Job)
    $employer = ConvertTo-NormalizedText ([string]$Job.employer)
    $title = ConvertTo-NormalizedText ([string]$Job.roleTitle)
    $url = ([string]$Job.jobUrl).Trim().ToLowerInvariant()
    return Get-Sha256Text -Text ($employer + '|' + $title + '|' + $url)
}

$currentPayload = Get-JsonFile -Path $CurrentPath -Default ([pscustomobject]@{ jobs = @() })
$previousPayload = Get-JsonFile -Path $PreviousPath -Default ([pscustomobject]@{ jobs = @() })
$currentJobs = @($currentPayload.jobs)
$previousJobs = @($previousPayload.jobs)

$currentMap = @{}
foreach ($job in $currentJobs) { $currentMap[(Get-JobKey -Job $job)] = $job }
$previousMap = @{}
foreach ($job in $previousJobs) { $previousMap[(Get-JobKey -Job $job)] = $job }

$newJobs = @()
foreach ($key in $currentMap.Keys) {
    if (-not $previousMap.ContainsKey($key)) { $newJobs += $currentMap[$key] }
}

$removedJobs = @()
foreach ($key in $previousMap.Keys) {
    if (-not $currentMap.ContainsKey($key)) { $removedJobs += $previousMap[$key] }
}

$changes = [pscustomobject][ordered]@{
    generatedAt     = (Get-Date).ToUniversalTime().ToString('o')
    currentJobCount = $currentJobs.Count
    previousJobCount = $previousJobs.Count
    newJobCount     = $newJobs.Count
    removedJobCount = $removedJobs.Count
    newJobs         = @($newJobs)
    removedJobs     = @($removedJobs)
    removalPolicy   = 'Removed postings are audit-only. ClickUp tasks are never automatically deleted or demoted.'
}

Set-JsonFile -Path $OutputPath -Value $changes -Depth 12
$changes
