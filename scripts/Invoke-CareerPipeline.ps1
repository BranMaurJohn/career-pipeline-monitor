[CmdletBinding()]
param(
    [ValidateSet('Hourly','HourlyAggregate','DeepDiscovery','Full')]
    [string]$Mode = 'Hourly',
    [switch]$ApplyChanges,
    [int]$ShardIndex = 0,
    [int]$ShardCount = 1,
    [string]$ArtifactDirectory = ''
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot '..\modules\Utilities.psm1') -Force
Import-Module (Join-Path $PSScriptRoot '..\modules\JobMatching.psm1') -Force

$root = Resolve-Path (Join-Path $PSScriptRoot '..')
$employersPath = Join-Path $root 'config\employers.json'
$settingsPath = Join-Path $root 'config\settings.json'
$keywordsPath = Join-Path $root 'config\keywords.json'
$currentJobsPath = Join-Path $root 'data\current\jobs.json'
$currentStatePath = Join-Path $root 'data\current\employer-state.json'
$previousJobsPath = Join-Path $root 'data\previous\jobs.json'
$changesPath = Join-Path $root 'data\changes\latest.json'
$auditDirectory = Join-Path $root 'data\audit'
$workbookPath = Join-Path $root 'input\us_health_systems_2026_current_map.xlsx'

function Ensure-EmployerRegistry {
    $registry = Get-JsonFile -Path $employersPath -Default $null
    if (-not $registry -or -not $registry.employers -or @($registry.employers).Count -eq 0) {
        Write-PipelineLog 'Employer registry is empty. Importing workbook.'
        & (Join-Path $PSScriptRoot 'Import-Employers.ps1') -WorkbookPath $workbookPath -OutputPath $employersPath | Out-Null
    }
}

function Rotate-AndCompare {
    param([Parameter(Mandatory)]$NewPayload)

    if (Test-Path -LiteralPath $currentJobsPath) {
        $existingCurrent = Get-JsonFile -Path $currentJobsPath -Default ([pscustomobject]@{ jobs = @() })
        Set-JsonFile -Path $previousJobsPath -Value $existingCurrent -Depth 12
    }
    Set-JsonFile -Path $currentJobsPath -Value $NewPayload -Depth 12
    & (Join-Path $PSScriptRoot 'Compare-JobSnapshots.ps1') -CurrentPath $currentJobsPath -PreviousPath $previousJobsPath -OutputPath $changesPath | Out-Null
}

function Write-AuditRow {
    param([Parameter(Mandatory)][string]$RunMode)

    $current = Get-JsonFile -Path $currentJobsPath -Default ([pscustomobject]@{ jobs = @() })
    $changes = Get-JsonFile -Path $changesPath -Default ([pscustomobject]@{ newJobCount = 0; removedJobCount = 0 })
    $state = Get-JsonFile -Path $currentStatePath -Default ([pscustomobject]@{ employers = @() })
    $registry = Get-JsonFile -Path $employersPath -Default ([pscustomobject]@{ employers = @() })

    if (-not (Test-Path -LiteralPath $auditDirectory)) { New-Item -ItemType Directory -Path $auditDirectory -Force | Out-Null }
    $auditPath = Join-Path $auditDirectory ((Get-Date).ToString('yyyy-MM-dd') + '.csv')
    $row = [pscustomobject][ordered]@{
        RunAtUtc = (Get-Date).ToUniversalTime().ToString('o')
        Mode = $RunMode
        Employers = @($registry.employers).Count
        EmployersChecked = @($state.employers).Count
        CurrentRoles = @($current.jobs).Count
        NewRoles = [int]$changes.newJobCount
        RemovedRolesAuditOnly = [int]$changes.removedJobCount
        ApplyChanges = [bool]$ApplyChanges
    }

    if (Test-Path -LiteralPath $auditPath) {
        $row | Export-Csv -LiteralPath $auditPath -NoTypeInformation -Append
    } else {
        $row | Export-Csv -LiteralPath $auditPath -NoTypeInformation
    }
}

Ensure-EmployerRegistry
$keywords = Get-JsonFile -Path $keywordsPath -Default $null
if (-not $keywords) { throw "Keywords file is empty: $keywordsPath" }

switch ($Mode) {
    'DeepDiscovery' {
        & (Join-Path $PSScriptRoot 'Discover-CareerSources.ps1') -EmployersPath $employersPath -SettingsPath $settingsPath -OutputPath $employersPath -RefreshExisting | Out-Null
        Write-PipelineLog 'Deep discovery complete.' -Level Success
    }

    'Hourly' {
        $hourlyJobsPath = if ($ShardCount -eq 1) { Join-Path $root 'data\current\jobs-hourly-temp.json' } else { $currentJobsPath }
        & (Join-Path $PSScriptRoot 'Monitor-CareerSources.ps1') `
            -EmployersPath $employersPath `
            -SettingsPath $settingsPath `
            -KeywordsPath $keywordsPath `
            -OutputPath $hourlyJobsPath `
            -EmployerStatePath $currentStatePath `
            -ShardIndex $ShardIndex `
            -ShardCount $ShardCount | Out-Null

        if ($ShardCount -eq 1) {
            $hourlyPayload = Get-JsonFile -Path $hourlyJobsPath -Default ([pscustomobject]@{ jobs = @() })
            Rotate-AndCompare -NewPayload $hourlyPayload
            Remove-Item -LiteralPath $hourlyJobsPath -Force -ErrorAction SilentlyContinue
            if ($ApplyChanges) { & (Join-Path $PSScriptRoot 'Sync-ClickUp.ps1') | Out-Null }
            Write-AuditRow -RunMode $Mode
        }
    }

    'HourlyAggregate' {
        if ([string]::IsNullOrWhiteSpace($ArtifactDirectory)) { throw 'ArtifactDirectory is required for HourlyAggregate mode.' }
        $jobFiles = @(Get-ChildItem -LiteralPath $ArtifactDirectory -Recurse -Filter 'jobs-shard-*.json' -File)
        $stateFiles = @(Get-ChildItem -LiteralPath $ArtifactDirectory -Recurse -Filter 'state-shard-*.json' -File)
        if ($jobFiles.Count -eq 0) { throw "No jobs-shard JSON files found under $ArtifactDirectory" }

        $allJobs = @()
        foreach ($file in $jobFiles) {
            $payload = Get-JsonFile -Path $file.FullName -Default ([pscustomobject]@{ jobs = @() })
            $allJobs += @($payload.jobs)
        }

        $preValidationCount = $allJobs.Count
        $allJobs = @(
            $allJobs |
                Where-Object {
                    JobMatching\Test-JobRelevance `
                        -Title ([string]$_.roleTitle) `
                        -Text '' `
                        -Keywords $keywords
                }
        )
        $removedByValidation = $preValidationCount - $allJobs.Count
        if ($removedByValidation -gt 0) {
            Write-PipelineLog ("Aggregate relevance revalidation removed {0} artifact role(s)." -f $removedByValidation) -Level Warn
        }

        $allJobs = @(
            $allJobs |
                Group-Object { (ConvertTo-NormalizedText $_.employer) + '|' + (ConvertTo-NormalizedText $_.roleTitle) + '|' + ([string]$_.jobUrl).ToLowerInvariant() } |
                ForEach-Object { $_.Group | Select-Object -First 1 }
        )

        $allStates = @()
        foreach ($file in $stateFiles) {
            $payload = Get-JsonFile -Path $file.FullName -Default ([pscustomobject]@{ employers = @() })
            $allStates += @($payload.employers)
        }

        Rotate-AndCompare -NewPayload ([pscustomobject][ordered]@{
            generatedAt = (Get-Date).ToUniversalTime().ToString('o')
            shardCount = $jobFiles.Count
            jobs = @($allJobs)
        })
        Set-JsonFile -Path $currentStatePath -Value ([pscustomobject][ordered]@{
            generatedAt = (Get-Date).ToUniversalTime().ToString('o')
            employers = @($allStates)
        }) -Depth 10

        if ($ApplyChanges) { & (Join-Path $PSScriptRoot 'Sync-ClickUp.ps1') | Out-Null }
        Write-AuditRow -RunMode $Mode
        Write-PipelineLog ("Aggregated {0} role(s) from {1} shard(s)." -f $allJobs.Count, $jobFiles.Count) -Level Success
    }

    'Full' {
        & (Join-Path $PSScriptRoot 'Discover-CareerSources.ps1') -EmployersPath $employersPath -SettingsPath $settingsPath -OutputPath $employersPath -RefreshExisting | Out-Null

        $tempJobs = Join-Path $root 'data\current\jobs-full-temp.json'
        $tempState = Join-Path $root 'data\current\state-full-temp.json'
        & (Join-Path $PSScriptRoot 'Monitor-CareerSources.ps1') `
            -EmployersPath $employersPath `
            -SettingsPath $settingsPath `
            -KeywordsPath $keywordsPath `
            -OutputPath $tempJobs `
            -EmployerStatePath $tempState `
            -DeepSearch | Out-Null

        $payload = Get-JsonFile -Path $tempJobs -Default ([pscustomobject]@{ jobs = @() })
        Rotate-AndCompare -NewPayload $payload
        $statePayload = Get-JsonFile -Path $tempState -Default ([pscustomobject]@{ employers = @() })
        Set-JsonFile -Path $currentStatePath -Value $statePayload -Depth 10
        Remove-Item -LiteralPath $tempJobs, $tempState -Force -ErrorAction SilentlyContinue

        if ($ApplyChanges) { & (Join-Path $PSScriptRoot 'Sync-ClickUp.ps1') | Out-Null }
        Write-AuditRow -RunMode $Mode
        Write-PipelineLog 'Full discovery and monitoring pass complete.' -Level Success
    }
}
