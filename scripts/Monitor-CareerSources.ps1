[CmdletBinding()]
param(
    [string]$EmployersPath = (Join-Path $PSScriptRoot '..\config\employers.json'),
    [string]$SettingsPath = (Join-Path $PSScriptRoot '..\config\settings.json'),
    [string]$KeywordsPath = (Join-Path $PSScriptRoot '..\config\keywords.json'),
    [string]$OutputPath = (Join-Path $PSScriptRoot '..\data\current\jobs.json'),
    [string]$EmployerStatePath = '',
    [int]$ShardIndex = 0,
    [int]$ShardCount = 1,
    [switch]$DeepSearch
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$utilitiesModule = Join-Path $PSScriptRoot '..\modules\Utilities.psm1'
$careerSourcesModule = Join-Path $PSScriptRoot '..\modules\CareerSources.psm1'

Import-Module $careerSourcesModule -Force -ErrorAction Stop
Import-Module $utilitiesModule -Force -ErrorAction Stop

$registry = Utilities\Get-JsonFile -Path $EmployersPath -Default $null
$settings = Utilities\Get-JsonFile -Path $SettingsPath -Default $null
$keywords = Utilities\Get-JsonFile -Path $KeywordsPath -Default $null

if (-not $registry -or -not $registry.employers) { throw "Employer registry is empty: $EmployersPath" }
if (-not $settings) { throw "Settings file is empty: $SettingsPath" }
if (-not $keywords) { throw "Keywords file is empty: $KeywordsPath" }
if ($ShardCount -lt 1) { throw 'ShardCount must be at least 1.' }
if ($ShardIndex -lt 0 -or $ShardIndex -ge $ShardCount) { throw 'ShardIndex must be between 0 and ShardCount - 1.' }

$jobs = @()
$states = @()
$employers = @($registry.employers)

for ($index = 0; $index -lt $employers.Count; $index++) {
    if (($index % $ShardCount) -ne $ShardIndex) { continue }

    $employer = $employers[$index]
    $isActive = $true
    if ($null -ne $employer.PSObject.Properties['active']) {
        $isActive = [bool]$employer.active
    }
    if (-not $isActive) { continue }

    $started = Get-Date
    $status = 'monitored'
    $errorMessage = ''
    $employerJobs = @()

    if (-not (Utilities\Test-HttpUrl ([string]$employer.careersUrl))) {
        $status = 'source missing'
    }
    else {
        try {
            Utilities\Write-PipelineLog (
                "Monitoring [{0}/{1}] {2}" -f `
                    ($index + 1),
                    $employers.Count,
                    [string]$employer.name
            )

            $employerJobs = @(
                CareerSources\Find-OfficialJobs `
                    -Employer $employer `
                    -Settings $settings `
                    -Keywords $keywords `
                    -DeepSearch:$DeepSearch
            )
            $jobs += $employerJobs
        }
        catch {
            $status = 'error'
            $errorMessage = $_.Exception.Message
            Utilities\Write-PipelineLog (
                "Monitor error for {0}: {1}" -f [string]$employer.name, $errorMessage
            ) -Level Warn
        }
    }

    $states += [pscustomobject][ordered]@{
        employer   = [string]$employer.name
        careersUrl = [string]$employer.careersUrl
        ats        = [string]$employer.ats
        status     = $status
        rolesFound = $employerJobs.Count
        checkedAt  = (Get-Date).ToUniversalTime().ToString('o')
        durationMs = [int]((Get-Date) - $started).TotalMilliseconds
        error      = $errorMessage
    }

    $delayMilliseconds = 0
    if ($null -ne $settings.request.PSObject.Properties['delayMilliseconds']) {
        $delayMilliseconds = [int]$settings.request.delayMilliseconds
    }
    if ($delayMilliseconds -gt 0) {
        Start-Sleep -Milliseconds $delayMilliseconds
    }
}

$payload = [pscustomobject][ordered]@{
    generatedAt = (Get-Date).ToUniversalTime().ToString('o')
    shardIndex  = $ShardIndex
    shardCount  = $ShardCount
    jobs        = @($jobs)
}

Utilities\Set-JsonFile -Path $OutputPath -Value $payload -Depth 12

if ($EmployerStatePath) {
    Utilities\Set-JsonFile `
        -Path $EmployerStatePath `
        -Value ([pscustomobject]@{
            generatedAt = (Get-Date).ToUniversalTime().ToString('o')
            shardIndex  = $ShardIndex
            shardCount  = $ShardCount
            employers   = @($states)
        }) `
        -Depth 10
}

Utilities\Write-PipelineLog (
    "Shard {0}/{1} complete. {2} qualifying official-source roles." -f `
        ($ShardIndex + 1),
        $ShardCount,
        $jobs.Count
) -Level Success

$payload
