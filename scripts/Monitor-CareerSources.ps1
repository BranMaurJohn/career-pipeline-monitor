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
Import-Module (Join-Path $PSScriptRoot '..\modules\Utilities.psm1') -Force
Import-Module (Join-Path $PSScriptRoot '..\modules\CareerSources.psm1') -Force

$registry = Get-JsonFile -Path $EmployersPath -Default $null
$settings = Get-JsonFile -Path $SettingsPath -Default $null
$keywords = Get-JsonFile -Path $KeywordsPath -Default $null
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
    if (-not [bool]$employer.active) { continue }

    $started = Get-Date
    $status = 'monitored'
    $errorMessage = ''
    $employerJobs = @()

    if (-not (Test-HttpUrl ([string]$employer.careersUrl))) {
        $status = 'source missing'
    } else {
        try {
            Write-PipelineLog ("Monitoring [{0}/{1}] {2}" -f ($index + 1), $employers.Count, [string]$employer.name)
            $employerJobs = @(Find-OfficialJobs -Employer $employer -Settings $settings -Keywords $keywords -DeepSearch:$DeepSearch)
            $jobs += $employerJobs
        } catch {
            $status = 'error'
            $errorMessage = $_.Exception.Message
            Write-PipelineLog ("Monitor error for {0}: {1}" -f [string]$employer.name, $errorMessage) -Level Warn
        }
    }

    $states += [pscustomobject][ordered]@{
        employer     = [string]$employer.name
        careersUrl   = [string]$employer.careersUrl
        ats          = [string]$employer.ats
        status       = $status
        rolesFound   = $employerJobs.Count
        checkedAt    = (Get-Date).ToUniversalTime().ToString('o')
        durationMs   = [int]((Get-Date) - $started).TotalMilliseconds
        error        = $errorMessage
    }

    Start-Sleep -Milliseconds ([int]$settings.request.delayMilliseconds)
}

$payload = [pscustomobject][ordered]@{
    generatedAt = (Get-Date).ToUniversalTime().ToString('o')
    shardIndex  = $ShardIndex
    shardCount  = $ShardCount
    jobs        = @($jobs)
}
Set-JsonFile -Path $OutputPath -Value $payload -Depth 12

if ($EmployerStatePath) {
    Set-JsonFile -Path $EmployerStatePath -Value ([pscustomobject]@{
        generatedAt = (Get-Date).ToUniversalTime().ToString('o')
        shardIndex = $ShardIndex
        shardCount = $ShardCount
        employers = @($states)
    }) -Depth 10
}

Write-PipelineLog ("Shard {0}/{1} complete. {2} qualifying official-source roles." -f ($ShardIndex + 1), $ShardCount, $jobs.Count) -Level Success
$payload
