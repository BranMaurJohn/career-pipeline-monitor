[CmdletBinding()]
param(
    [string]$EmployersPath = (Join-Path $PSScriptRoot '..\config\employers.json'),
    [string]$SettingsPath = (Join-Path $PSScriptRoot '..\config\settings.json'),
    [string]$KeywordsPath = (Join-Path $PSScriptRoot '..\config\keywords.json'),
    [string]$SourceOverridesPath = (Join-Path $PSScriptRoot '..\config\source-overrides.json'),
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
$jobMatchingModule = Join-Path $PSScriptRoot '..\modules\JobMatching.psm1'

Import-Module $careerSourcesModule -Force -ErrorAction Stop
Import-Module $jobMatchingModule -Force -ErrorAction Stop
Import-Module $utilitiesModule -Force -ErrorAction Stop

$registry = Utilities\Get-JsonFile -Path $EmployersPath -Default $null
$settings = Utilities\Get-JsonFile -Path $SettingsPath -Default $null
$keywords = Utilities\Get-JsonFile -Path $KeywordsPath -Default $null
$sourceOverrides = Utilities\Get-JsonFile -Path $SourceOverridesPath -Default $null

if (-not $registry -or -not $registry.employers) { throw "Employer registry is empty: $EmployersPath" }
if (-not $settings) { throw "Settings file is empty: $SettingsPath" }
if (-not $keywords) { throw "Keywords file is empty: $KeywordsPath" }
if ($ShardCount -lt 1) { throw 'ShardCount must be at least 1.' }
if ($ShardIndex -lt 0 -or $ShardIndex -ge $ShardCount) { throw 'ShardIndex must be between 0 and ShardCount - 1.' }

$overrideMap = @{}
if ($sourceOverrides -and $sourceOverrides.overrides) {
    foreach ($entry in @($sourceOverrides.overrides)) {
        $name = [string]$entry.name
        $url = [string]$entry.careersUrl
        if ([string]::IsNullOrWhiteSpace($name) -or -not (Utilities\Test-HttpUrl $url)) { continue }
        $overrideMap[$name] = $entry
    }
}

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

    $name = [string]$employer.name
    $hasVerifiedOverride = $overrideMap.ContainsKey($name)
    if ($hasVerifiedOverride) {
        $override = $overrideMap[$name]
        $employer.careersUrl = [string]$override.careersUrl
        if (-not [string]::IsNullOrWhiteSpace([string]$override.ats)) {
            $employer.ats = [string]$override.ats
        }
        Utilities\Write-PipelineLog ("Using verified source override for {0}: {1}" -f $name, [string]$employer.careersUrl)
    }

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
                    $name
            )

            $employerJobs = @(
                CareerSources\Find-OfficialJobs `
                    -Employer $employer `
                    -Settings $settings `
                    -Keywords $keywords `
                    -DeepSearch:$DeepSearch
            )

            # Branded source overrides are already verified as belonging to the employer.
            # Search that host directly without forcing the employer name into the query;
            # some official job pages do not repeat the corporate name in indexed snippets.
            if ($hasVerifiedOverride) {
                $domain = Utilities\Get-UriHostKey ([string]$employer.careersUrl)
                if ($domain) {
                    $overrideQueries = @(
                        ('site:' + $domain + ' (UKG OR Kronos OR Boomi OR HRIS OR "workforce management" OR WFM OR timekeeping OR "time and attendance" OR "advanced scheduling")')
                    )
                    if ($DeepSearch) {
                        $overrideQueries += ('site:' + $domain + ' (UKG OR Boomi OR HRIS OR "workforce management" OR timekeeping) (analyst OR manager OR developer OR engineer OR administrator OR specialist OR director)')
                    }

                    foreach ($query in $overrideQueries | Select-Object -Unique) {
                        $results = @(
                            CareerSources\Search-CareerWeb `
                                -Query $query `
                                -Settings $settings `
                                -Count ([int]$settings.search.jobResultCount)
                        )

                        foreach ($result in $results) {
                            $evidence = ([string]$result.Title) + ' ' + ([string]$result.Description)
                            if (-not (JobMatching\Test-JobRelevance -Title ([string]$result.Title) -Text ([string]$result.Description) -Keywords $keywords)) { continue }

                            $roleTitle = JobMatching\Get-JobTitleFromSearchTitle -Title ([string]$result.Title) -Company $name
                            if (-not (JobMatching\Test-LooksLikeRoleTitle -Title $roleTitle)) { continue }
                            if (-not (JobMatching\Test-OfficialJobUrl -Url ([string]$result.Url) -Employer $employer -Settings $settings -EvidenceText $evidence)) { continue }

                            $employerJobs += [pscustomobject]@{
                                employer   = $name
                                roleTitle  = $roleTitle
                                workModel  = JobMatching\Get-WorkModelFromText -Text $evidence
                                salary     = JobMatching\Get-SalaryFromText -Text $evidence
                                jobUrl     = [string]$result.Url
                                careersUrl = [string]$employer.careersUrl
                                source     = [string]$result.Source
                                verifiedAt = (Get-Date).ToUniversalTime().ToString('o')
                            }
                        }
                    }
                }
            }

            $employerJobs = @(
                $employerJobs |
                    Group-Object { (Utilities\ConvertTo-NormalizedText $_.roleTitle) + '|' + ([string]$_.jobUrl).ToLowerInvariant() } |
                    ForEach-Object { $_.Group | Select-Object -First 1 }
            )
            $jobs += $employerJobs
        }
        catch {
            $status = 'error'
            $errorMessage = $_.Exception.Message
            Utilities\Write-PipelineLog (
                "Monitor error for {0}: {1}" -f $name, $errorMessage
            ) -Level Warn
        }
    }

    $states += [pscustomobject][ordered]@{
        employer   = $name
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
