Set-StrictMode -Version Latest

Import-Module (Join-Path $PSScriptRoot 'Utilities.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'JobMatching.psm1') -Force

function Search-BingRss {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Query,
        [int]$Count = 10,
        [int]$TimeoutSec = 30,
        [int]$Retries = 2
    )

    $encoded = [Uri]::EscapeDataString($Query)
    $url = "https://www.bing.com/search?q=$encoded&format=rss&count=$Count"
    $response = Invoke-WebRequestRetry -Uri $url -TimeoutSec $TimeoutSec -Retries $Retries
    if (-not $response -or [string]::IsNullOrWhiteSpace([string]$response.Content)) { return @() }

    try { [xml]$xml = [string]$response.Content } catch { return @() }
    $items = @($xml.rss.channel.item)
    $results = @()

    foreach ($item in $items | Select-Object -First $Count) {
        $rawUrl = [string]$item.link
        if (-not $rawUrl) { continue }
        $results += [pscustomobject]@{
            Title       = [System.Net.WebUtility]::HtmlDecode([string]$item.title)
            Url         = Resolve-SearchResultUrl -Url $rawUrl
            Description = ConvertFrom-HtmlText ([string]$item.description)
            Source      = 'Bing RSS'
        }
    }

    return @($results)
}

function Search-DuckDuckGoHtml {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Query,
        [int]$Count = 10,
        [int]$TimeoutSec = 30,
        [int]$Retries = 2
    )

    $encoded = [Uri]::EscapeDataString($Query)
    $url = "https://html.duckduckgo.com/html/?q=$encoded"
    $response = Invoke-WebRequestRetry -Uri $url -TimeoutSec $TimeoutSec -Retries $Retries
    if (-not $response -or [string]::IsNullOrWhiteSpace([string]$response.Content)) { return @() }

    $html = [string]$response.Content
    $pattern = '(?is)<a[^>]+class="[^"]*result__a[^"]*"[^>]+href="([^"]+)"[^>]*>(.*?)</a>.*?(?:class="[^"]*result__snippet[^"]*"[^>]*>(.*?)</(?:a|div|span)>)?'
    $matches = [regex]::Matches($html, $pattern)
    $results = @()

    foreach ($match in $matches | Select-Object -First $Count) {
        $results += [pscustomobject]@{
            Title       = ConvertFrom-HtmlText $match.Groups[2].Value
            Url         = Resolve-SearchResultUrl -Url ([System.Net.WebUtility]::HtmlDecode($match.Groups[1].Value))
            Description = ConvertFrom-HtmlText $match.Groups[3].Value
            Source      = 'DuckDuckGo HTML'
        }
    }

    return @($results)
}

function Search-CareerWeb {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Query,
        [Parameter(Mandatory)]$Settings,
        [int]$Count = 10
    )

    $timeout = [int]$Settings.request.timeoutSeconds
    $retries = [int]$Settings.request.retries
    $results = @(Search-BingRss -Query $Query -Count $Count -TimeoutSec $timeout -Retries $retries)
    if ($results.Count -eq 0) {
        $results = @(Search-DuckDuckGoHtml -Query $Query -Count $Count -TimeoutSec $timeout -Retries $retries)
    }
    return @($results)
}

function Get-AtsNameFromUrl {
    [CmdletBinding()]
    param(
        [AllowNull()][string]$Url,
        [Parameter(Mandatory)]$Settings
    )

    $host = Get-UriHostKey $Url
    if (-not $host) { return 'unknown' }

    foreach ($property in $Settings.atsNames.PSObject.Properties) {
        $domain = [string]$property.Name
        if ($host -eq $domain -or $host.EndsWith('.' + $domain)) {
            return [string]$property.Value
        }
    }

    return 'custom'
}

function Test-BlockedDomain {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Url,
        [Parameter(Mandatory)]$Settings
    )

    $host = Get-UriHostKey $Url
    if (-not $host) { return $true }
    foreach ($blocked in @($Settings.aggregatorDomains)) {
        $domain = ([string]$blocked).ToLowerInvariant()
        if ($host -eq $domain -or $host.EndsWith('.' + $domain)) { return $true }
    }
    return $false
}

function Find-CareerSource {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Employer,
        [Parameter(Mandatory)]$Settings
    )

    if (Test-HttpUrl ([string]$Employer.careersUrl)) {
        return [pscustomobject]@{
            careersUrl = [string]$Employer.careersUrl
            ats        = if ($Employer.ats) { [string]$Employer.ats } else { Get-AtsNameFromUrl -Url ([string]$Employer.careersUrl) -Settings $Settings }
            source     = 'registry'
        }
    }

    $query = '"' + [string]$Employer.name + '" careers jobs'
    $results = @(Search-CareerWeb -Query $query -Settings $Settings -Count ([int]$Settings.search.careerResultCount))
    $verificationDomain = Get-UriHostKey ([string]$Employer.verificationSource)
    $scored = @()

    foreach ($result in $results) {
        if (-not (Test-HttpUrl $result.Url)) { continue }
        if (Test-BlockedDomain -Url $result.Url -Settings $Settings) { continue }

        $host = Get-UriHostKey $result.Url
        $evidence = ([string]$result.Title) + ' ' + ([string]$result.Description)
        $score = 0
        $officialSignal = $false

        if ($verificationDomain -and ($host -eq $verificationDomain -or $host.EndsWith('.' + $verificationDomain))) {
            $score += 7
            $officialSignal = $true
        }

        foreach ($atsDomain in @($Settings.officialAtsDomains)) {
            $ats = ([string]$atsDomain).ToLowerInvariant()
            if ($host -eq $ats -or $host.EndsWith('.' + $ats)) {
                if (Test-CompanyIdentity -Company ([string]$Employer.name) -Text $evidence) {
                    $score += 8
                    $officialSignal = $true
                }
                break
            }
        }

        foreach ($token in @(Get-CompanyMatchTokens -Company ([string]$Employer.name))) {
            if ($token.Length -ge 5 -and $host -match [regex]::Escape($token)) {
                if (Test-CompanyIdentity -Company ([string]$Employer.name) -Text $evidence) {
                    $score += 6
                    $officialSignal = $true
                }
                break
            }
        }

        if (-not $officialSignal) { continue }
        if ((ConvertTo-NormalizedText $result.Url) -match '\b(career|careers|job|jobs|employment)\b') { $score += 3 }
        if ((ConvertTo-NormalizedText $result.Title) -match '\b(career|careers|job|jobs|employment)\b') { $score += 2 }
        if (Test-CompanyIdentity -Company ([string]$Employer.name) -Text $evidence) { $score += 2 }

        if ($score -ge 5) {
            $scored += [pscustomobject]@{ Result = $result; Score = $score }
        }
    }

    $best = $scored | Sort-Object Score -Descending | Select-Object -First 1
    if (-not $best) { return $null }

    return [pscustomobject]@{
        careersUrl = [string]$best.Result.Url
        ats        = Get-AtsNameFromUrl -Url ([string]$best.Result.Url) -Settings $Settings
        source     = [string]$best.Result.Source
    }
}

function Get-RelevantLinksFromCareerPage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Employer,
        [Parameter(Mandatory)]$Settings,
        [Parameter(Mandatory)]$Keywords
    )

    if (-not (Test-HttpUrl ([string]$Employer.careersUrl))) { return @() }
    $response = Invoke-WebRequestRetry -Uri ([string]$Employer.careersUrl) -TimeoutSec ([int]$Settings.request.timeoutSeconds) -Retries 1
    if (-not $response -or [string]::IsNullOrWhiteSpace([string]$response.Content)) { return @() }

    $html = [string]$response.Content
    $anchorPattern = '(?is)<a\b[^>]*href=["'']([^"'']+)["''][^>]*>(.*?)</a>'
    $matches = [regex]::Matches($html, $anchorPattern)
    $jobs = @()

    foreach ($match in $matches | Select-Object -First 500) {
        $title = ConvertFrom-HtmlText $match.Groups[2].Value
        if (-not (Test-JobRelevance -Title $title -Text '' -Keywords $Keywords)) { continue }
        if (-not (Test-LooksLikeRoleTitle -Title $title)) { continue }

        $url = Get-AbsoluteUrl -BaseUrl ([string]$Employer.careersUrl) -Href ([string]$match.Groups[1].Value)
        if (-not $url) { continue }
        if (-not (Test-OfficialJobUrl -Url $url -Employer $Employer -Settings $Settings -EvidenceText ($title + ' ' + [string]$Employer.name))) { continue }

        $jobs += [pscustomobject]@{
            employer       = [string]$Employer.name
            roleTitle      = $title
            workModel      = 'Not stated'
            salary         = 'Not disclosed'
            jobUrl         = $url
            careersUrl     = [string]$Employer.careersUrl
            source         = 'Official career page'
            verifiedAt     = (Get-Date).ToUniversalTime().ToString('o')
        }
    }

    return @($jobs)
}

function Find-OfficialJobs {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Employer,
        [Parameter(Mandatory)]$Settings,
        [Parameter(Mandatory)]$Keywords,
        [switch]$DeepSearch
    )

    if (-not (Test-HttpUrl ([string]$Employer.careersUrl))) { return @() }

    $jobs = @(Get-RelevantLinksFromCareerPage -Employer $Employer -Settings $Settings -Keywords $Keywords)
    $domain = Get-UriHostKey ([string]$Employer.careersUrl)
    if (-not $domain) { return @($jobs) }

    $queryBase = 'site:' + $domain + ' "' + [string]$Employer.name + '" '
    $queries = @(
        $queryBase + '(UKG OR Kronos OR Boomi OR HRIS OR "workforce management" OR WFM OR timekeeping OR "time and attendance")'
    )
    if ($DeepSearch) {
        $queries += $queryBase + '(HRIS OR UKG OR Kronos OR Boomi OR "workforce management") (analyst OR manager OR developer OR engineer OR administrator OR specialist OR director)'
    }

    foreach ($query in $queries) {
        $results = @(Search-CareerWeb -Query $query -Settings $Settings -Count ([int]$Settings.search.jobResultCount))
        foreach ($result in $results) {
            $evidence = ([string]$result.Title) + ' ' + ([string]$result.Description)
            if (-not (Test-JobRelevance -Title ([string]$result.Title) -Text ([string]$result.Description) -Keywords $Keywords)) { continue }

            $roleTitle = Get-JobTitleFromSearchTitle -Title ([string]$result.Title) -Company ([string]$Employer.name)
            if (-not (Test-LooksLikeRoleTitle -Title $roleTitle)) { continue }
            if (-not (Test-OfficialJobUrl -Url ([string]$result.Url) -Employer $Employer -Settings $Settings -EvidenceText $evidence)) { continue }

            $jobs += [pscustomobject]@{
                employer       = [string]$Employer.name
                roleTitle      = $roleTitle
                workModel      = Get-WorkModelFromText -Text $evidence
                salary         = Get-SalaryFromText -Text $evidence
                jobUrl         = [string]$result.Url
                careersUrl     = [string]$Employer.careersUrl
                source         = [string]$result.Source
                verifiedAt     = (Get-Date).ToUniversalTime().ToString('o')
            }
        }
    }

    return @(
        $jobs |
            Group-Object { (ConvertTo-NormalizedText $_.employer) + '|' + (ConvertTo-NormalizedText $_.roleTitle) + '|' + ([string]$_.jobUrl).ToLowerInvariant() } |
            ForEach-Object { $_.Group | Select-Object -First 1 }
    )
}

Export-ModuleMember -Function @(
    'Search-BingRss',
    'Search-DuckDuckGoHtml',
    'Search-CareerWeb',
    'Get-AtsNameFromUrl',
    'Test-BlockedDomain',
    'Find-CareerSource',
    'Get-RelevantLinksFromCareerPage',
    'Find-OfficialJobs'
)
