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

function Test-OfficialAtsHost {
    [CmdletBinding()]
    param(
        [AllowNull()][string]$Url,
        [Parameter(Mandatory)]$Settings
    )

    $host = Get-UriHostKey $Url
    if (-not $host) { return $false }

    foreach ($atsDomain in @($Settings.officialAtsDomains)) {
        $ats = ([string]$atsDomain).ToLowerInvariant()
        if ($host -eq $ats -or $host.EndsWith('.' + $ats)) { return $true }
    }

    return $false
}

function Test-EmployerDomainMatch {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Company,
        [Parameter(Mandatory)][string]$Url
    )

    $host = Get-UriHostKey $Url
    if (-not $host) { return $false }
    $hostCompact = [regex]::Replace($host.ToLowerInvariant(), '[^a-z0-9]', '')

    $aliases = @($Company -split '\s*[|/]\s*')
    foreach ($alias in $aliases) {
        $tokens = @(Get-CompanyMatchTokens -Company ([string]$alias))
        if ($tokens.Count -eq 0) { continue }

        $hits = 0
        foreach ($token in $tokens) {
            $compact = [regex]::Replace(([string]$token).ToLowerInvariant(), '[^a-z0-9]', '')
            if ($compact -and $hostCompact.Contains($compact)) { $hits++ }
        }

        if ($tokens.Count -eq 1 -and $hits -eq 1) { return $true }
        if ($tokens.Count -ge 2 -and $hits -ge 2) { return $true }
    }

    return $false
}

function Test-CareerSourceCandidate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Employer,
        [Parameter(Mandatory)][string]$Url,
        [AllowNull()][string]$Title,
        [AllowNull()][string]$Description,
        [Parameter(Mandatory)]$Settings
    )

    if (-not (Test-HttpUrl $Url)) { return $false }
    if (Test-BlockedDomain -Url $Url -Settings $Settings) { return $false }

    $evidence = (([string]$Title) + ' ' + ([string]$Description)).Trim()
    if (-not (Test-CompanyIdentity -Company ([string]$Employer.name) -Text $evidence)) {
        return $false
    }

    $careerSignalText = (([string]$Url) + ' ' + ([string]$Title))
    $hasCareerSignal = $careerSignalText -match '(?i)(career|job|employment|candidate|requisition|opportunit|join[-_ ]?our[-_ ]?team|work[-_ ]?with[-_ ]?us)'
    if (-not $hasCareerSignal) { return $false }

    $isAts = Test-OfficialAtsHost -Url $Url -Settings $Settings
    if ($isAts) { return $true }

    $host = Get-UriHostKey $Url
    $verificationHost = Get-UriHostKey ([string]$Employer.verificationSource)
    $matchesVerificationHost = $verificationHost -and (
        $host -eq $verificationHost -or
        $host.EndsWith('.' + $verificationHost) -or
        $verificationHost.EndsWith('.' + $host)
    )
    if ($matchesVerificationHost) { return $true }

    if (-not (Test-EmployerDomainMatch -Company ([string]$Employer.name) -Url $Url)) {
        return $false
    }

    # For a custom domain whose employer identity reduces to only one usable
    # brand token, require the actual search-result title to name the health
    # organization. This stops collisions such as Albany.edu, TheAtlantic.com,
    # Delta.com, BannerEngineering.com, and similarly named unrelated sites.
    $tokens = @(Get-CompanyMatchTokens -Company ([string]$Employer.name))
    if ($tokens.Count -lt 2) {
        return (Test-CompanyIdentity -Company ([string]$Employer.name) -Text ([string]$Title))
    }

    return $true
}

function Find-CareerSource {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Employer,
        [Parameter(Mandatory)]$Settings,
        [switch]$IgnoreRegistry
    )

    if (-not $IgnoreRegistry -and (Test-HttpUrl ([string]$Employer.careersUrl))) {
        return [pscustomobject]@{
            careersUrl = [string]$Employer.careersUrl
            ats        = if ($Employer.ats) { [string]$Employer.ats } else { Get-AtsNameFromUrl -Url ([string]$Employer.careersUrl) -Settings $Settings }
            source     = 'registry'
        }
    }

    $queries = @()
    $queries += ('"' + [string]$Employer.name + '" careers jobs employment')

    foreach ($propertyName in @('currentCompany','parentEnterprise','operatingBrandRegion')) {
        if ($null -ne $Employer.PSObject.Properties[$propertyName]) {
            $alias = [string]$Employer.$propertyName
            if (-not [string]::IsNullOrWhiteSpace($alias) -and $alias -ne [string]$Employer.name) {
                $queries += ('"' + $alias + '" careers jobs employment')
            }
        }
    }

    $verificationHost = Get-UriHostKey ([string]$Employer.verificationSource)
    if ($verificationHost) {
        $queries += ('site:' + $verificationHost + ' careers jobs employment')
    }

    $results = @()
    foreach ($query in @($queries | Select-Object -Unique)) {
        $results += @(Search-CareerWeb -Query $query -Settings $Settings -Count ([int]$Settings.search.careerResultCount))
    }

    $results = @(
        $results |
            Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.Url) } |
            Group-Object { ([string]$_.Url).ToLowerInvariant() } |
            ForEach-Object { $_.Group | Select-Object -First 1 }
    )

    $scored = @()
    foreach ($result in $results) {
        if (-not (Test-CareerSourceCandidate `
            -Employer $Employer `
            -Url ([string]$result.Url) `
            -Title ([string]$result.Title) `
            -Description ([string]$result.Description) `
            -Settings $Settings)) {
            continue
        }

        $host = Get-UriHostKey ([string]$result.Url)
        $score = 0

        if (Test-OfficialAtsHost -Url ([string]$result.Url) -Settings $Settings) { $score += 100 }
        if ($verificationHost -and ($host -eq $verificationHost -or $host.EndsWith('.' + $verificationHost))) { $score += 80 }
        if (Test-EmployerDomainMatch -Company ([string]$Employer.name) -Url ([string]$result.Url)) { $score += 60 }
        if ([string]$result.Url -match '(?i)(career|job|employment|candidate|requisition)') { $score += 20 }
        if ([string]$result.Title -match '(?i)(career|job|employment|candidate|requisition)') { $score += 10 }
        if (Test-CompanyIdentity -Company ([string]$Employer.name) -Text ([string]$result.Title)) { $score += 10 }

        $scored += [pscustomobject]@{ Result = $result; Score = $score }
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
            employer   = [string]$Employer.name
            roleTitle  = $title
            workModel  = 'Not stated'
            salary     = 'Not disclosed'
            jobUrl     = $url
            careersUrl = [string]$Employer.careersUrl
            source     = 'Official career page'
            verifiedAt = (Get-Date).ToUniversalTime().ToString('o')
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
                employer   = [string]$Employer.name
                roleTitle  = $roleTitle
                workModel  = Get-WorkModelFromText -Text $evidence
                salary     = Get-SalaryFromText -Text $evidence
                jobUrl     = [string]$result.Url
                careersUrl = [string]$Employer.careersUrl
                source     = [string]$result.Source
                verifiedAt = (Get-Date).ToUniversalTime().ToString('o')
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
    'Test-CareerSourceCandidate',
    'Find-CareerSource',
    'Get-RelevantLinksFromCareerPage',
    'Find-OfficialJobs'
)
