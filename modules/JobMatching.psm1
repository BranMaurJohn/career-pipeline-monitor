Set-StrictMode -Version Latest

$utilityModule = Join-Path $PSScriptRoot 'Utilities.psm1'
Import-Module $utilityModule -Force

function Get-CompanyMatchTokens {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Company)

    $stop = @(
        'health','healthcare','system','systems','medical','center','centers',
        'hospital','hospitals','network','management','services','service','care',
        'clinic','clinics','regional','university','the','and','of'
    )

    $tokens = @()
    foreach ($match in [regex]::Matches($Company, '[A-Za-z0-9]+')) {
        $raw = [string]$match.Value
        $normalized = $raw.ToLowerInvariant()
        if ($stop -contains $normalized) { continue }

        $isAcronym = ($raw -cmatch '^[A-Z0-9]{2,5}$')
        if ($normalized.Length -ge 4 -or $isAcronym) {
            $tokens += $normalized
        }
    }

    return @($tokens | Select-Object -Unique)
}

function Test-CompanyIdentity {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Company,
        [AllowNull()][string]$Text
    )

    $normalized = ConvertTo-NormalizedText $Text
    if (-not $normalized) { return $false }

    $aliases = @($Company -split '\s*[|/]\s*')
    foreach ($alias in $aliases) {
        $aliasNormalized = ConvertTo-NormalizedText ([string]$alias)
        if ($aliasNormalized -and $normalized.Contains($aliasNormalized)) {
            return $true
        }
    }

    $companyNormalized = ConvertTo-NormalizedText $Company
    if ($companyNormalized -and $normalized.Contains($companyNormalized)) {
        return $true
    }

    $tokens = @(Get-CompanyMatchTokens -Company $Company)
    if ($tokens.Count -lt 2) { return $false }

    $hits = 0
    foreach ($token in $tokens) {
        if ($normalized -match ('\b' + [regex]::Escape($token) + '\b')) {
            $hits++
        }
    }

    return ($hits -ge 2)
}

function Test-NormalizedPhraseMatch {
    [CmdletBinding()]
    param(
        [AllowNull()][string]$Haystack,
        [AllowNull()][string]$Needle
    )

    if ([string]::IsNullOrWhiteSpace($Haystack) -or [string]::IsNullOrWhiteSpace($Needle)) {
        return $false
    }

    $pattern = '(?<![a-z0-9])' + [regex]::Escape($Needle) + '(?![a-z0-9])'
    return [regex]::IsMatch($Haystack, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
}

function Test-JobRelevance {
    [CmdletBinding()]
    param(
        [AllowNull()][string]$Title,
        [AllowNull()][string]$Text,
        [Parameter(Mandatory)]$Keywords
    )

    $combined = ConvertTo-NormalizedText (([string]$Title) + ' ' + ([string]$Text))
    if (-not $combined) { return $false }

    foreach ($exclude in @($Keywords.exclude)) {
        $needle = ConvertTo-NormalizedText ([string]$exclude)
        if ($needle -and (Test-NormalizedPhraseMatch -Haystack $combined -Needle $needle)) { return $false }
    }

    $primaryHit = $false
    foreach ($term in @($Keywords.primary)) {
        $needle = ConvertTo-NormalizedText ([string]$term)
        if ($needle -and (Test-NormalizedPhraseMatch -Haystack $combined -Needle $needle)) {
            $primaryHit = $true
            break
        }
    }

    $secondaryHit = $false
    foreach ($term in @($Keywords.secondary)) {
        $needle = ConvertTo-NormalizedText ([string]$term)
        if ($needle -and (Test-NormalizedPhraseMatch -Haystack $combined -Needle $needle)) {
            $secondaryHit = $true
            break
        }
    }

    $roleNoun = $combined -match '\b(analyst|developer|administrator|engineer|architect|specialist|manager|director|lead|consultant|product owner|program manager|project manager)\b'

    return ($primaryHit -or ($secondaryHit -and $roleNoun))
}

function Test-LooksLikeRoleTitle {
    [CmdletBinding()]
    param([AllowNull()][string]$Title)

    $value = ConvertTo-NormalizedText $Title
    if (-not $value) { return $false }

    $bad = @(
        'home','careers','jobs','job search','search jobs','events','contact us',
        'albany new york','things to do','visitor guide'
    )
    if ($bad -contains $value) { return $false }
    if ($value -match '\b(tourism|visitor|travel|events|news|press release)\b') { return $false }

    return ($value -match '\b(analyst|developer|administrator|admin|manager|engineer|architect|specialist|director|lead|consultant|product owner|program manager|project manager)\b')
}

function Get-JobTitleFromSearchTitle {
    [CmdletBinding()]
    param(
        [AllowNull()][string]$Title,
        [Parameter(Mandatory)][string]$Company
    )

    if ([string]::IsNullOrWhiteSpace($Title)) { return '' }
    $clean = [System.Net.WebUtility]::HtmlDecode($Title).Trim()

    $separators = @(' | ', ' - Careers', ' - Jobs', ' at ' + $Company, ' - ' + $Company)
    foreach ($separator in $separators) {
        $index = $clean.IndexOf($separator, [System.StringComparison]::OrdinalIgnoreCase)
        if ($index -gt 2) { $clean = $clean.Substring(0, $index).Trim() }
    }

    return $clean
}

function Get-WorkModelFromText {
    [CmdletBinding()]
    param([AllowNull()][string]$Text)

    $value = ConvertTo-NormalizedText $Text
    if ($value -match '\bremote\b') { return 'Remote' }
    if ($value -match '\bhybrid\b') { return 'Hybrid' }
    if ($value -match '\b(on site|onsite|on-site)\b') { return 'On-site' }
    return 'Not stated'
}

function Get-SalaryFromText {
    [CmdletBinding()]
    param([AllowNull()][string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return 'Not disclosed' }

    $patterns = @(
        '\$\s?\d{2,3}(?:,\d{3})?(?:\.\d{2})?\s*(?:-|to|\u2013)\s*\$?\s?\d{2,3}(?:,\d{3})?(?:\.\d{2})?',
        '\$\s?\d{2,3}(?:\.\d{2})?\s*(?:-|to|\u2013)\s*\$?\s?\d{2,3}(?:\.\d{2})?\s*(?:/hr|per hour|hourly)'
    )

    foreach ($pattern in $patterns) {
        $match = [regex]::Match($Text, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($match.Success) { return $match.Value.Trim() }
    }

    return 'Not disclosed'
}

function Test-OfficialJobUrl {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Url,
        [Parameter(Mandatory)]$Employer,
        [Parameter(Mandatory)]$Settings,
        [AllowNull()][string]$EvidenceText
    )

    if (-not (Test-HttpUrl $Url)) { return $false }

    $host = Get-UriHostKey $Url
    if (-not $host) { return $false }

    foreach ($blocked in @($Settings.aggregatorDomains)) {
        $blockedHost = ([string]$blocked).ToLowerInvariant()
        if ($host -eq $blockedHost -or $host.EndsWith('.' + $blockedHost)) { return $false }
    }

    $careerDomain = Get-UriHostKey ([string]$Employer.careersUrl)
    $verificationDomain = Get-UriHostKey ([string]$Employer.verificationSource)

    $matchesEmployerDomain = $false
    foreach ($domain in @($careerDomain, $verificationDomain)) {
        if ($domain -and ($host -eq $domain -or $host.EndsWith('.' + $domain) -or $domain.EndsWith('.' + $host))) {
            $matchesEmployerDomain = $true
            break
        }
    }

    $isAts = $false
    foreach ($atsDomain in @($Settings.officialAtsDomains)) {
        $ats = ([string]$atsDomain).ToLowerInvariant()
        if ($host -eq $ats -or $host.EndsWith('.' + $ats)) {
            $isAts = $true
            break
        }
    }

    if (-not ($matchesEmployerDomain -or $isAts)) { return $false }

    if ($isAts -and -not $matchesEmployerDomain) {
        if (-not (Test-CompanyIdentity -Company ([string]$Employer.name) -Text $EvidenceText)) { return $false }
    }

    return $true
}

function Get-ClickUpRoleTaskName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Company,
        [Parameter(Mandatory)][string]$RoleTitle
    )
    return ('{0} - {1}' -f $Company.Trim(), $RoleTitle.Trim())
}

function Test-CanPromoteRoleStatus {
    [CmdletBinding()]
    param([AllowNull()][string]$Status)
    $value = ConvertTo-NormalizedText $Status
    return ($value -eq '' -or $value -eq 'target company')
}

Export-ModuleMember -Function @(
    'Get-CompanyMatchTokens',
    'Test-CompanyIdentity',
    'Test-NormalizedPhraseMatch',
    'Test-JobRelevance',
    'Test-LooksLikeRoleTitle',
    'Get-JobTitleFromSearchTitle',
    'Get-WorkModelFromText',
    'Get-SalaryFromText',
    'Test-OfficialJobUrl',
    'Get-ClickUpRoleTaskName',
    'Test-CanPromoteRoleStatus'
)
