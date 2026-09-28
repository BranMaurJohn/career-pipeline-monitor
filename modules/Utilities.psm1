Set-StrictMode -Version Latest

function Write-PipelineLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('Info','Warn','Error','Success')][string]$Level = 'Info'
    )

    $timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $color = switch ($Level) {
        'Warn'    { 'Yellow' }
        'Error'   { 'Red' }
        'Success' { 'Green' }
        default   { 'Cyan' }
    }
    Write-Host ("[{0}] {1}" -f $timestamp, $Message) -ForegroundColor $color
}

function ConvertTo-NormalizedText {
    [CmdletBinding()]
    param([AllowNull()][string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $value = [System.Net.WebUtility]::HtmlDecode($Text).ToLowerInvariant()
    $value = [regex]::Replace($value, '[^a-z0-9]+', ' ')
    return ([regex]::Replace($value, '\s+', ' ').Trim())
}

function ConvertTo-NormalizedCompanyName {
    [CmdletBinding()]
    param([AllowNull()][string]$Name)

    $value = ConvertTo-NormalizedText $Name
    if (-not $value) { return '' }

    $remove = @(
        'health system','health systems','healthcare system','healthcare','health',
        'medical center','medical centers','hospital','hospitals','network','services',
        'inc','llc','corp','corporation','company','the'
    )

    foreach ($word in $remove) {
        $pattern = '(^|\s)' + [regex]::Escape($word) + '($|\s)'
        $value = [regex]::Replace($value, $pattern, ' ')
    }

    return ([regex]::Replace($value, '\s+', ' ').Trim())
}

function Test-HttpUrl {
    [CmdletBinding()]
    param([AllowNull()][string]$Url)
    return (-not [string]::IsNullOrWhiteSpace($Url) -and $Url -match '^https?://')
}

function Get-UriHostKey {
    [CmdletBinding()]
    param([AllowNull()][string]$Url)

    if (-not (Test-HttpUrl $Url)) { return '' }
    try {
        $host = ([Uri]$Url).Host.ToLowerInvariant()
        if ($host.StartsWith('www.')) { $host = $host.Substring(4) }
        return $host
    } catch {
        return ''
    }
}

function Get-RegistrableDomain {
    [CmdletBinding()]
    param([AllowNull()][string]$Url)

    $host = Get-UriHostKey $Url
    if (-not $host) { return '' }

    $parts = @($host.Split('.') | Where-Object { $_ })
    if ($parts.Count -lt 2) { return $host }

    $knownSecondLevel = @('co.uk','org.uk','com.au','com.ca')
    $lastTwo = ($parts[($parts.Count - 2)] + '.' + $parts[($parts.Count - 1)])
    if ($knownSecondLevel -contains $lastTwo -and $parts.Count -ge 3) {
        return ($parts[($parts.Count - 3)] + '.' + $lastTwo)
    }
    return $lastTwo
}

function Resolve-SearchResultUrl {
    [CmdletBinding()]
    param([AllowNull()][string]$Url)

    if ([string]::IsNullOrWhiteSpace($Url)) { return $Url }
    $decoded = [System.Net.WebUtility]::HtmlDecode($Url)

    try { $uri = [Uri]$decoded } catch { return $decoded }

    if ($uri.Host -match '(^|\.)duckduckgo\.com$') {
        $match = [regex]::Match($uri.Query, '(?:^|[?&])uddg=([^&]+)')
        if ($match.Success) {
            $target = [Uri]::UnescapeDataString($match.Groups[1].Value)
            if ($target -match '^https?://') { return $target }
        }
    }

    if ($uri.Host -match '(^|\.)bing\.com$') {
        $match = [regex]::Match($uri.Query, '(?:^|[?&])u=([^&]+)')
        if ($match.Success) {
            $candidate = [Uri]::UnescapeDataString($match.Groups[1].Value)
            if ($candidate.StartsWith('a1') -and $candidate.Length -gt 2) {
                $base64 = $candidate.Substring(2).Replace('-', '+').Replace('_', '/')
                while (($base64.Length % 4) -ne 0) { $base64 += '=' }
                try {
                    $bytes = [Convert]::FromBase64String($base64)
                    $target = [Text.Encoding]::UTF8.GetString($bytes)
                    if ($target -match '^https?://') { return $target }
                } catch {}
            }
            if ($candidate -match '^https?://') { return $candidate }
        }
    }

    return $decoded
}

function Invoke-WebRequestRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [int]$TimeoutSec = 30,
        [int]$Retries = 2,
        [string]$UserAgent = 'CareerPipelineMonitor/1.0 (+https://github.com/BranMaurJohn/career-pipeline-monitor)'
    )

    for ($attempt = 1; $attempt -le $Retries; $attempt++) {
        try {
            $params = @{
                Uri         = $Uri
                TimeoutSec  = $TimeoutSec
                Headers     = @{
                    'User-Agent'      = $UserAgent
                    'Accept-Language' = 'en-US,en;q=0.9'
                }
                ErrorAction = 'Stop'
            }
            if ($PSVersionTable.PSVersion.Major -lt 6) { $params['UseBasicParsing'] = $true }
            return Invoke-WebRequest @params
        } catch {
            if ($attempt -eq $Retries) { return $null }
            Start-Sleep -Seconds ([Math]::Min(5, [Math]::Pow(2, $attempt)))
        }
    }
}

function ConvertFrom-HtmlText {
    [CmdletBinding()]
    param([AllowNull()][string]$Html)

    if ([string]::IsNullOrWhiteSpace($Html)) { return '' }
    $text = [regex]::Replace($Html, '(?is)<script\b[^>]*>.*?</script>', ' ')
    $text = [regex]::Replace($text, '(?is)<style\b[^>]*>.*?</style>', ' ')
    $text = [regex]::Replace($text, '(?is)<[^>]+>', ' ')
    $text = [System.Net.WebUtility]::HtmlDecode($text)
    return ([regex]::Replace($text, '\s+', ' ').Trim())
}

function Get-AbsoluteUrl {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$BaseUrl,
        [AllowNull()][string]$Href
    )

    if ([string]::IsNullOrWhiteSpace($Href)) { return $null }
    try {
        if ($Href -match '^https?://') { return $Href }
        return ([Uri]::new([Uri]$BaseUrl, $Href)).AbsoluteUri
    } catch {
        return $null
    }
}

function Get-JsonFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        $Default = $null
    )

    if (-not (Test-Path -LiteralPath $Path)) { return $Default }
    $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction SilentlyContinue
    if ([string]::IsNullOrWhiteSpace($raw)) { return $Default }
    return ($raw | ConvertFrom-Json)
}

function Set-JsonFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)]$Value,
        [int]$Depth = 20
    )

    $directory = Split-Path -Parent $Path
    if ($directory -and -not (Test-Path -LiteralPath $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }

    $json = $Value | ConvertTo-Json -Depth $Depth
    [System.IO.File]::WriteAllText($Path, $json + [Environment]::NewLine, (New-Object System.Text.UTF8Encoding($false)))
}

function Get-Sha256Text {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Text)

    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [Text.Encoding]::UTF8.GetBytes($Text)
        $hash = $sha.ComputeHash($bytes)
        return (($hash | ForEach-Object { $_.ToString('x2') }) -join '')
    } finally {
        $sha.Dispose()
    }
}

Export-ModuleMember -Function @(
    'Write-PipelineLog',
    'ConvertTo-NormalizedText',
    'ConvertTo-NormalizedCompanyName',
    'Test-HttpUrl',
    'Get-UriHostKey',
    'Get-RegistrableDomain',
    'Resolve-SearchResultUrl',
    'Invoke-WebRequestRetry',
    'ConvertFrom-HtmlText',
    'Get-AbsoluteUrl',
    'Get-JsonFile',
    'Set-JsonFile',
    'Get-Sha256Text'
)
