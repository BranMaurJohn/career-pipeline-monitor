BeforeAll {
    $root = Resolve-Path (Join-Path $PSScriptRoot '..')
    Import-Module (Join-Path $root 'modules\Utilities.psm1') -Force
    Import-Module (Join-Path $root 'modules\CareerSources.psm1') -Force
    $script:settings = Get-JsonFile -Path (Join-Path $root 'config\settings.json')
}

Describe 'CareerSources' {
    It 'identifies Workday from an ATS URL' {
        Get-AtsNameFromUrl -Url 'https://example.wd5.myworkdayjobs.com/en-US/jobs' -Settings $script:settings | Should -Be 'Workday'
    }

    It 'blocks job aggregators' {
        Test-BlockedDomain -Url 'https://www.indeed.com/viewjob?jk=123' -Settings $script:settings | Should -BeTrue
    }

    It 'does not block an official health system domain' {
        Test-BlockedDomain -Url 'https://careers.examplehealth.org/jobs/123' -Settings $script:settings | Should -BeFalse
    }

    It 'unwraps a DuckDuckGo destination URL' {
        $target = 'https%3A%2F%2Fcareers.example.org%2Fjobs%2F123'
        Resolve-SearchResultUrl -Url ('https://duckduckgo.com/l/?uddg=' + $target) | Should -Be 'https://careers.example.org/jobs/123'
    }
}
