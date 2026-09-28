BeforeAll {
    $root = Resolve-Path (Join-Path $PSScriptRoot '..')

    Import-Module (Join-Path $root 'modules\JobMatching.psm1') `
        -Force `
        -Global

    Import-Module (Join-Path $root 'modules\CareerSources.psm1') `
        -Force `
        -Global

    Import-Module (Join-Path $root 'modules\Utilities.psm1') `
        -Force `
        -Global

    $script:settings = Utilities\Get-JsonFile `
        -Path (Join-Path $root 'config\settings.json')

    if ($null -eq $script:settings) {
        throw 'Unable to load config\settings.json.'
    }
}

Describe 'CareerSources' {

    It 'identifies Workday from an ATS URL' {
        CareerSources\Get-AtsNameFromUrl `
            -Url 'https://example.wd5.myworkdayjobs.com/en-US/jobs' `
            -Settings $script:settings |
            Should -Be 'Workday'
    }

    It 'blocks job aggregators' {
        CareerSources\Test-BlockedDomain `
            -Url 'https://www.indeed.com/viewjob?jk=123' `
            -Settings $script:settings |
            Should -BeTrue
    }

    It 'does not block an official health system domain' {
        CareerSources\Test-BlockedDomain `
            -Url 'https://careers.examplehealth.org/jobs/123' `
            -Settings $script:settings |
            Should -BeFalse
    }

    It 'rejects an unrelated university with the same city token' {
        $employer = [pscustomobject]@{
            name               = 'Albany Med Health System'
            verificationSource = ''
            careersUrl         = ''
            ats                = ''
        }

        CareerSources\Test-CareerSourceCandidate `
            -Employer $employer `
            -Url 'https://www.albany.edu/jobs' `
            -Title 'University at Albany Careers' `
            -Description 'University employment opportunities in Albany.' `
            -Settings $script:settings |
            Should -BeFalse
    }

    It 'rejects an unrelated publisher with an overlapping brand token' {
        $employer = [pscustomobject]@{
            name               = 'Atlantic Health System'
            verificationSource = ''
            careersUrl         = ''
            ats                = ''
        }

        CareerSources\Test-CareerSourceCandidate `
            -Employer $employer `
            -Url 'https://www.theatlantic.com/jobs/' `
            -Title 'Jobs at The Atlantic' `
            -Description 'Atlantic careers and employment opportunities.' `
            -Settings $script:settings |
            Should -BeFalse
    }

    It 'accepts an official branded career site' {
        $employer = [pscustomobject]@{
            name               = 'AdventHealth'
            verificationSource = ''
            careersUrl         = ''
            ats                = ''
        }

        CareerSources\Test-CareerSourceCandidate `
            -Employer $employer `
            -Url 'https://jobs.adventhealth.com/' `
            -Title 'AdventHealth Careers' `
            -Description 'Search AdventHealth jobs and career opportunities.' `
            -Settings $script:settings |
            Should -BeTrue
    }

    It 'accepts a company-identified official ATS result' {
        $employer = [pscustomobject]@{
            name               = 'Example Health'
            verificationSource = ''
            careersUrl         = ''
            ats                = ''
        }

        CareerSources\Test-CareerSourceCandidate `
            -Employer $employer `
            -Url 'https://example.wd5.myworkdayjobs.com/en-US/jobs' `
            -Title 'Example Health Careers' `
            -Description 'Search Example Health jobs.' `
            -Settings $script:settings |
            Should -BeTrue
    }

    It 'unwraps a DuckDuckGo destination URL' {
        $target = 'https%3A%2F%2Fcareers.example.org%2Fjobs%2F123'

        Utilities\Resolve-SearchResultUrl `
            -Url ('https://duckduckgo.com/l/?uddg=' + $target) |
            Should -Be 'https://careers.example.org/jobs/123'
    }
}
