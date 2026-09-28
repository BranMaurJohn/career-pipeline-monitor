BeforeAll {
    $root = Resolve-Path (Join-Path $PSScriptRoot '..')

    Import-Module (Join-Path $root 'modules\Utilities.psm1') `
        -Force `
        -Global

    Import-Module (Join-Path $root 'modules\JobMatching.psm1') `
        -Force `
        -Global

    $script:keywords = Get-JsonFile `
        -Path (Join-Path $root 'config\keywords.json')

    $script:settings = Get-JsonFile `
        -Path (Join-Path $root 'config\settings.json')

    if ($null -eq $script:keywords) {
        throw 'Unable to load config\keywords.json.'
    }

    if ($null -eq $script:settings) {
        throw 'Unable to load config\settings.json.'
    }
}

Describe 'JobMatching' {

    It 'matches a UKG HRIS role' {
        Test-JobRelevance `
            -Title 'Senior HRIS Analyst - UKG' `
            -Text 'Support UKG Pro WFM and timekeeping.' `
            -Keywords $script:keywords |
            Should -BeTrue
    }

    It 'matches a Boomi integration role' {
        Test-JobRelevance `
            -Title 'Integration Developer' `
            -Text 'Build Boomi integrations for HR technology platforms.' `
            -Keywords $script:keywords |
            Should -BeTrue
    }

    It 'rejects patient scheduling roles' {
        Test-JobRelevance `
            -Title 'Patient Scheduler' `
            -Text 'Schedule patient appointments.' `
            -Keywords $script:keywords |
            Should -BeFalse
    }

    It 'rejects a tourism/location title as a role title' {
        Test-LooksLikeRoleTitle `
            -Title 'Albany, New York' |
            Should -BeFalse
    }

    It 'accepts an official employer-domain job URL' {
        $employer = [pscustomobject]@{
            name               = 'Example Health'
            careersUrl         = 'https://careers.examplehealth.org/jobs'
            verificationSource = 'https://www.examplehealth.org/'
        }

        Test-OfficialJobUrl `
            -Url 'https://careers.examplehealth.org/jobs/123/senior-hris-analyst' `
            -Employer $employer `
            -Settings $script:settings `
            -EvidenceText 'Example Health Senior HRIS Analyst UKG' |
            Should -BeTrue
    }

    It 'rejects Indeed as an official source' {
        $employer = [pscustomobject]@{
            name               = 'Example Health'
            careersUrl         = 'https://careers.examplehealth.org/jobs'
            verificationSource = 'https://www.examplehealth.org/'
        }

        Test-OfficialJobUrl `
            -Url 'https://www.indeed.com/viewjob?jk=123' `
            -Employer $employer `
            -Settings $script:settings `
            -EvidenceText 'Example Health UKG Analyst' |
            Should -BeFalse
    }
}
