BeforeAll {
    $root = Resolve-Path (Join-Path $PSScriptRoot '..')

    Import-Module (Join-Path $root 'modules\JobMatching.psm1') `
        -Force `
        -Global

    Import-Module (Join-Path $root 'modules\Utilities.psm1') `
        -Force `
        -Global

    $script:keywords = Utilities\Get-JsonFile `
        -Path (Join-Path $root 'config\keywords.json')

    $script:settings = Utilities\Get-JsonFile `
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
        JobMatching\Test-JobRelevance `
            -Title 'Senior HRIS Analyst - UKG' `
            -Text 'Support UKG Pro WFM and timekeeping.' `
            -Keywords $script:keywords |
            Should -BeTrue
    }

    It 'matches a Boomi integration role' {
        JobMatching\Test-JobRelevance `
            -Title 'Integration Developer' `
            -Text 'Build Boomi integrations for HR technology platforms.' `
            -Keywords $script:keywords |
            Should -BeTrue
    }

    It 'rejects patient scheduling roles' {
        JobMatching\Test-JobRelevance `
            -Title 'Patient Scheduler' `
            -Text 'Schedule patient appointments.' `
            -Keywords $script:keywords |
            Should -BeFalse
    }

    It 'rejects a tourism/location title as a role title' {
        JobMatching\Test-LooksLikeRoleTitle `
            -Title 'Albany, New York' |
            Should -BeFalse
    }

    It 'rejects a one-token same-name false identity' {
        JobMatching\Test-CompanyIdentity `
            -Company 'Albany Med Health System' `
            -Text 'University at Albany jobs and careers' |
            Should -BeFalse
    }

    It 'accepts the exact health-system identity for a one-token brand' {
        JobMatching\Test-CompanyIdentity `
            -Company 'Albany Med Health System' `
            -Text 'Albany Med Health System careers and employment' |
            Should -BeTrue
    }

    It 'requires two distinctive tokens for multi-token fallback identity' {
        JobMatching\Test-CompanyIdentity `
            -Company 'NYU Langone Health' `
            -Text 'NYU Langone careers and technology jobs' |
            Should -BeTrue
    }

    It 'accepts an official employer-domain job URL' {
        $employer = [pscustomobject]@{
            name               = 'Example Health'
            careersUrl         = 'https://careers.examplehealth.org/jobs'
            verificationSource = 'https://www.examplehealth.org/'
        }

        JobMatching\Test-OfficialJobUrl `
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

        JobMatching\Test-OfficialJobUrl `
            -Url 'https://www.indeed.com/viewjob?jk=123' `
            -Employer $employer `
            -Settings $script:settings `
            -EvidenceText 'Example Health UKG Analyst' |
            Should -BeFalse
    }
}
