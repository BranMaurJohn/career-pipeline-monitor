BeforeAll {
    $root = Resolve-Path (Join-Path $PSScriptRoot '..')

    Import-Module (Join-Path $root 'modules\ClickUp.psm1') `
        -Force `
        -Global

    Import-Module (Join-Path $root 'modules\JobMatching.psm1') `
        -Force `
        -Global

    Import-Module (Join-Path $root 'modules\Utilities.psm1') `
        -Force `
        -Global
}

Describe 'ClickUp safety helpers' {

    It 'builds the expected role task name' {
        JobMatching\Get-ClickUpRoleTaskName `
            -Company 'NYU Langone Health' `
            -RoleTitle 'Sr Developer - UKG Pro WFM Solutions' |
            Should -Be 'NYU Langone Health - Sr Developer - UKG Pro WFM Solutions'
    }

    It 'finds an exact normalized task name' {
        $tasks = @(
            [pscustomobject]@{
                id     = '1'
                name   = 'Adena Health'
                status = [pscustomobject]@{
                    status = 'target company'
                }
            },
            [pscustomobject]@{
                id     = '2'
                name   = 'Other Health'
                status = [pscustomobject]@{
                    status = 'applied'
                }
            }
        )

        $result = ClickUp\Find-ClickUpTaskByExactName `
            -Tasks $tasks `
            -Name 'adena health'

        $result.id | Should -Be '1'
    }

    It 'only promotes unprogressed role tasks' {
        JobMatching\Test-CanPromoteRoleStatus `
            -Status 'target company' |
            Should -BeTrue

        JobMatching\Test-CanPromoteRoleStatus `
            -Status 'applied' |
            Should -BeFalse

        JobMatching\Test-CanPromoteRoleStatus `
            -Status 'technical interview' |
            Should -BeFalse
    }

    It 'documents the no-delete policy in role descriptions' {
        $job = [pscustomobject]@{
            employer   = 'Example Health'
            roleTitle  = 'UKG Analyst'
            workModel  = 'Remote'
            salary     = 'Not disclosed'
            careersUrl = 'https://example.org/careers'
            jobUrl     = 'https://example.org/careers/123'
            source     = 'Official career page'
            verifiedAt = '2026-09-27T00:00:00Z'
        }

        ClickUp\New-RoleDescription -Job $job |
            Should -Match 'never automatically deleted or demoted'
    }
}
