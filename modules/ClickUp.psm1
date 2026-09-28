Set-StrictMode -Version Latest

Import-Module (Join-Path $PSScriptRoot 'Utilities.psm1') -Force
Import-Module (Join-Path $PSScriptRoot 'JobMatching.psm1') -Force

function Invoke-ClickUpRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateSet('GET','POST','PUT')][string]$Method,
        [Parameter(Mandatory)][string]$Path,
        [AllowNull()]$Body,
        [Parameter(Mandatory)][string]$Token,
        [int]$MaxAttempts = 5
    )

    $uri = 'https://api.clickup.com/api/v2' + $Path
    $headers = @{ Authorization = $Token; 'Content-Type' = 'application/json' }

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        $params = @{
            Method      = $Method
            Uri         = $uri
            Headers     = $headers
            ErrorAction = 'Stop'
        }
        if ($null -ne $Body) { $params['Body'] = ($Body | ConvertTo-Json -Depth 10) }

        try {
            return Invoke-RestMethod @params
        } catch {
            $statusCode = 0
            try { $statusCode = [int]$_.Exception.Response.StatusCode } catch {}

            $retryable = ($statusCode -eq 429 -or $statusCode -ge 500 -or $statusCode -eq 0)
            if ($retryable -and $attempt -lt $MaxAttempts) {
                $sleepSeconds = [Math]::Min(30, [Math]::Pow(2, $attempt))
                try {
                    $retryAfter = $_.Exception.Response.Headers['Retry-After']
                    if ($retryAfter) { $sleepSeconds = [Math]::Max($sleepSeconds, [int]$retryAfter) }
                } catch {}
                Start-Sleep -Seconds $sleepSeconds
                continue
            }

            $message = $_.Exception.Message
            if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $message += ' | ' + $_.ErrorDetails.Message }
            throw "ClickUp API $Method $Path failed after $attempt attempt(s): $message"
        }
    }
}

function Get-ClickUpTasks {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ListId,
        [Parameter(Mandatory)][string]$Token
    )

    $all = @()
    for ($page = 0; $page -lt 100; $page++) {
        $path = "/list/$ListId/task?archived=false&include_closed=true&subtasks=true&page=$page"
        $response = Invoke-ClickUpRequest -Method GET -Path $path -Token $Token
        $tasks = @($response.tasks)
        if ($tasks.Count -eq 0) { break }
        $all += $tasks
        if ($tasks.Count -lt 100) { break }
    }
    return @($all)
}

function Get-TaskStatusName {
    [CmdletBinding()]
    param($Task)
    if ($null -eq $Task -or $null -eq $Task.status) { return '' }
    if ($Task.status.PSObject.Properties['status']) { return [string]$Task.status.status }
    return [string]$Task.status
}

function Find-ClickUpTaskByExactName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Tasks,
        [Parameter(Mandatory)][string]$Name
    )

    $target = ConvertTo-NormalizedText $Name
    return @($Tasks | Where-Object { (ConvertTo-NormalizedText ([string]$_.name)) -eq $target }) | Select-Object -First 1
}

function New-CompanyMonitorDescription {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Employer)

    return @(
        'Career Pipeline Monitor company record',
        '',
        ('Company: {0}' -f [string]$Employer.name),
        ('Careers URL: {0}' -f [string]$Employer.careersUrl),
        ('ATS: {0}' -f [string]$Employer.ats),
        ('Verification source: {0}' -f [string]$Employer.verificationSource),
        ('Workbook records: {0}' -f (@($Employer.sourceRecordNumbers) -join ', ')),
        '',
        'Automation rule: this task is permanent and is never automatically deleted when a job closes.'
    ) -join [Environment]::NewLine
}

function New-RoleDescription {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Job)

    return @(
        'Career Pipeline Monitor opportunity',
        '',
        ('Company: {0}' -f [string]$Job.employer),
        ('Role: {0}' -f [string]$Job.roleTitle),
        ('Work model: {0}' -f [string]$Job.workModel),
        ('Salary: {0}' -f [string]$Job.salary),
        ('Official careers URL: {0}' -f [string]$Job.careersUrl),
        ('Job URL: {0}' -f [string]$Job.jobUrl),
        ('Source: {0}' -f [string]$Job.source),
        ('Verified: {0}' -f [string]$Job.verifiedAt),
        '',
        'Automation rule: this opportunity is never automatically deleted or demoted if the posting disappears.'
    ) -join [Environment]::NewLine
}

function New-ClickUpTask {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ListId,
        [Parameter(Mandatory)][string]$Token,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Status,
        [Parameter(Mandatory)][string]$Description
    )

    $body = @{ name = $Name; status = $Status; description = $Description }
    return Invoke-ClickUpRequest -Method POST -Path "/list/$ListId/task" -Body $body -Token $Token
}

function Set-ClickUpTaskStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$TaskId,
        [Parameter(Mandatory)][string]$Token,
        [Parameter(Mandatory)][string]$Status
    )

    return Invoke-ClickUpRequest -Method PUT -Path "/task/$TaskId" -Body @{ status = $Status } -Token $Token
}

function Sync-CareerPipelineToClickUp {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Employers,
        [Parameter(Mandatory)]$Jobs,
        [Parameter(Mandatory)][string]$ListId,
        [Parameter(Mandatory)][string]$Token,
        [string]$TargetCompanyStatus = 'target company',
        [string]$RoleIdentifiedStatus = 'role identified',
        [int]$WriteDelayMilliseconds = 750,
        [switch]$DryRun
    )

    $tasks = @(Get-ClickUpTasks -ListId $ListId -Token $Token)
    $actions = @()

    foreach ($employer in @($Employers)) {
        $name = [string]$employer.name
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        $existing = Find-ClickUpTaskByExactName -Tasks $tasks -Name $name
        if (-not $existing) {
            $actions += [pscustomobject]@{ action = 'CREATE COMPANY MONITOR'; taskName = $name; status = $TargetCompanyStatus }
            if (-not $DryRun) {
                $created = New-ClickUpTask -ListId $ListId -Token $Token -Name $name -Status $TargetCompanyStatus -Description (New-CompanyMonitorDescription -Employer $employer)
                $tasks += $created
                Start-Sleep -Milliseconds $WriteDelayMilliseconds
            }
        }
    }

    foreach ($job in @($Jobs)) {
        $taskName = Get-ClickUpRoleTaskName -Company ([string]$job.employer) -RoleTitle ([string]$job.roleTitle)
        $existing = Find-ClickUpTaskByExactName -Tasks $tasks -Name $taskName

        if (-not $existing) {
            $actions += [pscustomobject]@{ action = 'CREATE ROLE'; taskName = $taskName; status = $RoleIdentifiedStatus }
            if (-not $DryRun) {
                $created = New-ClickUpTask -ListId $ListId -Token $Token -Name $taskName -Status $RoleIdentifiedStatus -Description (New-RoleDescription -Job $job)
                $tasks += $created
                Start-Sleep -Milliseconds $WriteDelayMilliseconds
            }
            continue
        }

        $currentStatus = Get-TaskStatusName -Task $existing
        if (Test-CanPromoteRoleStatus -Status $currentStatus) {
            $actions += [pscustomobject]@{ action = 'PROMOTE EXISTING ROLE'; taskName = $taskName; status = $RoleIdentifiedStatus }
            if (-not $DryRun) {
                Set-ClickUpTaskStatus -TaskId ([string]$existing.id) -Token $Token -Status $RoleIdentifiedStatus | Out-Null
                Start-Sleep -Milliseconds $WriteDelayMilliseconds
            }
        } else {
            $actions += [pscustomobject]@{ action = 'PRESERVE EXISTING ROLE'; taskName = $taskName; status = $currentStatus }
        }
    }

    return @($actions)
}

Export-ModuleMember -Function @(
    'Invoke-ClickUpRequest',
    'Get-ClickUpTasks',
    'Get-TaskStatusName',
    'Find-ClickUpTaskByExactName',
    'New-CompanyMonitorDescription',
    'New-RoleDescription',
    'New-ClickUpTask',
    'Set-ClickUpTaskStatus',
    'Sync-CareerPipelineToClickUp'
)
