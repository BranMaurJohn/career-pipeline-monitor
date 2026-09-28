[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Title,
    [string]$Text = '',
    [string]$KeywordsPath = (Join-Path $PSScriptRoot '..\config\keywords.json')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot '..\modules\Utilities.psm1') -Force
Import-Module (Join-Path $PSScriptRoot '..\modules\JobMatching.psm1') -Force

$keywords = Get-JsonFile -Path $KeywordsPath -Default $null
if (-not $keywords) { throw "Keywords file is empty: $KeywordsPath" }

$result = Test-JobRelevance -Title $Title -Text $Text -Keywords $keywords
[pscustomobject]@{
    Title = $Title
    Match = [bool]$result
}
