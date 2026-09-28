[CmdletBinding()]
param(
    [string]$WorkbookPath = (Join-Path $PSScriptRoot '..\input\us_health_systems_2026_current_map.xlsx'),
    [string]$OutputPath = (Join-Path $PSScriptRoot '..\config\employers.json')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot '..\modules\Utilities.psm1') -Force

function Get-ExcelColumnIndex {
    param([Parameter(Mandatory)][string]$Letters)
    $index = 0
    foreach ($char in $Letters.ToUpperInvariant().ToCharArray()) {
        $index = ($index * 26) + ([int][char]$char - [int][char]'A' + 1)
    }
    return $index
}

function Get-ZipEntryText {
    param(
        [Parameter(Mandatory)]$Archive,
        [Parameter(Mandatory)][string]$EntryName
    )
    $entry = $Archive.GetEntry($EntryName)
    if (-not $entry) { return $null }
    $stream = $entry.Open()
    try {
        $reader = New-Object System.IO.StreamReader($stream)
        try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
    } finally {
        $stream.Dispose()
    }
}

function Get-XlsxWorksheetRows {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$WorksheetName,
        [int]$HeaderRow = 4
    )

    Add-Type -AssemblyName System.IO.Compression
    try { Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue } catch {}

    $file = [System.IO.File]::OpenRead((Resolve-Path -LiteralPath $Path))
    try {
        $archive = New-Object System.IO.Compression.ZipArchive($file, [System.IO.Compression.ZipArchiveMode]::Read, $false)
        try {
            [xml]$workbookXml = Get-ZipEntryText -Archive $archive -EntryName 'xl/workbook.xml'
            [xml]$relsXml = Get-ZipEntryText -Archive $archive -EntryName 'xl/_rels/workbook.xml.rels'

            $workbookNs = New-Object System.Xml.XmlNamespaceManager($workbookXml.NameTable)
            $workbookNs.AddNamespace('d', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
            $workbookNs.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')

            $sheetNode = $workbookXml.SelectSingleNode("//d:sheets/d:sheet[@name='$WorksheetName']", $workbookNs)
            if (-not $sheetNode) { throw "Worksheet '$WorksheetName' was not found in $Path." }

            $relationshipId = $sheetNode.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
            $relsNs = New-Object System.Xml.XmlNamespaceManager($relsXml.NameTable)
            $relsNs.AddNamespace('p', 'http://schemas.openxmlformats.org/package/2006/relationships')
            $relationship = $relsXml.SelectSingleNode("//p:Relationship[@Id='$relationshipId']", $relsNs)
            if (-not $relationship) { throw "Worksheet relationship '$relationshipId' was not found." }

            $target = [string]$relationship.Target
            if ($target.StartsWith('/')) { $target = $target.TrimStart('/') }
            elseif (-not $target.StartsWith('xl/')) { $target = 'xl/' + $target.TrimStart('.').TrimStart('/') }

            $sharedStrings = @()
            $sharedRaw = Get-ZipEntryText -Archive $archive -EntryName 'xl/sharedStrings.xml'
            if ($sharedRaw) {
                [xml]$sharedXml = $sharedRaw
                $sharedNs = New-Object System.Xml.XmlNamespaceManager($sharedXml.NameTable)
                $sharedNs.AddNamespace('d', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
                foreach ($item in $sharedXml.SelectNodes('//d:si', $sharedNs)) {
                    $sharedStrings += [string]$item.InnerText
                }
            }

            [xml]$sheetXml = Get-ZipEntryText -Archive $archive -EntryName $target
            $sheetNs = New-Object System.Xml.XmlNamespaceManager($sheetXml.NameTable)
            $sheetNs.AddNamespace('d', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
            $rows = @($sheetXml.SelectNodes('//d:sheetData/d:row', $sheetNs))

            $headerMap = @{}
            $output = @()

            foreach ($row in $rows) {
                $rowNumber = [int]$row.r
                $valuesByColumn = @{}

                foreach ($cell in @($row.SelectNodes('d:c', $sheetNs))) {
                    $reference = [string]$cell.r
                    $letters = ([regex]::Match($reference, '^[A-Z]+')).Value
                    if (-not $letters) { continue }
                    $columnIndex = Get-ExcelColumnIndex -Letters $letters
                    $cellType = [string]$cell.t
                    $value = ''

                    if ($cellType -eq 'inlineStr') {
                        $inline = $cell.SelectSingleNode('d:is', $sheetNs)
                        if ($inline) { $value = [string]$inline.InnerText }
                    } else {
                        $valueNode = $cell.SelectSingleNode('d:v', $sheetNs)
                        if ($valueNode) { $value = [string]$valueNode.InnerText }
                        if ($cellType -eq 's' -and $value -match '^\d+$') {
                            $sharedIndex = [int]$value
                            if ($sharedIndex -ge 0 -and $sharedIndex -lt $sharedStrings.Count) {
                                $value = [string]$sharedStrings[$sharedIndex]
                            }
                        }
                    }
                    $valuesByColumn[$columnIndex] = $value
                }

                if ($rowNumber -eq $HeaderRow) {
                    foreach ($entry in $valuesByColumn.GetEnumerator()) {
                        if (-not [string]::IsNullOrWhiteSpace([string]$entry.Value)) {
                            $headerMap[[int]$entry.Key] = [string]$entry.Value
                        }
                    }
                    continue
                }

                if ($rowNumber -le $HeaderRow) { continue }
                $record = [ordered]@{}
                foreach ($columnIndex in ($headerMap.Keys | Sort-Object)) {
                    $header = [string]$headerMap[$columnIndex]
                    $record[$header] = if ($valuesByColumn.ContainsKey($columnIndex)) { [string]$valuesByColumn[$columnIndex] } else { '' }
                }
                $output += [pscustomobject]$record
            }

            return @($output)
        } finally {
            $archive.Dispose()
        }
    } finally {
        $file.Dispose()
    }
}

if (-not (Test-Path -LiteralPath $WorkbookPath)) { throw "Workbook not found: $WorkbookPath" }
Write-PipelineLog "Reading employer universe from $WorkbookPath"
$rows = @(Get-XlsxWorksheetRows -Path $WorkbookPath -WorksheetName '2026 Systems' -HeaderRow 4)

$existingRegistry = Get-JsonFile -Path $OutputPath -Default $null
$existingByName = @{}
if ($existingRegistry -and $existingRegistry.employers) {
    foreach ($employer in @($existingRegistry.employers)) {
        $key = ConvertTo-NormalizedText ([string]$employer.name)
        if ($key) { $existingByName[$key] = $employer }
    }
}

$grouped = $rows | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.'2026 Job-Search Employer Name') } | Group-Object '2026 Job-Search Employer Name'
$employers = @()
$id = 0
foreach ($group in $grouped | Sort-Object Name) {
    $id++
    $first = $group.Group | Select-Object -First 1
    $name = [string]$group.Name
    $key = ConvertTo-NormalizedText $name
    $existing = if ($existingByName.ContainsKey($key)) { $existingByName[$key] } else { $null }

    $verificationSource = @($group.Group | ForEach-Object { [string]$_.'2026 Verification Source' } | Where-Object { $_ } | Select-Object -First 1)
    $recordNumbers = @($group.Group | ForEach-Object { [int]$_.'Record #' })

    $employers += [pscustomobject][ordered]@{
        id                    = $id
        name                  = $name
        sourceRecordNumbers   = $recordNumbers
        currentCompany        = [string]$first.'Current 2026 Company / System'
        parentEnterprise      = [string]$first.'2026 Parent / Enterprise'
        operatingBrandRegion  = [string]$first.'2026 Operating Brand / Region'
        structureStatus       = [string]$first.'2026 Structure Status'
        verificationSource    = if ($verificationSource.Count -gt 0) { [string]$verificationSource[0] } else { '' }
        verifiedAsOf          = [string]$first.'Verified As Of'
        structureCategory     = [string]$first.'2026 Structure Category'
        careersUrl            = if ($existing) { [string]$existing.careersUrl } else { '' }
        ats                   = if ($existing) { [string]$existing.ats } else { '' }
        priority              = if ($existing -and $existing.priority) { [int]$existing.priority } else { 2 }
        active                = if ($existing -and $null -ne $existing.active) { [bool]$existing.active } else { $true }
        lastDiscoveredAt      = if ($existing) { [string]$existing.lastDiscoveredAt } else { '' }
    }
}

$registry = [pscustomobject][ordered]@{
    generatedAt    = (Get-Date).ToUniversalTime().ToString('o')
    sourceWorkbook = [System.IO.Path]::GetFileName($WorkbookPath)
    recordCount    = $rows.Count
    employerCount  = $employers.Count
    employers      = $employers
}

Set-JsonFile -Path $OutputPath -Value $registry -Depth 12
Write-PipelineLog ("Imported {0} workbook records into {1} unique employers." -f $rows.Count, $employers.Count) -Level Success
$registry
