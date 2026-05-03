param(
    [string]$Source = "S:\#Depts\380\6SIGMABRANDLIST\BRANDING\Projects Folder",
    [string]$Output = (Join-Path (Get-Location) "Brand Lists"),
    [string]$Pd = "",
    [int]$FromYear = (Get-Date).Year,
    [switch]$LatestOnly,
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"
$FromDate = Get-Date -Year $FromYear -Month 1 -Day 1

function Normalize-PdNumber {
    param([string]$Value)

    return $Value.Trim().ToUpperInvariant()
}

function Get-ProjectCandidate {
    param([System.IO.DirectoryInfo]$Directory)

    if ($Directory.Name -notlike "*_*") {
        return $null
    }

    $pdNumber = Normalize-PdNumber (($Directory.Name -split "_", 2)[0])

    if ([string]::IsNullOrWhiteSpace($pdNumber)) {
        return $null
    }

    return [pscustomobject]@{
        ProjectFolderName = $Directory.Name
        ProjectFolderPath = $Directory.FullName
        PdNumber          = $pdNumber
    }
}

function Is-ExcelFile {
    param([string]$FileName)

    $extension = [IO.Path]::GetExtension($FileName).ToLowerInvariant()
    return @(".xlsx", ".xls", ".xlsm", ".xlsb") -contains $extension
}

if (-not (Test-Path -LiteralPath $Source -PathType Container)) {
    throw "Source root does not exist or is not a directory: $Source"
}

$pdFilter = $null

if (-not [string]::IsNullOrWhiteSpace($Pd)) {
    $pdFilter = Normalize-PdNumber $Pd
}

Write-Host "Source: $Source"
Write-Host "Output: $Output"
Write-Host "Modified on/after: Jan 1, $FromYear"
Write-Host "Mode: $(if ($DryRun) { 'dry run' } else { 'copy' })"
Write-Host "Copy behavior: $(if ($LatestOnly) { 'latest only' } else { 'all matching Excel files' })"

if ($pdFilter) {
    Write-Host "PD filter: $pdFilter"
}

Write-Host ""

$projects = Get-ChildItem -LiteralPath $Source -Directory |
    ForEach-Object { Get-ProjectCandidate -Directory $_ } |
    Where-Object { $null -ne $_ } |
    Sort-Object ProjectFolderName

$scanned = 0
$projectsCopied = 0
$filesCopied = 0
$skippedFilter = 0
$skippedNoExcel = 0

foreach ($project in $projects) {
    $scanned++

    if ($pdFilter -and $project.PdNumber -ne $pdFilter) {
        $skippedFilter++
        continue
    }

    $matchingFiles = Get-ChildItem -LiteralPath $project.ProjectFolderPath -File |
        Where-Object {
            (Is-ExcelFile -FileName $_.Name) -and
            (-not $_.Name.StartsWith("~$")) -and
            $_.LastWriteTime -ge $FromDate
        } |
        Sort-Object LastWriteTime -Descending

    if ($LatestOnly) {
        $matchingFiles = $matchingFiles | Select-Object -First 1
    }

    if (-not $matchingFiles) {
        $skippedNoExcel++
        continue
    }

    $destinationDirectory = Join-Path $Output $project.PdNumber

    if (-not $DryRun) {
        New-Item -ItemType Directory -Path $destinationDirectory -Force | Out-Null
    }

    $projectsCopied++

    Write-Host "$($project.ProjectFolderName)"

    foreach ($file in $matchingFiles) {
        $destinationPath = Join-Path $destinationDirectory $file.Name

        if (-not $DryRun) {
            Copy-Item -LiteralPath $file.FullName -Destination $destinationPath -Force
        }

        $filesCopied++

        Write-Host "  $($file.Name)"
        Write-Host "    Edited: $($file.LastWriteTime)"
        Write-Host "    From:   $($file.FullName)"
        Write-Host "    To:     $destinationPath"
    }
}

if ($projectsCopied -eq 0) {
    Write-Host "No matching project folders with Excel files were copied."
}

Write-Host ""
Write-Host "Scanned projects: $scanned"
Write-Host "Projects copied: $projectsCopied"
Write-Host "Files copied: $filesCopied"
Write-Host "Skipped by filter: $skippedFilter"
Write-Host "Skipped with no matching Excel files: $skippedNoExcel"
