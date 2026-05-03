param(
    [string]$Source = "S:\Legal Drawings\Drawings",
    [string]$Output = (Join-Path (Get-Location) "Share\Legal Drawings"),
    [int]$FromYear = 2026,
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"
$FromDate = Get-Date -Year $FromYear -Month 1 -Day 1 -Hour 0 -Minute 0 -Second 0

function Get-ProjectNumber {
    param([string]$ProjectFolderName)

    if ($ProjectFolderName -like "*_*") {
        return ($ProjectFolderName -split "_", 2)[0].Trim()
    }

    return $ProjectFolderName.Trim()
}

function Get-MatchKind {
    param([string]$FileName)

    $extension = [System.IO.Path]::GetExtension($FileName).ToLowerInvariant()
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FileName).ToLowerInvariant()
    $normalizedBaseName = ($baseName -replace "[^a-z0-9]+", "")
    $isSpreadsheet = @(".xlsx", ".xlsm", ".xls", ".xlsb", ".xslx") -contains $extension

    if ($isSpreadsheet -and $normalizedBaseName.Contains("ucpwlcompare")) {
        return "UCP WL Compare spreadsheet"
    }

    if ($isSpreadsheet -and (
        $normalizedBaseName.Contains("ucpwirelist") -or
        $normalizedBaseName.Contains("ucpwiringlist")
    )) {
        return "UCP wire list spreadsheet"
    }

    if ($isSpreadsheet -and $baseName.Contains("ucp")) {
        return "UCP spreadsheet"
    }

    if ($extension -eq ".pdf" -and $baseName.Contains("lay")) {
        return "LAY PDF"
    }

    return $null
}

if (-not (Test-Path -LiteralPath $Source -PathType Container)) {
    throw "Source root does not exist or is not a directory: $Source"
}

Write-Host "Source: $Source"
Write-Host "Output: $Output"
Write-Host "Modified on/after: $($FromDate.ToShortDateString())"
Write-Host "Mode: $(if ($DryRun) { 'dry run' } else { 'copy' })"
Write-Host ""

$projectDirectories = Get-ChildItem -LiteralPath $Source -Directory |
    Sort-Object Name

if (-not $projectDirectories) {
    Write-Host "No project folders were found."
    exit 0
}

$foundElectricalProject = $false

foreach ($projectDirectory in $projectDirectories) {
    $electricalPath = Join-Path $projectDirectory.FullName "Electrical"

    if (-not (Test-Path -LiteralPath $electricalPath -PathType Container)) {
        continue
    }

    $foundElectricalProject = $true

    $projectNumber = Get-ProjectNumber -ProjectFolderName $projectDirectory.Name
    $destinationDirectory = Join-Path $Output $projectNumber

    $latestByKind = @{}

    Get-ChildItem -LiteralPath $electricalPath -File -Recurse | ForEach-Object {
        $kind = Get-MatchKind -FileName $_.Name

        if (-not $kind) {
            return
        }

        if ($_.LastWriteTime -lt $FromDate) {
            return
        }

        if (
            -not $latestByKind.ContainsKey($kind) -or
            $_.LastWriteTime -gt $latestByKind[$kind].LastWriteTime
        ) {
            $latestByKind[$kind] = $_
        }
    }

    Write-Host "$($projectDirectory.Name) -> $destinationDirectory"

    if ($latestByKind.Count -eq 0) {
        Write-Host "  No matching UCP spreadsheet, UCP WL Compare spreadsheet, UCP wire list spreadsheet, or LAY PDF found."
        continue
    }

    if (-not $DryRun) {
        New-Item -ItemType Directory -Path $destinationDirectory -Force | Out-Null
    }

    foreach ($kind in @(
        "LAY PDF",
        "UCP spreadsheet",
        "UCP wire list spreadsheet",
        "UCP WL Compare spreadsheet"
    )) {
        if (-not $latestByKind.ContainsKey($kind)) {
            continue
        }

        $file = $latestByKind[$kind]
        $destinationPath = Join-Path $destinationDirectory $file.Name

        if (-not $DryRun) {
            Copy-Item -LiteralPath $file.FullName -Destination $destinationPath -Force
        }

        Write-Host "  ${kind}: $($file.Name)"
        Write-Host "    Last edited: $($file.LastWriteTime)"
        Write-Host "    From: $($file.FullName)"
    }
}

if (-not $foundElectricalProject) {
    Write-Host "No project folders with an Electrical directory were found."
}
