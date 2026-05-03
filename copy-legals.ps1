param(
    [string]$Source = "S:\Legal Drawings\Drawings",
    [string]$Output = (Join-Path (Get-Location) "Share\Legal Drawings"),
    [string]$BrandListSource = "S:\#Depts\380\6SIGMABRANDLIST\BRANDING\Projects Folder",
    [switch]$CopyBrandLists,
    [int]$FromYear = 2026,
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"
$FromDate = Get-Date -Year $FromYear -Month 1 -Day 1

function Get-ProjectNumber {
    param([string]$ProjectFolderName)

    if ($ProjectFolderName -like "*_*") {
        return ($ProjectFolderName -split "_", 2)[0].Trim().ToUpperInvariant()
    }

    return $ProjectFolderName.Trim().ToUpperInvariant()
}

function Get-MatchKind {
    param([string]$FileName)

    $extension = [IO.Path]::GetExtension($FileName).ToLowerInvariant()
    $baseName = [IO.Path]::GetFileNameWithoutExtension($FileName).ToLowerInvariant()
    $normalized = $baseName -replace "[^a-z0-9]+", ""
    $isSpreadsheet = @(".xlsx", ".xlsm", ".xls", ".xlsb", ".xslx") -contains $extension

    if ($isSpreadsheet -and $normalized.Contains("ucpwlcompare")) {
        return "UCP WL Compare spreadsheet"
    }

    if ($isSpreadsheet -and ($normalized.Contains("ucpwirelist") -or $normalized.Contains("ucpwiringlist"))) {
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

function Is-ExcelFile {
    param([string]$FileName)

    $extension = [IO.Path]::GetExtension($FileName).ToLowerInvariant()
    return @(".xlsx", ".xls", ".xlsm", ".xlsb") -contains $extension
}

function Copy-BrandListForProject {
    param(
        [string]$PdNumber,
        [string]$DestinationDirectory
    )

    if (-not (Test-Path -LiteralPath $BrandListSource -PathType Container)) {
        Write-Host "  Brandlist source not found: $BrandListSource"
        return
    }

    $brandProjectFolder = Get-ChildItem -LiteralPath $BrandListSource -Directory |
        Where-Object {
            $parsedPd = Get-ProjectNumber -ProjectFolderName $_.Name
            $parsedPd -eq $PdNumber
        } |
        Sort-Object Name |
        Select-Object -First 1

    if (-not $brandProjectFolder) {
        Write-Host "  Brandlist: No matching Brandlist project folder found for $PdNumber"
        return
    }

    $latestBrandList = Get-ChildItem -LiteralPath $brandProjectFolder.FullName -File |
        Where-Object {
            (Is-ExcelFile -FileName $_.Name) -and
            (-not $_.Name.StartsWith("~$")) -and
            $_.LastWriteTime -ge $FromDate
        } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1

    if (-not $latestBrandList) {
        Write-Host "  Brandlist: No matching Excel file found for $PdNumber"
        return
    }

    $brandListDestination = Join-Path $DestinationDirectory $latestBrandList.Name

    if (-not $DryRun) {
        New-Item -ItemType Directory -Path $DestinationDirectory -Force | Out-Null
        Copy-Item -LiteralPath $latestBrandList.FullName -Destination $brandListDestination -Force
    }

    Write-Host "  Brandlist Excel: $($latestBrandList.Name)"
    Write-Host "    Last edited: $($latestBrandList.LastWriteTime)"
    Write-Host "    From: $($latestBrandList.FullName)"
    Write-Host "    To: $brandListDestination"
}

if (-not (Test-Path -LiteralPath $Source -PathType Container)) {
    throw "Source root does not exist or is not a directory: $Source"
}

Write-Host "Source: $Source"
Write-Host "Output: $Output"
Write-Host "Modified on/after: $($FromDate.ToShortDateString())"
Write-Host "Mode: $(if ($DryRun) { 'dry run' } else { 'copy' })"
Write-Host "Brandlists: $(if ($CopyBrandLists) { 'enabled' } else { 'disabled' })"
Write-Host ""

$projectDirectories = Get-ChildItem -LiteralPath $Source -Directory | Sort-Object Name

foreach ($projectDirectory in $projectDirectories) {
    $electricalPath = Join-Path $projectDirectory.FullName "Electrical"

    if (-not (Test-Path -LiteralPath $electricalPath -PathType Container)) {
        continue
    }

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

        if (-not $latestByKind.ContainsKey($kind) -or $_.LastWriteTime -gt $latestByKind[$kind].LastWriteTime) {
            $latestByKind[$kind] = $_
        }
    }

    Write-Host "$($projectDirectory.Name) -> $destinationDirectory"

    if ($latestByKind.Count -gt 0 -and -not $DryRun) {
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
        Write-Host "    To: $destinationPath"
    }

    if ($CopyBrandLists) {
        Copy-BrandListForProject -PdNumber $projectNumber -DestinationDirectory $destinationDirectory
    }

    Write-Host ""
}
