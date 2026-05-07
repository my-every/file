$legalDir = "S:\Legal Drawings\Drawings"
$brandDir  = "S:\#Depts\380\6SIGMABRANDLIST\BRANDING\Projects Folder"
$outFile   = "$env:USERPROFILE\Desktop\dir-tree-export.txt"

"=== LEGAL DRAWINGS ===" | Out-File $outFile -Encoding utf8
"Root: $legalDir" | Add-Content $outFile

Get-ChildItem -LiteralPath $legalDir -Directory | Sort-Object Name | ForEach-Object {
    "  [$($_.Name)]" | Add-Content $outFile
    $electricalPath = Join-Path $_.FullName "Electrical"
    if (Test-Path -LiteralPath $electricalPath) {
        Get-ChildItem -LiteralPath $electricalPath -File -Recurse |
            Sort-Object LastWriteTime -Descending |
            ForEach-Object {
                "    $($_.LastWriteTime.ToString('yyyy-MM-dd HH:mm'))  $($_.Name)  ($([math]::Round($_.Length/1KB,1)) KB)" | Add-Content $outFile
                "    $($_.FullName)" | Add-Content $outFile
            }
    } else {
        "    (no Electrical subfolder)" | Add-Content $outFile
    }
}

"`n=== BRAND LIST ===" | Add-Content $outFile
"Root: $brandDir" | Add-Content $outFile

Get-ChildItem -LiteralPath $brandDir -Directory | Sort-Object Name | ForEach-Object {
    "  [$($_.Name)]" | Add-Content $outFile
    Get-ChildItem -LiteralPath $_.FullName -File |
        Where-Object { -not $_.Name.StartsWith("~$") } |
        Sort-Object LastWriteTime -Descending |
        ForEach-Object {
            "    $($_.LastWriteTime.ToString('yyyy-MM-dd HH:mm'))  $($_.Name)  ($([math]::Round($_.Length/1KB,1)) KB)" | Add-Content $outFile
        }
}

Write-Host "Done -> $outFile"
Invoke-Item $outFile
