# Set these to your actual paths
$legalDir  = "C:\Path\To\Legal Drawings"
$brandDir  = "C:\Path\To\Brand List"
$outFile   = "$env:USERPROFILE\Desktop\dir-tree-export.txt"

@"
=== LEGAL DRAWINGS ===
$legalDir
"@ | Out-File $outFile -Encoding utf8

Get-ChildItem -Path $legalDir -Recurse -ErrorAction SilentlyContinue |
  Select-Object FullName, Length, LastWriteTime |
  Format-Table -AutoSize |
  Out-String -Width 300 |
  Add-Content $outFile

@"

=== BRAND LIST ===
$brandDir
"@ | Add-Content $outFile

Get-ChildItem -Path $brandDir -Recurse -ErrorAction SilentlyContinue |
  Select-Object FullName, Length, LastWriteTime |
  Format-Table -AutoSize |
  Out-String -Width 300 |
  Add-Content $outFile

Write-Host "Saved to $outFile"
