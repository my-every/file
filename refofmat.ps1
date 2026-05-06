$keepPd = (
    (Get-Content '.\Share\Legal Drawings\prebuild-report.json' -Raw | ConvertFrom-Json).details.dueThisMonth |
    ForEach-Object { $_.pdNumber.ToUpper() } |
    Sort-Object -Unique
)

Get-ChildItem '.\Share\Projects' -Directory |
    Where-Object {
        $n = $_.Name.ToUpper()
        -not (
            $keepPd | Where-Object {
                $n -eq $_ -or $n.StartsWith("$_" + "_")
            }
        )
    } |
    Remove-Item -Recurse -Force -WhatIf
