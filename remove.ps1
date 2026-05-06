Get-ChildItem .\Share\users -Recurse -Force |
  Where-Object {
    ($_.PSIsContainer -and $_.Name -eq 'activity-index') -or
    (-not $_.PSIsContainer -and @('activity.json','activity-events.jsonl') -contains $_.Name)
  } |
  Remove-Item -Recurse -Force
