---
description: Show the installed office-tools binary version and the latest published release version.
---

Run the following in a shell and report the result to the user:

```powershell
$exe = Join-Path $env:LOCALAPPDATA 'Temp\office-tools\office-tools.exe'
if (Test-Path $exe) {
    $current = ((& $exe --version 2>&1) -replace 'office-tools\s+', '').Trim()
} else {
    $current = "(not installed at $exe)"
}

try {
    $rel = Invoke-RestMethod 'https://api.github.com/repos/gunba/office-tools/releases/latest'
    $latest = $rel.tag_name
    $published = $rel.published_at
} catch {
    $latest = "(unreachable: $($_.Exception.Message))"
    $published = ""
}

Write-Host "Installed: $current"
Write-Host "Latest:    $latest  $(if ($published) { "($published)" })"
```
