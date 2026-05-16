param(
    [string]$OfficeTools = "",
    [switch]$KeepTemp
)

$ErrorActionPreference = "Stop"

function Resolve-OfficeToolsBinary {
    param([string]$Requested)
    if ($Requested -and (Test-Path -LiteralPath $Requested)) {
        return (Resolve-Path -LiteralPath $Requested).Path
    }

    $repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
    $candidates = @(
        (Join-Path $repoRoot "target\release\office-tools.exe"),
        (Join-Path $repoRoot "plugins\office-tools\bin\office-tools.exe")
    )
    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            return (Resolve-Path -LiteralPath $candidate).Path
        }
    }
    throw "office-tools.exe not found. Run cargo build --release or pass -OfficeTools <path>."
}

function Invoke-OfficeTools {
    param(
        [string]$Exe,
        [string[]]$ToolArgs
    )
    $output = & $Exe @ToolArgs 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "office-tools $($ToolArgs -join ' ') failed with exit $LASTEXITCODE`n$output"
    }
    return ($output -join "`n")
}

function New-SmokeWorkbook {
    param(
        [string]$Path,
        [string]$SheetName,
        [string]$Value
    )
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    try {
        $workbook = $excel.Workbooks.Add()
        $sheet = $workbook.Worksheets.Item(1)
        $sheet.Name = $SheetName
        $sheet.Range("A1").Value2 = $Value
        $sheet.Range("B1").Formula = "=1+1"
        $workbook.SaveAs($Path, 51)
        $workbook.Close($false)
    } finally {
        $excel.Quit()
    }
}

function New-SmokeDocx {
    param([string]$Path)
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    try {
        $document = $word.Documents.Add()
        $document.Content.Text = "old text"
        $document.SaveAs2($Path, 12)
        $document.Close($false)
    } finally {
        $word.Quit()
    }
}

$exe = Resolve-OfficeToolsBinary $OfficeTools
$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ("office-tools-wincom-" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Force -Path $tempRoot | Out-Null

try {
    Write-Host "office-tools: $exe"
    Write-Host "temp:         $tempRoot"

    Write-Host "Checking COM availability..."
    $doctor = Invoke-OfficeTools $exe @("doctor") | ConvertFrom-Json
    $failed = @($doctor.wincom | Where-Object { -not $_.ok })
    if ($failed.Count -gt 0) {
        throw "COM doctor failed: $($failed | ConvertTo-Json -Compress)"
    }

    $targetXlsx = Join-Path $tempRoot "target.xlsx"
    $sourceXlsx = Join-Path $tempRoot "source.xlsx"
    New-SmokeWorkbook $targetXlsx "Sheet1" "target"
    New-SmokeWorkbook $sourceXlsx "Source" "source"

    Write-Host "Exercising Excel COM validate..."
    Invoke-OfficeTools $exe @("xlsx", "validate", $targetXlsx, "--full-calc", "--save", "--check-errors") | Out-Null

    Write-Host "Exercising Excel COM insert..."
    Invoke-OfficeTools $exe @("xlsx", "insert", $targetXlsx, "--sheet", "Sheet1", "--axis", "rows", "--range", "2:2") | Out-Null

    Write-Host "Exercising Excel COM rename-sheet..."
    Invoke-OfficeTools $exe @("xlsx", "rename-sheet", $targetXlsx, "--sheet", "Sheet1", "--name", "Renamed") | Out-Null
    $renamedSheets = Invoke-OfficeTools $exe @("xlsx", "list-sheets", $targetXlsx, "--json") | ConvertFrom-Json
    if (-not ($renamedSheets | Where-Object { $_.name -eq "Renamed" })) {
        throw "Renamed sheet was not found after xlsx rename-sheet."
    }

    Write-Host "Exercising Excel COM autofit..."
    Invoke-OfficeTools $exe @("xlsx", "autofit", $targetXlsx, "--sheet", "Renamed", "--axis", "columns") | Out-Null

    Write-Host "Exercising Excel COM format..."
    Invoke-OfficeTools $exe @("xlsx", "format", $targetXlsx, "--sheet", "Renamed", "--range", "A1:B1", "--bold", "true", "--fill-color", "D9EAF7") | Out-Null

    Write-Host "Exercising Excel COM copy-sheets..."
    Invoke-OfficeTools $exe @("xlsx", "copy-sheets", $targetXlsx, "--source", $sourceXlsx, "--sheets", "Source=Copied", "--replace") | Out-Null
    $sheets = Invoke-OfficeTools $exe @("xlsx", "list-sheets", $targetXlsx, "--json") | ConvertFrom-Json
    if (-not ($sheets | Where-Object { $_.name -eq "Copied" })) {
        throw "Copied sheet was not found after xlsx copy-sheets."
    }

    $docx = Join-Path $tempRoot "word.docx"
    New-SmokeDocx $docx
    Write-Host "Exercising Word COM replace..."
    Invoke-OfficeTools $exe @("docx", "replace", $docx, "--find", "old", "--replace", "new", "--engine", "com") | Out-Null
    $docText = Invoke-OfficeTools $exe @("docx", "read", $docx)
    if ($docText -notmatch "new text") {
        throw "Word COM replacement did not produce expected text."
    }

    Write-Host "Exercising read-only Outlook query..."
    $outlook = Invoke-OfficeTools $exe @("outlook", "--hours", "24", "--count", "1", "--include-read", "--folder", "inbox") | ConvertFrom-Json
    if ($null -eq $outlook.emails -or $null -eq $outlook.meeting_requests) {
        throw "Outlook output did not include expected arrays."
    }

    Write-Host "WinCOM smoke test passed."
} finally {
    if ($KeepTemp) {
        Write-Host "Keeping temp directory: $tempRoot"
    } else {
        Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
