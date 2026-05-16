use crate::docx::{DocxReplaceArgs, ReplaceResult, Replacement, ReplacementCount};
use crate::xlsx::{
    AutoFitArgs, AutoFitAxis, CopySheetsArgs, FormatArgs, InsertArgs, InsertAxis, RenameSheetArgs,
    ValidateArgs,
};
use anyhow::{Context, Result, bail};
use std::io::Write;
use std::path::Path;
use std::process::Command;

pub fn excel_validate(args: &ValidateArgs) -> Result<String> {
    ensure_windows()?;
    let script = format!(
        r#"
$ErrorActionPreference = 'Stop'
$path = {}
$excel = New-Object -ComObject Excel.Application
$excel.Visible = {keep_open}
$excel.DisplayAlerts = $false
$excel.AskToUpdateLinks = $false
$wb = $excel.Workbooks.Open($path, 0, {read_only})
try {{
  if ({full_calc}) {{
    $excel.CalculateFullRebuild()
  }} else {{
    $excel.Calculate()
  }}
  $result = [ordered]@{{ opened = $path; sheets = $wb.Worksheets.Count; errors = @() }}
  if ({check_errors}) {{
    foreach ($ws in $wb.Worksheets) {{
      $used = $ws.UsedRange
      foreach ($cell in $used.Cells) {{
        $value = [string]$cell.Text
        if ($value -match '^#(DIV/0!|N/A|NAME\?|NULL!|NUM!|REF!|VALUE!|SPILL!|CALC!)') {{
          $result.errors += [ordered]@{{ sheet = $ws.Name; address = $cell.Address($false,$false); value = $value }}
          if ($result.errors.Count -ge {max_errors}) {{ break }}
        }}
      }}
      if ($result.errors.Count -ge {max_errors}) {{ break }}
    }}
  }}
  if ({save}) {{ $wb.Save() }}
  $result.saved = {save}
  $result | ConvertTo-Json -Depth 5
}} finally {{
  if (-not {keep_open}) {{
    $wb.Close($false)
    $excel.Quit()
  }}
}}
"#,
        ps_string(args.file.as_path()),
        keep_open = ps_bool(args.keep_open),
        read_only = ps_bool(!args.save),
        full_calc = ps_bool(args.full_calc),
        check_errors = ps_bool(args.check_errors),
        max_errors = args.max_errors,
        save = ps_bool(args.save),
    );
    run_powershell(&script)
}

pub fn excel_insert(args: &InsertArgs) -> Result<String> {
    ensure_windows()?;
    let insert_call = match args.axis {
        InsertAxis::Rows => "$target.EntireRow.Insert()",
        InsertAxis::Columns => "$target.EntireColumn.Insert()",
    };
    let script = format!(
        r#"
$ErrorActionPreference = 'Stop'
$path = {}
$excel = New-Object -ComObject Excel.Application
$excel.Visible = {keep_open}
$excel.DisplayAlerts = $false
$excel.AskToUpdateLinks = $false
$wb = $excel.Workbooks.Open($path, 0, $false)
try {{
  $ws = $wb.Worksheets.Item({})
  $target = $ws.Range({})
  {insert_call}
  $wb.Save()
  [ordered]@{{ file = $path; sheet = {}; range = {}; axis = {}; saved = $true }} | ConvertTo-Json
}} finally {{
  if (-not {keep_open}) {{
    $wb.Close($false)
    $excel.Quit()
  }}
}}
"#,
        ps_string(args.file.as_path()),
        ps_literal(&args.sheet),
        ps_literal(&args.range),
        ps_literal(&args.sheet),
        ps_literal(&args.range),
        ps_literal(match args.axis {
            InsertAxis::Rows => "rows",
            InsertAxis::Columns => "columns",
        }),
        keep_open = ps_bool(args.keep_open),
    );
    run_powershell(&script)
}

pub fn excel_rename_sheet(args: &RenameSheetArgs) -> Result<String> {
    ensure_windows()?;
    let script = format!(
        r#"
$ErrorActionPreference = 'Stop'
$path = {}
$excel = New-Object -ComObject Excel.Application
$excel.Visible = {keep_open}
$excel.DisplayAlerts = $false
$excel.AskToUpdateLinks = $false
$wb = $excel.Workbooks.Open($path, 0, $false)
try {{
  $ws = $wb.Worksheets.Item({})
  $oldName = $ws.Name
  $ws.Name = {}
  $wb.Save()
  [ordered]@{{ file = $path; old_name = $oldName; new_name = $ws.Name; saved = $true }} | ConvertTo-Json
}} finally {{
  if (-not {keep_open}) {{
    $wb.Close($false)
    $excel.Quit()
  }}
}}
"#,
        ps_string(args.file.as_path()),
        ps_literal(&args.sheet),
        ps_literal(&args.name),
        keep_open = ps_bool(args.keep_open),
    );
    run_powershell(&script)
}

pub fn excel_autofit(args: &AutoFitArgs) -> Result<String> {
    ensure_windows()?;
    let range = args
        .range
        .as_ref()
        .map(|value| ps_literal(value))
        .unwrap_or_else(|| "$null".to_string());
    let autofit_columns = matches!(args.axis, AutoFitAxis::All | AutoFitAxis::Columns);
    let autofit_rows = matches!(args.axis, AutoFitAxis::All | AutoFitAxis::Rows);
    let script = format!(
        r#"
$ErrorActionPreference = 'Stop'
$path = {}
$rangeAddress = {range}
$excel = New-Object -ComObject Excel.Application
$excel.Visible = {keep_open}
$excel.DisplayAlerts = $false
$excel.AskToUpdateLinks = $false
$wb = $excel.Workbooks.Open($path, 0, $false)
try {{
  $ws = $wb.Worksheets.Item({})
  $target = if ($rangeAddress) {{ $ws.Range($rangeAddress) }} else {{ $ws.UsedRange }}
  if ({autofit_columns}) {{ $target.EntireColumn.AutoFit() | Out-Null }}
  if ({autofit_rows}) {{ $target.EntireRow.AutoFit() | Out-Null }}
  $wb.Save()
  [ordered]@{{ file = $path; sheet = $ws.Name; range = $rangeAddress; axis = {}; saved = $true }} | ConvertTo-Json
}} finally {{
  if (-not {keep_open}) {{
    $wb.Close($false)
    $excel.Quit()
  }}
}}
"#,
        ps_string(args.file.as_path()),
        ps_literal(&args.sheet),
        ps_literal(match args.axis {
            AutoFitAxis::All => "all",
            AutoFitAxis::Rows => "rows",
            AutoFitAxis::Columns => "columns",
        }),
        range = range,
        autofit_columns = ps_bool(autofit_columns),
        autofit_rows = ps_bool(autofit_rows),
        keep_open = ps_bool(args.keep_open),
    );
    run_powershell(&script)
}

pub fn excel_format(args: &FormatArgs) -> Result<String> {
    ensure_windows()?;
    let mut operations = Vec::new();
    if let Some(number_format) = &args.number_format {
        operations.push(format!(
            "$target.NumberFormat = {}",
            ps_literal(number_format)
        ));
    }
    if let Some(bold) = args.bold {
        operations.push(format!("$target.Font.Bold = {}", ps_bool(bold)));
    }
    if let Some(italic) = args.italic {
        operations.push(format!("$target.Font.Italic = {}", ps_bool(italic)));
    }
    if let Some(wrap_text) = args.wrap_text {
        operations.push(format!("$target.WrapText = {}", ps_bool(wrap_text)));
    }
    if let Some(fill_color) = &args.fill_color {
        operations.push(format!(
            "$target.Interior.Color = Convert-HexColor {}",
            ps_literal(fill_color)
        ));
    }
    if let Some(font_color) = &args.font_color {
        operations.push(format!(
            "$target.Font.Color = Convert-HexColor {}",
            ps_literal(font_color)
        ));
    }
    if operations.is_empty() {
        bail!("provide at least one formatting option");
    }
    let operations = operations.join("\n  ");
    let script = format!(
        r##"
$ErrorActionPreference = 'Stop'
function Convert-HexColor($value) {{
  $hex = ([string]$value).Trim().TrimStart('#')
  if ($hex.Length -ne 6) {{ throw "Expected RRGGBB color, got '$value'" }}
  $r = [Convert]::ToInt32($hex.Substring(0, 2), 16)
  $g = [Convert]::ToInt32($hex.Substring(2, 2), 16)
  $b = [Convert]::ToInt32($hex.Substring(4, 2), 16)
  return ($b -shl 16) -bor ($g -shl 8) -bor $r
}}
$path = {}
$excel = New-Object -ComObject Excel.Application
$excel.Visible = {keep_open}
$excel.DisplayAlerts = $false
$excel.AskToUpdateLinks = $false
$wb = $excel.Workbooks.Open($path, 0, $false)
try {{
  $ws = $wb.Worksheets.Item({})
  $target = $ws.Range({})
  {operations}
  $wb.Save()
  [ordered]@{{ file = $path; sheet = $ws.Name; range = {}; saved = $true }} | ConvertTo-Json
}} finally {{
  if (-not {keep_open}) {{
    $wb.Close($false)
    $excel.Quit()
  }}
}}
"##,
        ps_string(args.file.as_path()),
        ps_literal(&args.sheet),
        ps_literal(&args.range),
        ps_literal(&args.range),
        operations = operations,
        keep_open = ps_bool(args.keep_open),
    );
    run_powershell(&script)
}

pub fn excel_copy_sheets(args: &CopySheetsArgs) -> Result<String> {
    ensure_windows()?;
    let sheet_specs = args
        .sheets
        .iter()
        .map(|sheet| ps_literal(sheet))
        .collect::<Vec<_>>()
        .join(",");
    let after = args
        .after
        .as_ref()
        .map(|value| ps_literal(value))
        .unwrap_or_else(|| "$null".to_string());
    let script = format!(
        r#"
$ErrorActionPreference = 'Stop'
$targetPath = {}
$sourcePath = {}
$sheetSpecs = @({sheet_specs})
$afterName = {after}
$excel = New-Object -ComObject Excel.Application
$excel.Visible = {keep_open}
$excel.DisplayAlerts = $false
$excel.AskToUpdateLinks = $false
$target = $excel.Workbooks.Open($targetPath, 0, $false)
$source = $excel.Workbooks.Open($sourcePath, 0, $true)
$copied = @()
try {{
  if ($afterName) {{ $after = $target.Worksheets.Item($afterName) }} else {{ $after = $target.Worksheets.Item($target.Worksheets.Count) }}
  foreach ($spec in $sheetSpecs) {{
    $parts = $spec.Split('=', 2)
    $sourceName = $parts[0]
    $targetName = if ($parts.Length -gt 1) {{ $parts[1] }} else {{ $sourceName }}
    $existing = $null
    foreach ($ws in $target.Worksheets) {{ if ($ws.Name -eq $targetName) {{ $existing = $ws; break }} }}
    if ($existing -ne $null) {{
      if (-not {replace}) {{ throw "Target sheet '$targetName' already exists; use --replace" }}
      $existing.Delete()
    }}
    $source.Worksheets.Item($sourceName).Copy([Type]::Missing, $after)
    $newSheet = $target.Worksheets.Item($after.Index + 1)
    $newSheet.Name = $targetName
    $after = $newSheet
    $copied += [ordered]@{{ source = $sourceName; target = $targetName }}
  }}
  $source.Close($false)
  $target.Save()
  [ordered]@{{ file = $targetPath; copied = $copied; saved = $true }} | ConvertTo-Json -Depth 5
}} finally {{
  if ($source -ne $null) {{ try {{ $source.Close($false) }} catch {{}} }}
  if (-not {keep_open}) {{
    $target.Close($false)
    $excel.Quit()
  }}
}}
"#,
        ps_string(args.file.as_path()),
        ps_string(args.source.as_path()),
        sheet_specs = sheet_specs,
        after = after,
        replace = ps_bool(args.replace),
        keep_open = ps_bool(args.keep_open),
    );
    run_powershell(&script)
}

pub fn word_replace(args: &DocxReplaceArgs, replacements: &[Replacement]) -> Result<ReplaceResult> {
    ensure_windows()?;
    let target = args.output.as_ref().unwrap_or(&args.file);
    if let Some(output) = &args.output {
        std::fs::copy(&args.file, output)
            .with_context(|| format!("copy {} to {}", args.file.display(), output.display()))?;
    }
    let replacements_json = serde_json::to_string(replacements)?;
    let script = format!(
        r#"
$ErrorActionPreference = 'Stop'
$path = {}
$replacements = ConvertFrom-Json @'
{replacements_json}
'@
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$doc = $word.Documents.Open($path)
$counts = @()
try {{
  foreach ($rep in $replacements) {{
    $count = 0
    while ($true) {{
      $range = $doc.Content
      $find = $range.Find
      $find.ClearFormatting()
      $find.Text = $rep.find
      $find.Forward = $true
      $find.Wrap = 0
      $find.MatchCase = $false
      $find.MatchWholeWord = $false
      $find.MatchWildcards = $false
      if ($find.Execute()) {{
        $range.Text = $rep.replace
        $count += 1
      }} else {{
        break
      }}
    }}
    $counts += [ordered]@{{ find = $rep.find; count = $count }}
  }}
  $doc.SaveAs2($path, 12)
  [ordered]@{{ file = $path; counts = $counts }} | ConvertTo-Json -Depth 5
}} finally {{
  $doc.Close($false)
  $word.Quit()
}}
"#,
        ps_string(target.as_path()),
    );
    let output = run_powershell(&script)?;
    let value: serde_json::Value = serde_json::from_str(&output)?;
    let counts = value
        .get("counts")
        .and_then(|v| v.as_array())
        .map(|items| {
            items
                .iter()
                .map(|item| ReplacementCount {
                    find: item
                        .get("find")
                        .and_then(|v| v.as_str())
                        .unwrap_or_default()
                        .to_string(),
                    count: item.get("count").and_then(|v| v.as_u64()).unwrap_or(0) as usize,
                })
                .collect::<Vec<_>>()
        })
        .unwrap_or_default();
    Ok(ReplaceResult {
        file: target.to_path_buf(),
        counts,
    })
}

pub fn run_outlook_script(script: &str) -> Result<String> {
    ensure_windows()?;
    run_powershell(script)
}

pub fn office_doctor() -> Result<String> {
    if !cfg!(windows) {
        return Ok(serde_json::to_string_pretty(&serde_json::json!({
            "platform": std::env::consts::OS,
            "wincom": [
                {
                    "name": "Excel.Application",
                    "ok": false,
                    "error": "WinCOM is only available on Windows"
                },
                {
                    "name": "Word.Application",
                    "ok": false,
                    "error": "WinCOM is only available on Windows"
                },
                {
                    "name": "Outlook.Application",
                    "ok": false,
                    "error": "WinCOM is only available on Windows"
                }
            ]
        }))?);
    }
    ensure_windows()?;
    let script = r#"
$ErrorActionPreference = 'Stop'
function Test-Com($name, $probe) {
  $obj = $null
  try {
    $obj = New-Object -ComObject $name
    $detail = & $probe $obj
    return [ordered]@{ name = $name; ok = $true; detail = $detail }
  } catch {
    return [ordered]@{ name = $name; ok = $false; error = $_.Exception.Message }
  } finally {
    if ($obj -ne $null -and ($name -eq 'Excel.Application' -or $name -eq 'Word.Application')) {
      try { $obj.Quit() } catch {}
    }
  }
}
$checks = @()
$checks += Test-Com 'Excel.Application' { param($excel) "version=$($excel.Version)" }
$checks += Test-Com 'Word.Application' { param($word) "version=$($word.Version)" }
$checks += Test-Com 'Outlook.Application' {
  param($outlook)
  $ns = $outlook.GetNamespace('MAPI')
  "stores=$($ns.Folders.Count)"
}
[ordered]@{ platform = 'windows'; wincom = $checks } | ConvertTo-Json -Depth 5
"#;
    run_powershell(script)
}

fn run_powershell(script: &str) -> Result<String> {
    let mut file = tempfile::NamedTempFile::new()?;
    file.write_all(script.as_bytes())?;
    let output = Command::new("powershell.exe")
        .arg("-NoProfile")
        .arg("-ExecutionPolicy")
        .arg("Bypass")
        .arg("-File")
        .arg(file.path())
        .output()
        .context("run powershell.exe")?;
    if !output.status.success() {
        bail!(
            "PowerShell COM command failed (exit {}): {}",
            output.status,
            String::from_utf8_lossy(&output.stderr)
        );
    }
    Ok(String::from_utf8_lossy(&output.stdout).trim().to_string())
}

fn ensure_windows() -> Result<()> {
    if cfg!(windows) {
        Ok(())
    } else {
        bail!("WinCOM commands require Windows with Microsoft Office installed");
    }
}

fn ps_bool(value: bool) -> &'static str {
    if value { "$true" } else { "$false" }
}

pub fn ps_literal(value: &str) -> String {
    format!("'{}'", value.replace('\'', "''"))
}

pub fn ps_string(path: &Path) -> String {
    ps_literal(&path.to_string_lossy())
}
