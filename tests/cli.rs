use office_tools::ooxml::{PartMap, read_part, read_part_to_string, write_new_package};
use std::io::Write;
use std::process::{Command, Stdio};

fn bin() -> &'static str {
    env!("CARGO_BIN_EXE_office-tools")
}

#[test]
fn xlsx_cli_edit_preserves_validation_and_conditional_formatting() -> anyhow::Result<()> {
    let dir = tempfile::tempdir()?;
    let path = dir.path().join("book.xlsx");
    write_new_package(&path, &minimal_xlsx_parts())?;

    let output = Command::new(bin())
        .args([
            "xlsx",
            "edit",
            path.to_str().unwrap(),
            "--sheet",
            "Sheet1",
            "--cell",
            "A1",
            "--value",
            "via-cli",
            "--json",
        ])
        .output()?;

    assert!(
        output.status.success(),
        "stdout:\n{}\nstderr:\n{}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );
    let sheet = read_part_to_string(&path, "xl/worksheets/sheet1.xml")?;
    assert!(sheet.contains("via-cli"));
    assert!(sheet.contains("conditionalFormatting"));
    assert!(sheet.contains("dataValidations"));
    Ok(())
}

#[test]
fn xlsx_cli_cells_lists_addressed_values() -> anyhow::Result<()> {
    let dir = tempfile::tempdir()?;
    let path = dir.path().join("book.xlsx");
    write_new_package(&path, &minimal_xlsx_parts())?;

    let output = Command::new(bin())
        .args([
            "xlsx",
            "cells",
            path.to_str().unwrap(),
            "--sheet",
            "Sheet1",
            "--range",
            "A1:B1",
            "--include-empty",
            "--json",
        ])
        .output()?;
    assert!(
        output.status.success(),
        "stdout:\n{}\nstderr:\n{}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("\"cell\": \"A1\""));
    assert!(stdout.contains("\"value\": \"old\""));
    assert!(stdout.contains("\"cell\": \"B1\""));
    Ok(())
}

#[test]
fn xlsx_cli_list_sheets_reports_sheet_metadata() -> anyhow::Result<()> {
    let dir = tempfile::tempdir()?;
    let path = dir.path().join("book.xlsx");
    let mut parts = minimal_xlsx_parts();
    let workbook = String::from_utf8(parts.remove("xl/workbook.xml").unwrap())?;
    parts.insert(
        "xl/workbook.xml".to_string(),
        workbook
            .replace(
                r#"<sheet name="Sheet1" sheetId="1" r:id="rId1"/>"#,
                r#"<sheet name="Sheet1" sheetId="7" state="hidden" r:id="rId1"/>"#,
            )
            .into_bytes(),
    );
    write_new_package(&path, &parts)?;

    let output = Command::new(bin())
        .args(["xlsx", "list-sheets", path.to_str().unwrap(), "--json"])
        .output()?;
    assert!(
        output.status.success(),
        "stdout:\n{}\nstderr:\n{}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("\"name\": \"Sheet1\""));
    assert!(stdout.contains("\"sheet_id\": \"7\""));
    assert!(stdout.contains("\"state\": \"hidden\""));
    Ok(())
}

#[test]
fn xlsx_cli_relationships_lists_sheet_targets() -> anyhow::Result<()> {
    let dir = tempfile::tempdir()?;
    let path = dir.path().join("book.xlsx");
    write_new_package(&path, &minimal_xlsx_parts())?;

    let output = Command::new(bin())
        .args([
            "xlsx",
            "relationships",
            path.to_str().unwrap(),
            "--sheet",
            "Sheet1",
            "--json",
        ])
        .output()?;
    assert!(
        output.status.success(),
        "stdout:\n{}\nstderr:\n{}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("\"scope\": \"worksheet\""));
    assert!(stdout.contains("relationships/hyperlink"));
    assert!(stdout.contains("https://example.com"));
    assert!(stdout.contains("\"target_mode\": \"External\""));
    Ok(())
}

#[test]
fn xlsx_cli_tables_lists_table_metadata() -> anyhow::Result<()> {
    let dir = tempfile::tempdir()?;
    let path = dir.path().join("book.xlsx");
    write_new_package(&path, &minimal_xlsx_parts())?;

    let output = Command::new(bin())
        .args(["xlsx", "tables", path.to_str().unwrap(), "--json"])
        .output()?;
    assert!(
        output.status.success(),
        "stdout:\n{}\nstderr:\n{}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("Table1"));
    assert!(stdout.contains("A1:B2"));
    Ok(())
}

#[test]
fn xlsx_cli_validations_lists_validation_metadata() -> anyhow::Result<()> {
    let dir = tempfile::tempdir()?;
    let path = dir.path().join("book.xlsx");
    write_new_package(&path, &minimal_xlsx_parts())?;

    let output = Command::new(bin())
        .args(["xlsx", "validations", path.to_str().unwrap(), "--json"])
        .output()?;
    assert!(
        output.status.success(),
        "stdout:\n{}\nstderr:\n{}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("B1:B10"));
    assert!(stdout.contains("whole"));
    assert!(stdout.contains("\"formula1\": \"1\""));
    assert!(stdout.contains("\"formula2\": \"10\""));
    Ok(())
}

#[test]
fn xlsx_cli_conditional_formatting_lists_rule_metadata() -> anyhow::Result<()> {
    let dir = tempfile::tempdir()?;
    let path = dir.path().join("book.xlsx");
    write_new_package(&path, &minimal_xlsx_parts())?;

    let output = Command::new(bin())
        .args([
            "xlsx",
            "conditional-formatting",
            path.to_str().unwrap(),
            "--json",
        ])
        .output()?;
    assert!(
        output.status.success(),
        "stdout:\n{}\nstderr:\n{}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("A1:A10"));
    assert!(stdout.contains("cellIs"));
    assert!(stdout.contains("greaterThan"));
    assert!(stdout.contains("\"0\""));
    Ok(())
}

#[test]
fn xlsx_cli_hyperlinks_lists_targets() -> anyhow::Result<()> {
    let dir = tempfile::tempdir()?;
    let path = dir.path().join("book.xlsx");
    write_new_package(&path, &minimal_xlsx_parts())?;

    let output = Command::new(bin())
        .args(["xlsx", "hyperlinks", path.to_str().unwrap(), "--json"])
        .output()?;
    assert!(
        output.status.success(),
        "stdout:\n{}\nstderr:\n{}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("https://example.com"));
    assert!(stdout.contains("Example"));
    assert!(stdout.contains("A1"));
    Ok(())
}

#[test]
fn xlsx_cli_comments_lists_comment_text() -> anyhow::Result<()> {
    let dir = tempfile::tempdir()?;
    let path = dir.path().join("book.xlsx");
    write_new_package(&path, &minimal_xlsx_parts())?;

    let output = Command::new(bin())
        .args(["xlsx", "comments", path.to_str().unwrap(), "--json"])
        .output()?;
    assert!(
        output.status.success(),
        "stdout:\n{}\nstderr:\n{}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("Reviewer"));
    assert!(stdout.contains("Check this cell"));
    assert!(stdout.contains("A1"));
    Ok(())
}

#[test]
fn xlsx_cli_defined_names_lists_scoped_names() -> anyhow::Result<()> {
    let dir = tempfile::tempdir()?;
    let path = dir.path().join("book.xlsx");
    write_new_package(&path, &minimal_xlsx_parts())?;

    let output = Command::new(bin())
        .args(["xlsx", "defined-names", path.to_str().unwrap(), "--json"])
        .output()?;
    assert!(
        output.status.success(),
        "stdout:\n{}\nstderr:\n{}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("TotalCell"));
    assert!(stdout.contains("Sheet1!$B$2"));
    assert!(stdout.contains("_xlnm.Print_Titles"));
    assert!(stdout.contains("\"sheet\": \"Sheet1\""));
    Ok(())
}

#[test]
fn xlsx_cli_merged_ranges_lists_ranges() -> anyhow::Result<()> {
    let dir = tempfile::tempdir()?;
    let path = dir.path().join("book.xlsx");
    write_new_package(&path, &minimal_xlsx_parts())?;

    let output = Command::new(bin())
        .args([
            "xlsx",
            "merged-ranges",
            path.to_str().unwrap(),
            "--sheet",
            "Sheet1",
            "--json",
        ])
        .output()?;
    assert!(
        output.status.success(),
        "stdout:\n{}\nstderr:\n{}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("\"sheet\": \"Sheet1\""));
    assert!(stdout.contains("\"range\": \"A1:B1\""));
    Ok(())
}

#[test]
fn xlsx_cli_auto_filters_lists_filter_ranges() -> anyhow::Result<()> {
    let dir = tempfile::tempdir()?;
    let path = dir.path().join("book.xlsx");
    write_new_package(&path, &minimal_xlsx_parts())?;

    let output = Command::new(bin())
        .args([
            "xlsx",
            "auto-filters",
            path.to_str().unwrap(),
            "--sheet",
            "Sheet1",
            "--json",
        ])
        .output()?;
    assert!(
        output.status.success(),
        "stdout:\n{}\nstderr:\n{}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("\"sheet\": \"Sheet1\""));
    assert!(stdout.contains("\"range\": \"A1:B10\""));
    Ok(())
}

#[test]
fn xlsx_cli_protections_lists_sheet_protection_settings() -> anyhow::Result<()> {
    let dir = tempfile::tempdir()?;
    let path = dir.path().join("book.xlsx");
    let mut parts = minimal_xlsx_parts();
    let sheet = String::from_utf8(parts.remove("xl/worksheets/sheet1.xml").unwrap())?;
    parts.insert(
        "xl/worksheets/sheet1.xml".to_string(),
        sheet
            .replace(
                "<sheetData>",
                r#"<sheetProtection sheet="1" objects="1" scenarios="1"/><sheetData>"#,
            )
            .into_bytes(),
    );
    write_new_package(&path, &parts)?;

    let output = Command::new(bin())
        .args([
            "xlsx",
            "protections",
            path.to_str().unwrap(),
            "--sheet",
            "Sheet1",
            "--json",
        ])
        .output()?;
    assert!(
        output.status.success(),
        "stdout:\n{}\nstderr:\n{}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("\"sheet\": \"Sheet1\""));
    assert!(stdout.contains("\"objects\": \"1\""));
    assert!(stdout.contains("\"scenarios\": \"1\""));
    Ok(())
}

#[test]
fn xlsx_cli_create_builds_workbook_from_json() -> anyhow::Result<()> {
    let dir = tempfile::tempdir()?;
    let spec = dir.path().join("workbook.json");
    let output_xlsx = dir.path().join("created.xlsx");
    std::fs::write(
        &spec,
        serde_json::json!({
            "sheets": [
                {
                    "name": "Data",
                    "rows": [["Label", "Value"], ["Revenue", 123]],
                    "cells": [{ "cell": "C2", "value": "=B2*2" }]
                },
                { "name": "Notes", "rows": [["Created without Python"]] }
            ]
        })
        .to_string(),
    )?;

    let create = Command::new(bin())
        .args([
            "xlsx",
            "create",
            spec.to_str().unwrap(),
            "--output",
            output_xlsx.to_str().unwrap(),
            "--json",
        ])
        .output()?;
    assert!(
        create.status.success(),
        "stdout:\n{}\nstderr:\n{}",
        String::from_utf8_lossy(&create.stdout),
        String::from_utf8_lossy(&create.stderr)
    );

    let workbook_xml = read_part_to_string(&output_xlsx, "xl/workbook.xml")?;
    assert!(workbook_xml.contains("Data"));
    assert!(workbook_xml.contains("Notes"));
    assert!(workbook_xml.contains("fullCalcOnLoad"));
    let sheet_xml = read_part_to_string(&output_xlsx, "xl/worksheets/sheet1.xml")?;
    assert!(sheet_xml.contains("Revenue"));
    assert!(sheet_xml.contains("<f>B2*2</f>"));

    let formulas = Command::new(bin())
        .args([
            "xlsx",
            "formulas",
            output_xlsx.to_str().unwrap(),
            "--sheet",
            "Data",
            "--json",
        ])
        .output()?;
    assert!(formulas.status.success());
    assert!(String::from_utf8_lossy(&formulas.stdout).contains("B2*2"));

    let read = Command::new(bin())
        .args([
            "xlsx",
            "read",
            output_xlsx.to_str().unwrap(),
            "--sheet",
            "Data",
            "--range",
            "A1:C2",
            "--json",
        ])
        .output()?;
    assert!(read.status.success());
    assert!(String::from_utf8_lossy(&read.stdout).contains("Revenue"));
    Ok(())
}

#[test]
fn docx_cli_compose_then_read() -> anyhow::Result<()> {
    let dir = tempfile::tempdir()?;
    let spec = dir.path().join("docx.json");
    let output_docx = dir.path().join("out.docx");
    std::fs::write(
        &spec,
        serde_json::json!({
            "meta": { "title": "Memo", "footer_text": "Generic footer" },
            "blocks": [
                { "type": "heading", "text": "Memo", "level": 1 },
                { "type": "body", "text": "Body text", "footnote": "Footnote text" }
            ]
        })
        .to_string(),
    )?;

    let compose = Command::new(bin())
        .args([
            "docx",
            "compose",
            spec.to_str().unwrap(),
            "--output",
            output_docx.to_str().unwrap(),
        ])
        .output()?;
    assert!(compose.status.success());

    let read = Command::new(bin())
        .args(["docx", "read", output_docx.to_str().unwrap()])
        .output()?;
    assert!(read.status.success());
    let text = String::from_utf8_lossy(&read.stdout);
    assert!(text.contains("Memo"));
    assert!(text.contains("Body text"));

    let read_json = Command::new(bin())
        .args(["docx", "read", output_docx.to_str().unwrap(), "--json"])
        .output()?;
    assert!(read_json.status.success());
    let json_text = String::from_utf8_lossy(&read_json.stdout);
    assert!(json_text.contains("\"markdown\""));
    assert!(json_text.contains("Body text"));

    let footnotes = read_part_to_string(&output_docx, "word/footnotes.xml")?;
    assert!(footnotes.contains("Footnote text"));
    Ok(())
}

#[test]
fn pptx_cli_build_then_read_notes() -> anyhow::Result<()> {
    let dir = tempfile::tempdir()?;
    let spec = dir.path().join("deck.json");
    let output_pptx = dir.path().join("deck.pptx");
    std::fs::write(
        &spec,
        serde_json::json!({
            "meta": { "title": "Deck", "footer_text": "Footer" },
            "slides": [
                { "layout": "content", "title": "Slide title", "bullets": ["Point"], "notes": "Speaker note" }
            ]
        })
        .to_string(),
    )?;

    let build = Command::new(bin())
        .args([
            "pptx",
            "build",
            spec.to_str().unwrap(),
            "--output",
            output_pptx.to_str().unwrap(),
        ])
        .output()?;
    assert!(build.status.success());

    let read = Command::new(bin())
        .args(["pptx", "read", output_pptx.to_str().unwrap()])
        .output()?;
    assert!(read.status.success());
    assert!(String::from_utf8_lossy(&read.stdout).contains("Slide title"));

    let read_json = Command::new(bin())
        .args(["pptx", "read", output_pptx.to_str().unwrap(), "--json"])
        .output()?;
    assert!(read_json.status.success());
    let read_json_text = String::from_utf8_lossy(&read_json.stdout);
    assert!(read_json_text.contains("\"index\": 1"));
    assert!(read_json_text.contains("Slide title"));

    let notes = Command::new(bin())
        .args(["pptx", "notes", output_pptx.to_str().unwrap()])
        .output()?;
    assert!(notes.status.success());
    assert!(String::from_utf8_lossy(&notes.stdout).contains("Speaker note"));

    let notes_json = Command::new(bin())
        .args(["pptx", "notes", output_pptx.to_str().unwrap(), "--json"])
        .output()?;
    assert!(notes_json.status.success());
    let notes_json_text = String::from_utf8_lossy(&notes_json.stdout);
    assert!(notes_json_text.contains("\"index\": 1"));
    assert!(notes_json_text.contains("Speaker note"));
    Ok(())
}

#[test]
fn package_cli_list_parts_reports_content_types() -> anyhow::Result<()> {
    let dir = tempfile::tempdir()?;
    let path = dir.path().join("book.xlsx");
    write_new_package(&path, &minimal_xlsx_parts())?;

    let output = Command::new(bin())
        .args(["package", "list-parts", path.to_str().unwrap(), "--json"])
        .output()?;
    assert!(
        output.status.success(),
        "stdout:\n{}\nstderr:\n{}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("\"name\": \"xl/workbook.xml\""));
    assert!(stdout.contains(
        "\"content_type\": \"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\""
    ));
    assert!(stdout.contains("\"content_type\": \"application/xml\""));
    Ok(())
}

#[test]
fn package_cli_read_part_can_emit_base64() -> anyhow::Result<()> {
    let dir = tempfile::tempdir()?;
    let path = dir.path().join("package.zip");
    let mut parts = minimal_xlsx_parts();
    parts.insert("custom/data.bin".to_string(), vec![0, 1, 2, 255]);
    write_new_package(&path, &parts)?;

    let output = Command::new(bin())
        .args([
            "package",
            "read-part",
            path.to_str().unwrap(),
            "custom/data.bin",
            "--base64",
        ])
        .output()?;
    assert!(
        output.status.success(),
        "stdout:\n{}\nstderr:\n{}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );
    assert_eq!(String::from_utf8(output.stdout)?, "AAEC/w==");
    Ok(())
}

#[test]
fn package_cli_write_part_accepts_base64() -> anyhow::Result<()> {
    let dir = tempfile::tempdir()?;
    let path = dir.path().join("package.zip");
    write_new_package(&path, &minimal_xlsx_parts())?;

    let output = Command::new(bin())
        .args([
            "package",
            "write-part",
            path.to_str().unwrap(),
            "custom/data.bin",
            "--base64",
            "AAEC/w==",
        ])
        .output()?;
    assert!(
        output.status.success(),
        "stdout:\n{}\nstderr:\n{}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );
    assert_eq!(read_part(&path, "custom/data.bin")?, vec![0, 1, 2, 255]);
    Ok(())
}

#[test]
fn package_cli_write_part_accepts_text() -> anyhow::Result<()> {
    let dir = tempfile::tempdir()?;
    let path = dir.path().join("package.zip");
    write_new_package(&path, &minimal_xlsx_parts())?;

    let output = Command::new(bin())
        .args([
            "package",
            "write-part",
            path.to_str().unwrap(),
            "custom/text.xml",
            "--text",
            "<root>hello</root>",
        ])
        .output()?;
    assert!(
        output.status.success(),
        "stdout:\n{}\nstderr:\n{}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );
    assert_eq!(
        read_part_to_string(&path, "custom/text.xml")?,
        "<root>hello</root>"
    );
    Ok(())
}

#[test]
fn package_cli_write_part_rejects_ambiguous_payload_sources() -> anyhow::Result<()> {
    let dir = tempfile::tempdir()?;
    let path = dir.path().join("package.zip");
    let input = dir.path().join("payload.txt");
    write_new_package(&path, &minimal_xlsx_parts())?;
    std::fs::write(&input, "payload")?;

    let output = Command::new(bin())
        .args([
            "package",
            "write-part",
            path.to_str().unwrap(),
            "custom/data.bin",
            "--input",
            input.to_str().unwrap(),
            "--base64",
            "AAEC/w==",
        ])
        .output()?;
    assert!(!output.status.success());
    assert!(
        String::from_utf8_lossy(&output.stderr)
            .contains("provide exactly one of --input, --base64, or --text")
    );
    Ok(())
}

#[test]
fn office_doctor_returns_machine_readable_json() -> anyhow::Result<()> {
    let output = Command::new(bin()).args(["doctor"]).output()?;
    assert!(
        output.status.success(),
        "stdout:\n{}\nstderr:\n{}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8_lossy(&output.stdout);
    assert!(stdout.contains("\"platform\""));
    assert!(stdout.contains("\"wincom\""));
    assert!(stdout.contains("Excel.Application"));
    Ok(())
}

#[test]
fn mcp_office_doctor_returns_machine_readable_json() -> anyhow::Result<()> {
    let response = mcp_request(serde_json::json!({
        "jsonrpc": "2.0",
        "id": 1,
        "method": "tools/call",
        "params": {
            "name": "office_doctor",
            "arguments": {}
        }
    }))?;
    assert_eq!(response["result"]["isError"].as_bool(), Some(false));
    let text = response["result"]["content"][0]["text"]
        .as_str()
        .unwrap_or_default();
    let value: serde_json::Value = serde_json::from_str(text)?;
    assert!(value.get("platform").is_some());
    assert!(
        value
            .get("wincom")
            .and_then(|value| value.as_array())
            .is_some()
    );
    assert!(text.contains("Excel.Application"));
    Ok(())
}

#[test]
fn mcp_tools_list_exposes_office_primitives() -> anyhow::Result<()> {
    let response = mcp_request(serde_json::json!({
        "jsonrpc": "2.0",
        "id": 1,
        "method": "tools/list",
        "params": {}
    }))?;
    let stdout = response.to_string();
    for tool in [
        "xlsx_edit",
        "xlsx_cells",
        "xlsx_create",
        "xlsx_relationships",
        "xlsx_formulas",
        "xlsx_tables",
        "xlsx_validations",
        "xlsx_conditional_formatting",
        "xlsx_hyperlinks",
        "xlsx_comments",
        "xlsx_defined_names",
        "xlsx_merged_ranges",
        "xlsx_auto_filters",
        "xlsx_protections",
        "xlsx_validate",
        "xlsx_insert",
        "xlsx_autofit",
        "xlsx_format",
        "xlsx_rename_sheet",
        "xlsx_copy_sheets",
        "docx_compose",
        "engine",
        "com",
        "pptx_build",
        "office_doctor",
        "package_write_part",
        "package_delete_part",
        "oneOf",
        "date1904",
        "blocks",
        "slides",
    ] {
        assert!(stdout.contains(tool), "missing MCP tool {tool}");
    }
    Ok(())
}

#[test]
fn mcp_xlsx_edit_accepts_json_scalar_values() -> anyhow::Result<()> {
    let dir = tempfile::tempdir()?;
    let path = dir.path().join("book.xlsx");
    write_new_package(&path, &minimal_xlsx_parts())?;

    let response = mcp_request(serde_json::json!({
        "jsonrpc": "2.0",
        "id": 1,
        "method": "tools/call",
        "params": {
            "name": "xlsx_edit",
            "arguments": {
                "file": path,
                "sheet": "Sheet1",
                "cell": "B2",
                "value": 456
            }
        }
    }))?;
    assert!(response["result"]["isError"].as_bool() == Some(false));
    let sheet = read_part_to_string(&path, "xl/worksheets/sheet1.xml")?;
    assert!(sheet.contains(r#"<c r="B2"><v>456</v></c>"#));
    Ok(())
}

#[test]
fn mcp_can_compose_docx_and_write_delete_package_parts() -> anyhow::Result<()> {
    let dir = tempfile::tempdir()?;
    let output_docx = dir.path().join("mcp.docx");

    let response = mcp_request(serde_json::json!({
        "jsonrpc": "2.0",
        "id": 1,
        "method": "tools/call",
        "params": {
            "name": "docx_compose",
            "arguments": {
                "output": output_docx,
                "spec": {
                    "meta": { "title": "MCP Memo" },
                    "blocks": [
                        { "type": "heading", "text": "MCP Memo", "level": 1 },
                        { "type": "body", "text": "Composed through MCP" }
                    ]
                }
            }
        }
    }))?;
    assert!(response["result"]["isError"].as_bool() == Some(false));
    let text = read_part_to_string(&output_docx, "word/document.xml")?;
    assert!(text.contains("Composed through MCP"));

    let list_response = mcp_request(serde_json::json!({
        "jsonrpc": "2.0",
        "id": 5,
        "method": "tools/call",
        "params": {
            "name": "package_list_parts",
            "arguments": { "file": output_docx }
        }
    }))?;
    let listed_parts = list_response["result"]["content"][0]["text"]
        .as_str()
        .unwrap_or_default();
    assert!(listed_parts.contains("\"name\": \"word/document.xml\""));
    assert!(listed_parts.contains("wordprocessingml.document.main+xml"));

    let write_response = mcp_request(serde_json::json!({
        "jsonrpc": "2.0",
        "id": 2,
        "method": "tools/call",
        "params": {
            "name": "package_write_part",
            "arguments": {
                "file": output_docx,
                "part": "custom/data.bin",
                "base64": "aGVsbG8="
            }
        }
    }))?;
    assert!(write_response["result"]["isError"].as_bool() == Some(false));
    assert_eq!(read_part(&output_docx, "custom/data.bin")?, b"hello");

    let read_response = mcp_request(serde_json::json!({
        "jsonrpc": "2.0",
        "id": 3,
        "method": "tools/call",
        "params": {
            "name": "package_read_part",
            "arguments": {
                "file": output_docx,
                "part": "custom/data.bin",
                "base64": true
            }
        }
    }))?;
    assert_eq!(
        read_response["result"]["content"][0]["text"].as_str(),
        Some("aGVsbG8=")
    );

    let delete_response = mcp_request(serde_json::json!({
        "jsonrpc": "2.0",
        "id": 4,
        "method": "tools/call",
        "params": {
            "name": "package_delete_part",
            "arguments": { "file": output_docx, "parts": ["custom/data.bin"] }
        }
    }))?;
    assert!(delete_response["result"]["isError"].as_bool() == Some(false));
    assert!(read_part(&output_docx, "custom/data.bin").is_err());
    Ok(())
}

#[test]
fn mcp_docx_replace_accepts_engine_argument() -> anyhow::Result<()> {
    let dir = tempfile::tempdir()?;
    let output_docx = dir.path().join("replace.docx");

    let compose = mcp_request(serde_json::json!({
        "jsonrpc": "2.0",
        "id": 1,
        "method": "tools/call",
        "params": {
            "name": "docx_compose",
            "arguments": {
                "output": output_docx,
                "spec": {
                    "blocks": [
                        { "type": "body", "text": "old text" }
                    ]
                }
            }
        }
    }))?;
    assert!(compose["result"]["isError"].as_bool() == Some(false));

    let replace = mcp_request(serde_json::json!({
        "jsonrpc": "2.0",
        "id": 2,
        "method": "tools/call",
        "params": {
            "name": "docx_replace",
            "arguments": {
                "file": output_docx,
                "find": "old",
                "replace": "new",
                "engine": "direct"
            }
        }
    }))?;
    assert!(replace["result"]["isError"].as_bool() == Some(false));
    assert!(read_part_to_string(&output_docx, "word/document.xml")?.contains("new text"));

    let read = mcp_request(serde_json::json!({
        "jsonrpc": "2.0",
        "id": 3,
        "method": "tools/call",
        "params": {
            "name": "docx_read",
            "arguments": { "file": output_docx }
        }
    }))?;
    assert!(
        read["result"]["content"][0]["text"]
            .as_str()
            .unwrap_or_default()
            .contains("new text")
    );
    Ok(())
}

#[test]
fn mcp_can_build_and_read_pptx_notes() -> anyhow::Result<()> {
    let dir = tempfile::tempdir()?;
    let output_pptx = dir.path().join("mcp-deck.pptx");

    let build = mcp_request(serde_json::json!({
        "jsonrpc": "2.0",
        "id": 1,
        "method": "tools/call",
        "params": {
            "name": "pptx_build",
            "arguments": {
                "output": output_pptx,
                "spec": {
                    "slides": [
                        {
                            "layout": "content",
                            "title": "MCP slide",
                            "bullets": ["Point"],
                            "notes": "MCP speaker note"
                        }
                    ]
                }
            }
        }
    }))?;
    assert!(build["result"]["isError"].as_bool() == Some(false));

    let read = mcp_request(serde_json::json!({
        "jsonrpc": "2.0",
        "id": 2,
        "method": "tools/call",
        "params": {
            "name": "pptx_read",
            "arguments": { "file": output_pptx }
        }
    }))?;
    assert!(
        read["result"]["content"][0]["text"]
            .as_str()
            .unwrap_or_default()
            .contains("MCP slide")
    );

    let notes = mcp_request(serde_json::json!({
        "jsonrpc": "2.0",
        "id": 3,
        "method": "tools/call",
        "params": {
            "name": "pptx_notes",
            "arguments": { "file": output_pptx }
        }
    }))?;
    assert!(
        notes["result"]["content"][0]["text"]
            .as_str()
            .unwrap_or_default()
            .contains("MCP speaker note")
    );
    Ok(())
}

#[test]
fn mcp_package_write_part_rejects_ambiguous_payloads() -> anyhow::Result<()> {
    let dir = tempfile::tempdir()?;
    let path = dir.path().join("book.xlsx");
    write_new_package(&path, &minimal_xlsx_parts())?;

    let response = mcp_request(serde_json::json!({
        "jsonrpc": "2.0",
        "id": 1,
        "method": "tools/call",
        "params": {
            "name": "package_write_part",
            "arguments": {
                "file": path,
                "part": "custom/data.bin",
                "text": "hello",
                "base64": "aGVsbG8="
            }
        }
    }))?;
    assert_eq!(response["error"]["code"].as_i64(), Some(-32000));
    assert!(
        response["error"]["message"]
            .as_str()
            .unwrap_or_default()
            .contains("provide exactly one of text or base64")
    );
    Ok(())
}

fn mcp_request(request: serde_json::Value) -> anyhow::Result<serde_json::Value> {
    let request = request.to_string();
    let mut child = Command::new(bin())
        .args(["mcp", "serve"])
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .spawn()?;
    {
        let stdin = child.stdin.as_mut().unwrap();
        stdin.write_all(request.as_bytes())?;
        stdin.write_all(b"\n")?;
    }
    drop(child.stdin.take());
    let output = child.wait_with_output()?;
    assert!(
        output.status.success(),
        "stdout:\n{}\nstderr:\n{}",
        String::from_utf8_lossy(&output.stdout),
        String::from_utf8_lossy(&output.stderr)
    );
    let stdout = String::from_utf8(output.stdout)?;
    let first_line = stdout
        .lines()
        .next()
        .ok_or_else(|| anyhow::anyhow!("MCP response missing first line: {stdout}"))?;
    Ok(serde_json::from_str(first_line)?)
}

fn minimal_xlsx_parts() -> PartMap {
    let mut parts = PartMap::new();
    parts.insert(
        "[Content_Types].xml".to_string(),
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
<Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>
</Types>"#
            .to_vec(),
    );
    parts.insert(
        "_rels/.rels".to_string(),
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"#
            .to_vec(),
    );
    parts.insert(
        "xl/workbook.xml".to_string(),
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
<definedNames>
<definedName name="TotalCell">Sheet1!$B$2</definedName>
<definedName name="_xlnm.Print_Titles" localSheetId="0">'Sheet1'!$1:$1</definedName>
</definedNames>
</workbook>"#
            .to_vec(),
    );
    parts.insert(
        "xl/_rels/workbook.xml.rels".to_string(),
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"#
            .to_vec(),
    );
    parts.insert(
        "xl/worksheets/sheet1.xml".to_string(),
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<dimension ref="A1:B2"/>
<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>old</t></is></c></row></sheetData>
<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>
<autoFilter ref="A1:B10"/>
<conditionalFormatting sqref="A1:A10"><cfRule type="cellIs" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>
<dataValidations count="1"><dataValidation type="whole" allowBlank="1" sqref="B1:B10"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>
<hyperlinks><hyperlink ref="A1" r:id="rId2" display="Example"/></hyperlinks>
</worksheet>"#
            .to_vec(),
    );
    parts.insert(
        "xl/worksheets/_rels/sheet1.xml.rels".to_string(),
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/>
</Relationships>"#
            .to_vec(),
    );
    parts.insert(
        "xl/tables/table1.xml".to_string(),
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:B2" totalsRowShown="0">
<autoFilter ref="A1:B2"/>
<tableColumns count="2"><tableColumn id="1" name="Column1"/><tableColumn id="2" name="Column2"/></tableColumns>
<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>
</table>"#
            .to_vec(),
    );
    parts.insert(
        "xl/comments1.xml".to_string(),
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<authors><author>Reviewer</author></authors>
<commentList><comment ref="A1" authorId="0"><text><r><t>Check this cell</t></r></text></comment></commentList>
</comments>"#
            .to_vec(),
    );
    parts
}
