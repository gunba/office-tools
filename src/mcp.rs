use crate::docx::{ComposeSpec, DocxReplaceArgs, DocxReplaceEngine, Replacement, compose_docx};
use crate::ooxml::{PartMap, list_parts, read_part, rewrite_parts};
use crate::outlook::OutlookArgs;
use crate::pptx::{DeckSpec, build_deck, read_notes, read_slides};
use crate::xlsx::{
    AutoFiltersArgs, AutoFitArgs, AutoFitAxis, CellsArgs, CommentsArgs, ConditionalFormattingArgs,
    CopySheetsArgs, CreateWorkbookSpec, DefinedNamesArgs, EditArgs, FormatArgs, FormulasArgs,
    HyperlinksArgs, InsertArgs, InsertAxis, InspectArgs, MergedRangesArgs, ProtectionsArgs,
    ReadArgs, RelationshipsArgs, RenameSheetArgs, SearchArgs, TablesArgs, ValidateArgs,
    ValidationsArgs, add_sheet, auto_filters, cells, comments, conditional_formatting,
    create_workbook, defined_names, edit, formulas, hyperlinks, inspect, list_sheets,
    merged_ranges, move_sheet, protections, read, relationships, search, tables, validations,
};
use anyhow::{Context, Result, anyhow, bail};
use base64::Engine;
use base64::engine::general_purpose::STANDARD as BASE64_STANDARD;
use clap::{Args, Subcommand};
use serde_json::{Value, json};
use std::io::{self, BufRead, Write};
use std::path::PathBuf;

#[derive(Debug, Args)]
pub struct McpArgs {
    #[command(subcommand)]
    command: McpCommand,
}

#[derive(Debug, Subcommand)]
pub enum McpCommand {
    /// Serve MCP over stdio.
    Serve,
}

impl McpArgs {
    pub fn run(self) -> Result<()> {
        match self.command {
            McpCommand::Serve => serve_stdio(),
        }
    }
}

pub fn serve_stdio() -> Result<()> {
    let stdin = io::stdin();
    let mut reader = io::BufReader::new(stdin.lock());
    let stdout = io::stdout();
    let mut writer = stdout.lock();
    while let Some(message) = read_message(&mut reader)? {
        let Some(response) = handle_message(message) else {
            continue;
        };
        write_message(&mut writer, &response)?;
    }
    Ok(())
}

fn handle_message(message: Value) -> Option<Value> {
    let id = message.get("id").cloned();
    let method = message
        .get("method")
        .and_then(Value::as_str)
        .unwrap_or_default();
    let id = id?;
    let result = match method {
        "initialize" => Ok(json!({
            "protocolVersion": "2024-11-05",
            "serverInfo": { "name": "office-tools", "version": env!("CARGO_PKG_VERSION") },
            "capabilities": { "tools": {} }
        })),
        "ping" => Ok(json!({})),
        "tools/list" => Ok(json!({ "tools": tool_list() })),
        "tools/call" => {
            let params = message.get("params").cloned().unwrap_or_else(|| json!({}));
            call_tool(params).map(
                |text| json!({ "content": [{ "type": "text", "text": text }], "isError": false }),
            )
        }
        _ => Err(anyhow!("unsupported MCP method: {method}")),
    };
    Some(match result {
        Ok(result) => json!({ "jsonrpc": "2.0", "id": id, "result": result }),
        Err(error) => json!({
            "jsonrpc": "2.0",
            "id": id,
            "error": { "code": -32000, "message": error.to_string() }
        }),
    })
}

fn tool_list() -> Vec<Value> {
    vec![
        tool(
            "xlsx_read",
            "Read an XLSX sheet or range without saving the workbook.",
            json!({
                "type": "object", "properties": {
                    "file": { "type": "string" },
                    "sheet": { "type": "string" },
                    "range": { "type": "string" },
                    "show_empty_rows": { "type": "boolean" },
                    "show_empty_cols": { "type": "boolean" }
                }, "required": ["file"]
            }),
        ),
        tool(
            "xlsx_cells",
            "List XLSX cells as address/value records without saving the workbook.",
            json!({
                "type": "object", "properties": {
                    "file": { "type": "string" },
                    "sheet": { "type": "string" },
                    "range": { "type": "string" },
                    "include_empty": { "type": "boolean" }
                }, "required": ["file"]
            }),
        ),
        tool(
            "xlsx_list_sheets",
            "List XLSX workbook sheets.",
            json!({
                "type": "object", "properties": { "file": { "type": "string" } }, "required": ["file"]
            }),
        ),
        tool(
            "xlsx_search",
            "Search visible XLSX cell values.",
            json!({
                "type": "object", "properties": { "file": { "type": "string" }, "query": { "type": "string" }, "sheet": { "type": "string" } }, "required": ["file", "query"]
            }),
        ),
        tool(
            "xlsx_inspect",
            "Inspect XLSX workbook OOXML structure, formulas, validation, and formatting metadata.",
            json!({
                "type": "object", "properties": { "file": { "type": "string" }, "sheet": { "type": "string" } }, "required": ["file"]
            }),
        ),
        tool(
            "xlsx_relationships",
            "List XLSX workbook and worksheet relationships without opening or saving the workbook.",
            json!({
                "type": "object", "properties": { "file": { "type": "string" }, "sheet": { "type": "string" } }, "required": ["file"]
            }),
        ),
        tool(
            "xlsx_formulas",
            "List XLSX formulas by sheet and cell without opening or saving the workbook.",
            json!({
                "type": "object", "properties": { "file": { "type": "string" }, "sheet": { "type": "string" } }, "required": ["file"]
            }),
        ),
        tool(
            "xlsx_tables",
            "List XLSX table names and ranges without opening or saving the workbook.",
            json!({
                "type": "object", "properties": { "file": { "type": "string" }, "sheet": { "type": "string" } }, "required": ["file"]
            }),
        ),
        tool(
            "xlsx_validations",
            "List XLSX data validation rules without opening or saving the workbook.",
            json!({
                "type": "object", "properties": { "file": { "type": "string" }, "sheet": { "type": "string" } }, "required": ["file"]
            }),
        ),
        tool(
            "xlsx_conditional_formatting",
            "List XLSX conditional formatting rules without opening or saving the workbook.",
            json!({
                "type": "object", "properties": { "file": { "type": "string" }, "sheet": { "type": "string" } }, "required": ["file"]
            }),
        ),
        tool(
            "xlsx_hyperlinks",
            "List XLSX hyperlinks and targets without opening or saving the workbook.",
            json!({
                "type": "object", "properties": { "file": { "type": "string" }, "sheet": { "type": "string" } }, "required": ["file"]
            }),
        ),
        tool(
            "xlsx_comments",
            "List XLSX comments/notes without opening or saving the workbook.",
            json!({
                "type": "object", "properties": { "file": { "type": "string" }, "sheet": { "type": "string" } }, "required": ["file"]
            }),
        ),
        tool(
            "xlsx_defined_names",
            "List XLSX workbook defined names and scoped named ranges without opening or saving the workbook.",
            json!({
                "type": "object", "properties": { "file": { "type": "string" }, "name": { "type": "string" } }, "required": ["file"]
            }),
        ),
        tool(
            "xlsx_merged_ranges",
            "List XLSX merged cell ranges without opening or saving the workbook.",
            json!({
                "type": "object", "properties": { "file": { "type": "string" }, "sheet": { "type": "string" } }, "required": ["file"]
            }),
        ),
        tool(
            "xlsx_auto_filters",
            "List XLSX worksheet auto-filter ranges without opening or saving the workbook.",
            json!({
                "type": "object", "properties": { "file": { "type": "string" }, "sheet": { "type": "string" } }, "required": ["file"]
            }),
        ),
        tool(
            "xlsx_protections",
            "List XLSX worksheet protection settings without opening or saving the workbook.",
            json!({
                "type": "object", "properties": { "file": { "type": "string" }, "sheet": { "type": "string" } }, "required": ["file"]
            }),
        ),
        tool(
            "xlsx_create",
            "Create a new XLSX workbook from structured JSON rows and sparse cell edits.",
            xlsx_create_input_schema(),
        ),
        tool(
            "xlsx_edit",
            "Direct OOXML cell/range edits that preserve unrelated workbook XML parts.",
            json!({
                "type": "object", "properties": {
                    "file": { "type": "string" }, "sheet": { "type": "string" },
                    "cell": { "type": "string" }, "range": { "type": "string" },
                    "value": {}, "values": {},
                    "type": { "type": "string", "enum": ["text", "string", "date", "datetime", "int", "integer", "float", "number", "numeric", "bool", "boolean"] },
                    "value_type": { "type": "string", "enum": ["text", "string", "date", "datetime", "int", "integer", "float", "number", "numeric", "bool", "boolean"] },
                    "edits": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "sheet": { "type": "string" },
                                "cell": { "type": "string" },
                                "range": { "type": "string" },
                                "value": {},
                                "values": {},
                                "type": { "type": "string", "enum": ["text", "string", "date", "datetime", "int", "integer", "float", "number", "numeric", "bool", "boolean"] }
                            }
                        }
                    },
                    "allow_protected": { "type": "boolean" }
                }, "required": ["file"]
            }),
        ),
        tool(
            "xlsx_add_sheet",
            "Add a blank XLSX worksheet by editing workbook package parts.",
            json!({
                "type": "object", "properties": { "file": { "type": "string" }, "name": { "type": "string" }, "before": { "type": "string" }, "after": { "type": "string" } }, "required": ["file", "name"]
            }),
        ),
        tool(
            "xlsx_move_sheet",
            "Move an XLSX worksheet in workbook order.",
            json!({
                "type": "object", "properties": { "file": { "type": "string" }, "name": { "type": "string" }, "before": { "type": "string" }, "after": { "type": "string" } }, "required": ["file", "name"]
            }),
        ),
        tool(
            "xlsx_rename_sheet",
            "Rename an XLSX worksheet through Excel COM on Windows so Excel updates dependent references.",
            json!({
                "type": "object", "properties": {
                    "file": { "type": "string" },
                    "sheet": { "type": "string" },
                    "name": { "type": "string" },
                    "keep_open": { "type": "boolean" }
                }, "required": ["file", "sheet", "name"]
            }),
        ),
        tool(
            "xlsx_validate",
            "Open an XLSX workbook through Excel COM on Windows, recalculate, optionally save, and return validation JSON.",
            json!({
                "type": "object", "properties": {
                    "file": { "type": "string" },
                    "full_calc": { "type": "boolean" },
                    "save": { "type": "boolean" },
                    "check_errors": { "type": "boolean" },
                    "max_errors": { "type": "integer" },
                    "keep_open": { "type": "boolean" }
                }, "required": ["file"]
            }),
        ),
        tool(
            "xlsx_insert",
            "Insert rows or columns through Excel COM on Windows and save the workbook.",
            json!({
                "type": "object", "properties": {
                    "file": { "type": "string" },
                    "sheet": { "type": "string" },
                    "axis": { "type": "string", "enum": ["rows", "columns"] },
                    "range": { "type": "string" },
                    "keep_open": { "type": "boolean" }
                }, "required": ["file", "sheet", "axis", "range"]
            }),
        ),
        tool(
            "xlsx_autofit",
            "Autofit rows or columns through Excel COM on Windows and save the workbook.",
            json!({
                "type": "object", "properties": {
                    "file": { "type": "string" },
                    "sheet": { "type": "string" },
                    "range": { "type": "string" },
                    "axis": { "type": "string", "enum": ["all", "rows", "columns"] },
                    "keep_open": { "type": "boolean" }
                }, "required": ["file", "sheet"]
            }),
        ),
        tool(
            "xlsx_format",
            "Apply simple range formatting through Excel COM on Windows and save the workbook.",
            json!({
                "type": "object", "properties": {
                    "file": { "type": "string" },
                    "sheet": { "type": "string" },
                    "range": { "type": "string" },
                    "number_format": { "type": "string" },
                    "bold": { "type": "boolean" },
                    "italic": { "type": "boolean" },
                    "wrap_text": { "type": "boolean" },
                    "fill_color": { "type": "string" },
                    "font_color": { "type": "string" },
                    "keep_open": { "type": "boolean" }
                }, "required": ["file", "sheet", "range"]
            }),
        ),
        tool(
            "xlsx_copy_sheets",
            "Copy worksheets between workbooks through Excel COM on Windows, preserving Excel-native sheet content.",
            json!({
                "type": "object", "properties": {
                    "file": { "type": "string" },
                    "source": { "type": "string" },
                    "sheets": { "type": "array", "items": { "type": "string" } },
                    "after": { "type": "string" },
                    "replace": { "type": "boolean" },
                    "keep_open": { "type": "boolean" }
                }, "required": ["file", "source", "sheets"]
            }),
        ),
        tool(
            "docx_read",
            "Read DOCX body text and tables as markdown.",
            json!({
                "type": "object", "properties": { "file": { "type": "string" } }, "required": ["file"]
            }),
        ),
        tool(
            "docx_replace",
            "DOCX text-node replacements through direct XML or Word COM.",
            json!({
                "type": "object",
                "properties": {
                    "file": { "type": "string" },
                    "find": { "type": "string" },
                    "replace": { "type": "string" },
                    "output": { "type": "string" },
                    "replacements": { "type": "array" },
                    "engine": { "type": "string", "enum": ["direct", "com"] }
                },
                "required": ["file"]
            }),
        ),
        tool(
            "docx_compose",
            "Compose a DOCX from a structured JSON spec. Optionally provide spec.template (path to an existing .docx) to compose into a branded template - the template's styles, headers, footers, embedded media (logo), and page setup are preserved; only the body is replaced with the composed blocks.",
            docx_compose_input_schema(),
        ),
        tool(
            "pptx_read",
            "Read PPTX slide text.",
            json!({
                "type": "object", "properties": { "file": { "type": "string" } }, "required": ["file"]
            }),
        ),
        tool(
            "pptx_notes",
            "Read PPTX speaker notes.",
            json!({
                "type": "object", "properties": { "file": { "type": "string" } }, "required": ["file"]
            }),
        ),
        tool(
            "pptx_build",
            "Build a simple 16:9 PPTX deck from a structured JSON spec.",
            pptx_build_input_schema(),
        ),
        tool(
            "outlook_mail",
            "Read-only Outlook COM mail search. Never sends mail.",
            json!({
                "type": "object", "properties": {
                    "hours": { "type": "integer" }, "count": { "type": "integer" }, "days": { "type": "integer" }, "since": { "type": "string" },
                    "search": { "type": "string" }, "sender": { "type": "string" }, "subject": { "type": "string" }, "to": { "type": "string" },
                    "include_read": { "type": "boolean" }, "full_body": { "type": "boolean" }, "folder": { "type": "string" },
                    "save_attachments": { "type": "string" }, "save_msg": { "type": "string" }
                }
            }),
        ),
        tool(
            "office_doctor",
            "Non-destructive Windows Office COM availability check for Excel, Word, and Outlook.",
            json!({ "type": "object", "properties": {} }),
        ),
        tool(
            "package_list_parts",
            "List raw OOXML zip package parts.",
            json!({
                "type": "object", "properties": { "file": { "type": "string" } }, "required": ["file"]
            }),
        ),
        tool(
            "package_read_part",
            "Read a raw OOXML package part as UTF-8 text or base64 bytes.",
            json!({
                "type": "object", "properties": { "file": { "type": "string" }, "part": { "type": "string" }, "base64": { "type": "boolean" }, "encoding": { "type": "string", "enum": ["text", "base64"] } }, "required": ["file", "part"]
            }),
        ),
        tool(
            "package_write_part",
            "Replace or add a raw OOXML package part. Provide exactly one of `text` (UTF-8 string) or `base64` (base64-encoded bytes). The handler returns an error if neither or both are set.",
            json!({
                "type": "object",
                "properties": {
                    "file": { "type": "string" },
                    "part": { "type": "string" },
                    "text": { "type": "string", "description": "UTF-8 text body. Mutually exclusive with base64." },
                    "base64": { "type": "string", "description": "Base64-encoded byte body. Mutually exclusive with text." }
                },
                "required": ["file", "part"]
            }),
        ),
        tool(
            "package_delete_part",
            "Delete one or more raw OOXML package parts.",
            json!({
                "type": "object", "properties": { "file": { "type": "string" }, "parts": { "type": "array", "items": { "type": "string" } } }, "required": ["file", "parts"]
            }),
        ),
    ]
}

fn xlsx_create_input_schema() -> Value {
    json!({
        "type": "object",
        "properties": {
            "output": { "type": "string" },
            "overwrite": { "type": "boolean" },
            "spec": {
                "type": "object",
                "properties": {
                    "date1904": { "type": "boolean" },
                    "sheets": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "name": { "type": "string" },
                                "rows": {
                                    "type": "array",
                                    "items": { "type": "array", "items": {} }
                                },
                                "cells": {
                                    "type": "array",
                                    "items": {
                                        "type": "object",
                                        "properties": {
                                            "cell": { "type": "string" },
                                            "range": { "type": "string" },
                                            "value": {},
                                            "values": {},
                                            "type": value_type_schema()
                                        }
                                    }
                                }
                            },
                            "required": ["name"]
                        }
                    }
                }
            }
        },
        "required": ["output", "spec"]
    })
}

fn docx_compose_input_schema() -> Value {
    json!({
        "type": "object",
        "properties": {
            "output": { "type": "string" },
            "spec": {
                "type": "object",
                "properties": {
                    "meta": {
                        "type": "object",
                        "properties": {
                            "title": { "type": "string" },
                            "subject": { "type": "string" },
                            "creator": { "type": "string" },
                            "footer_text": { "type": "string" }
                        }
                    },
                    "brand": {
                        "type": "object",
                        "properties": {
                            "font_family": { "type": "string" },
                            "body_color": { "type": "string" },
                            "accent_color": { "type": "string" },
                            "heading_color": { "type": "string" },
                            "muted_color": { "type": "string" },
                            "table_header_fill": { "type": "string" },
                            "alt_row_fill": { "type": "string" }
                        }
                    },
                    "blocks": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "type": {
                                    "type": "string",
                                    "enum": ["title_page", "heading", "body", "body_rich", "bullet", "numbered", "quote", "quote_block", "table", "borderless_table", "divider", "spacer", "page_break"]
                                },
                                "text": { "type": "string" },
                                "level": { "type": "integer" },
                                "segments": { "type": "array" },
                                "headers": { "type": "array", "items": { "type": "string" } },
                                "rows": { "type": "array", "items": { "type": "array", "items": { "type": "string" } } }
                            },
                            "required": ["type"]
                        }
                    },
                    "template": {
                        "type": "string",
                        "description": "Optional path to an existing .docx to compose into. When set, the template's styles, headers, footers, embedded media (logo etc.), and page setup are preserved; only the document body is replaced with the composed blocks. meta.footer_text is ignored in template mode. Footnote blocks are not yet supported in template mode."
                    }
                }
            }
        },
        "required": ["spec", "output"]
    })
}

fn pptx_build_input_schema() -> Value {
    json!({
        "type": "object",
        "properties": {
            "output": { "type": "string" },
            "spec": {
                "type": "object",
                "properties": {
                    "meta": {
                        "type": "object",
                        "properties": {
                            "title": { "type": "string" },
                            "subtitle": { "type": "string" },
                            "author": { "type": "string" },
                            "date": { "type": "string" },
                            "footer_text": { "type": "string" }
                        }
                    },
                    "brand": {
                        "type": "object",
                        "properties": {
                            "font_family": { "type": "string" },
                            "background": { "type": "string" },
                            "text": { "type": "string" },
                            "accent": { "type": "string" },
                            "muted": { "type": "string" }
                        }
                    },
                    "slides": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "layout": {
                                    "type": "string",
                                    "enum": ["title", "content", "two_column", "section", "quote", "table", "image_prompt"]
                                },
                                "title": { "type": "string" },
                                "subtitle": { "type": "string" },
                                "body": { "type": "string" },
                                "bullets": { "type": "array" },
                                "headers": { "type": "array", "items": { "type": "string" } },
                                "rows": { "type": "array", "items": { "type": "array", "items": { "type": "string" } } },
                                "image_prompt": { "type": "string" },
                                "notes": { "type": "string" }
                            }
                        }
                    }
                }
            }
        },
        "required": ["spec", "output"]
    })
}

fn value_type_schema() -> Value {
    json!({
        "type": "string",
        "enum": ["text", "string", "date", "datetime", "int", "integer", "float", "number", "numeric", "bool", "boolean"]
    })
}

fn tool(name: &str, description: &str, input_schema: Value) -> Value {
    json!({ "name": name, "description": description, "inputSchema": input_schema })
}

fn call_tool(params: Value) -> Result<String> {
    let name = params
        .get("name")
        .and_then(Value::as_str)
        .ok_or_else(|| anyhow!("tools/call missing name"))?;
    let args = params
        .get("arguments")
        .cloned()
        .unwrap_or_else(|| json!({}));
    match name {
        "xlsx_read" => {
            let result = read(&ReadArgs {
                file: path_arg(&args, "file")?,
                sheet: string_arg(&args, "sheet"),
                range: string_arg(&args, "range"),
                show_empty_rows: bool_arg(&args, "show_empty_rows"),
                show_empty_cols: bool_arg(&args, "show_empty_cols"),
                json: true,
            })?;
            Ok(serde_json::to_string_pretty(&result)?)
        }
        "xlsx_cells" => {
            let result = cells(&CellsArgs {
                file: path_arg(&args, "file")?,
                sheet: string_arg(&args, "sheet"),
                range: string_arg(&args, "range"),
                include_empty: bool_arg(&args, "include_empty"),
                json: true,
            })?;
            Ok(serde_json::to_string_pretty(&result)?)
        }
        "xlsx_list_sheets" => Ok(serde_json::to_string_pretty(&list_sheets(&path_arg(
            &args, "file",
        )?)?)?),
        "xlsx_search" => {
            let result = search(&SearchArgs {
                file: path_arg(&args, "file")?,
                query: required_string_arg(&args, "query")?,
                sheet: string_arg(&args, "sheet"),
                json: true,
            })?;
            Ok(serde_json::to_string_pretty(&result)?)
        }
        "xlsx_inspect" => {
            let result = inspect(&InspectArgs {
                file: path_arg(&args, "file")?,
                sheet: string_arg(&args, "sheet"),
            })?;
            Ok(serde_json::to_string_pretty(&result)?)
        }
        "xlsx_relationships" => {
            let result = relationships(&RelationshipsArgs {
                file: path_arg(&args, "file")?,
                sheet: string_arg(&args, "sheet"),
                json: true,
            })?;
            Ok(serde_json::to_string_pretty(&result)?)
        }
        "xlsx_formulas" => {
            let result = formulas(&FormulasArgs {
                file: path_arg(&args, "file")?,
                sheet: string_arg(&args, "sheet"),
                json: true,
            })?;
            Ok(serde_json::to_string_pretty(&result)?)
        }
        "xlsx_tables" => {
            let result = tables(&TablesArgs {
                file: path_arg(&args, "file")?,
                sheet: string_arg(&args, "sheet"),
                json: true,
            })?;
            Ok(serde_json::to_string_pretty(&result)?)
        }
        "xlsx_validations" => {
            let result = validations(&ValidationsArgs {
                file: path_arg(&args, "file")?,
                sheet: string_arg(&args, "sheet"),
                json: true,
            })?;
            Ok(serde_json::to_string_pretty(&result)?)
        }
        "xlsx_conditional_formatting" => {
            let result = conditional_formatting(&ConditionalFormattingArgs {
                file: path_arg(&args, "file")?,
                sheet: string_arg(&args, "sheet"),
                json: true,
            })?;
            Ok(serde_json::to_string_pretty(&result)?)
        }
        "xlsx_hyperlinks" => {
            let result = hyperlinks(&HyperlinksArgs {
                file: path_arg(&args, "file")?,
                sheet: string_arg(&args, "sheet"),
                json: true,
            })?;
            Ok(serde_json::to_string_pretty(&result)?)
        }
        "xlsx_comments" => {
            let result = comments(&CommentsArgs {
                file: path_arg(&args, "file")?,
                sheet: string_arg(&args, "sheet"),
                json: true,
            })?;
            Ok(serde_json::to_string_pretty(&result)?)
        }
        "xlsx_defined_names" => {
            let result = defined_names(&DefinedNamesArgs {
                file: path_arg(&args, "file")?,
                name: string_arg(&args, "name"),
                json: true,
            })?;
            Ok(serde_json::to_string_pretty(&result)?)
        }
        "xlsx_merged_ranges" => {
            let result = merged_ranges(&MergedRangesArgs {
                file: path_arg(&args, "file")?,
                sheet: string_arg(&args, "sheet"),
                json: true,
            })?;
            Ok(serde_json::to_string_pretty(&result)?)
        }
        "xlsx_auto_filters" => {
            let result = auto_filters(&AutoFiltersArgs {
                file: path_arg(&args, "file")?,
                sheet: string_arg(&args, "sheet"),
                json: true,
            })?;
            Ok(serde_json::to_string_pretty(&result)?)
        }
        "xlsx_protections" => {
            let result = protections(&ProtectionsArgs {
                file: path_arg(&args, "file")?,
                sheet: string_arg(&args, "sheet"),
                json: true,
            })?;
            Ok(serde_json::to_string_pretty(&result)?)
        }
        "xlsx_create" => {
            let spec: CreateWorkbookSpec =
                serde_json::from_value(required_value_arg(&args, "spec")?)
                    .context("parse xlsx create spec")?;
            let result = create_workbook(
                &spec,
                &path_arg(&args, "output")?,
                bool_arg(&args, "overwrite"),
            )?;
            Ok(serde_json::to_string_pretty(&result)?)
        }
        "xlsx_edit" => {
            let edits = xlsx_edit_payload_arg(&args);
            let result = edit(&EditArgs {
                file: path_arg(&args, "file")?,
                sheet: string_arg(&args, "sheet"),
                cell: string_arg(&args, "cell"),
                range: string_arg(&args, "range"),
                value: None,
                clear: false,
                edits,
                edits_file: None,
                value_type: None,
                allow_protected: bool_arg(&args, "allow_protected"),
                json: true,
            })?;
            Ok(serde_json::to_string_pretty(&result)?)
        }
        "xlsx_add_sheet" => {
            let part = add_sheet(
                &path_arg(&args, "file")?,
                &required_string_arg(&args, "name")?,
                string_arg(&args, "before").as_deref(),
                string_arg(&args, "after").as_deref(),
            )?;
            Ok(serde_json::to_string_pretty(&json!({ "part": part }))?)
        }
        "xlsx_move_sheet" => {
            move_sheet(
                &path_arg(&args, "file")?,
                &required_string_arg(&args, "name")?,
                string_arg(&args, "before").as_deref(),
                string_arg(&args, "after").as_deref(),
            )?;
            Ok("{}".to_string())
        }
        "xlsx_rename_sheet" => crate::wincom::excel_rename_sheet(&RenameSheetArgs {
            file: path_arg(&args, "file")?,
            sheet: required_string_arg(&args, "sheet")?,
            name: required_string_arg(&args, "name")?,
            keep_open: bool_arg(&args, "keep_open"),
        }),
        "xlsx_validate" => crate::wincom::excel_validate(&ValidateArgs {
            file: path_arg(&args, "file")?,
            full_calc: bool_arg(&args, "full_calc"),
            save: bool_arg(&args, "save"),
            check_errors: bool_arg(&args, "check_errors"),
            max_errors: usize_arg(&args, "max_errors")?.unwrap_or(50),
            keep_open: bool_arg(&args, "keep_open"),
        }),
        "xlsx_insert" => crate::wincom::excel_insert(&InsertArgs {
            file: path_arg(&args, "file")?,
            sheet: required_string_arg(&args, "sheet")?,
            axis: insert_axis_arg(&args)?,
            range: required_string_arg(&args, "range")?,
            keep_open: bool_arg(&args, "keep_open"),
        }),
        "xlsx_autofit" => crate::wincom::excel_autofit(&AutoFitArgs {
            file: path_arg(&args, "file")?,
            sheet: required_string_arg(&args, "sheet")?,
            range: string_arg(&args, "range"),
            axis: autofit_axis_arg(&args)?,
            keep_open: bool_arg(&args, "keep_open"),
        }),
        "xlsx_format" => crate::wincom::excel_format(&FormatArgs {
            file: path_arg(&args, "file")?,
            sheet: required_string_arg(&args, "sheet")?,
            range: required_string_arg(&args, "range")?,
            number_format: string_arg(&args, "number_format"),
            bold: optional_bool_arg(&args, "bold")?,
            italic: optional_bool_arg(&args, "italic")?,
            wrap_text: optional_bool_arg(&args, "wrap_text")?,
            fill_color: string_arg(&args, "fill_color"),
            font_color: string_arg(&args, "font_color"),
            keep_open: bool_arg(&args, "keep_open"),
        }),
        "xlsx_copy_sheets" => crate::wincom::excel_copy_sheets(&CopySheetsArgs {
            file: path_arg(&args, "file")?,
            source: path_arg(&args, "source")?,
            sheets: string_list_arg(&args, "sheets")?,
            after: string_arg(&args, "after"),
            replace: bool_arg(&args, "replace"),
            keep_open: bool_arg(&args, "keep_open"),
        }),
        "docx_read" => Ok(crate::docx::read_docx_markdown(&path_arg(&args, "file")?)?),
        "docx_replace" => {
            let replacements = args
                .get("replacements")
                .map(|v| serde_json::from_value::<Vec<Replacement>>(v.clone()))
                .transpose()?;
            let (find, replace_with) = if replacements.is_none() {
                (string_arg(&args, "find"), string_arg(&args, "replace"))
            } else {
                (None, None)
            };
            let result = crate::docx::replace(&DocxReplaceArgs {
                file: path_arg(&args, "file")?,
                find,
                replace: replace_with,
                replacements: replacements
                    .map(|items| serde_json::to_string(&items))
                    .transpose()?,
                replacements_file: None,
                output: string_arg(&args, "output").map(PathBuf::from),
                engine: docx_replace_engine_arg(&args)?,
            })?;
            Ok(serde_json::to_string_pretty(&result)?)
        }
        "docx_compose" => {
            let spec: ComposeSpec = serde_json::from_value(required_value_arg(&args, "spec")?)
                .context("parse docx compose spec")?;
            let output = path_arg(&args, "output")?;
            compose_docx(&spec, &output)?;
            Ok(serde_json::to_string_pretty(&json!({ "file": output }))?)
        }
        "pptx_read" => Ok(serde_json::to_string_pretty(&read_slides(&path_arg(
            &args, "file",
        )?)?)?),
        "pptx_notes" => Ok(serde_json::to_string_pretty(&read_notes(&path_arg(
            &args, "file",
        )?)?)?),
        "pptx_build" => {
            let spec: DeckSpec = serde_json::from_value(required_value_arg(&args, "spec")?)
                .context("parse pptx build spec")?;
            let output = path_arg(&args, "output")?;
            build_deck(&spec, &output)?;
            Ok(serde_json::to_string_pretty(&json!({ "file": output }))?)
        }
        "outlook_mail" => {
            let outlook = OutlookArgs {
                hours: int_arg(&args, "hours").unwrap_or(24),
                count: int_arg(&args, "count").unwrap_or(20) as usize,
                days: int_arg(&args, "days"),
                since: string_arg(&args, "since"),
                search: string_arg(&args, "search"),
                sender: string_arg(&args, "sender"),
                subject: string_arg(&args, "subject"),
                to: string_arg(&args, "to"),
                include_read: bool_arg(&args, "include_read"),
                full_body: bool_arg(&args, "full_body"),
                folder: string_arg(&args, "folder").unwrap_or_else(|| "inbox".to_string()),
                save_attachments: string_arg(&args, "save_attachments").map(PathBuf::from),
                save_msg: string_arg(&args, "save_msg").map(PathBuf::from),
            };
            outlook.fetch_json()
        }
        "office_doctor" => crate::wincom::office_doctor(),
        "package_list_parts" => Ok(serde_json::to_string_pretty(&list_parts(path_arg(
            &args, "file",
        )?)?)?),
        "package_read_part" => {
            let data = read_part(
                path_arg(&args, "file")?,
                &required_string_arg(&args, "part")?,
            )?;
            if bool_arg(&args, "base64")
                || string_arg(&args, "encoding").as_deref() == Some("base64")
            {
                Ok(BASE64_STANDARD.encode(data))
            } else {
                Ok(String::from_utf8_lossy(&data).to_string())
            }
        }
        "package_write_part" => {
            let mut updates = PartMap::new();
            let text = string_arg(&args, "text");
            let base64 = string_arg(&args, "base64");
            if text.is_some() == base64.is_some() {
                bail!("provide exactly one of text or base64");
            }
            let data = if let Some(text) = text {
                text.into_bytes()
            } else {
                BASE64_STANDARD
                    .decode(base64.expect("checked above"))
                    .context("decode base64 package part")?
            };
            updates.insert(required_string_arg(&args, "part")?, data);
            rewrite_parts(path_arg(&args, "file")?, &updates, Vec::<String>::new())?;
            Ok("{}".to_string())
        }
        "package_delete_part" => {
            let parts = string_list_arg(&args, "parts")?;
            if parts.is_empty() {
                bail!("provide at least one package part to delete");
            }
            rewrite_parts(path_arg(&args, "file")?, &PartMap::new(), parts)?;
            Ok("{}".to_string())
        }
        _ => bail!("unknown tool: {name}"),
    }
}

fn read_message(reader: &mut impl BufRead) -> Result<Option<Value>> {
    let mut headers = Vec::new();
    loop {
        let mut line = String::new();
        let bytes = reader.read_line(&mut line)?;
        if bytes == 0 {
            return Ok(None);
        }
        if line.starts_with('{') {
            return Ok(Some(serde_json::from_str(line.trim())?));
        }
        let trimmed = line.trim_end_matches(['\r', '\n']);
        if trimmed.is_empty() {
            break;
        }
        headers.push(trimmed.to_string());
    }
    let content_length = headers
        .iter()
        .find_map(|header| {
            let (name, value) = header.split_once(':')?;
            name.eq_ignore_ascii_case("content-length")
                .then(|| value.trim().parse::<usize>().ok())
                .flatten()
        })
        .ok_or_else(|| anyhow!("MCP frame missing Content-Length header"))?;
    let mut buf = vec![0u8; content_length];
    reader.read_exact(&mut buf)?;
    Ok(Some(serde_json::from_slice(&buf)?))
}

fn write_message(writer: &mut impl Write, message: &Value) -> Result<()> {
    let data = serde_json::to_vec(message)?;
    writer.write_all(&data)?;
    writer.write_all(b"\n")?;
    writer.flush()?;
    Ok(())
}

fn path_arg(args: &Value, name: &str) -> Result<PathBuf> {
    Ok(PathBuf::from(required_string_arg(args, name)?))
}

fn required_value_arg(args: &Value, name: &str) -> Result<Value> {
    args.get(name)
        .cloned()
        .ok_or_else(|| anyhow!("missing required argument: {name}"))
}

fn required_string_arg(args: &Value, name: &str) -> Result<String> {
    string_arg(args, name).ok_or_else(|| anyhow!("missing required argument: {name}"))
}

fn string_arg(args: &Value, name: &str) -> Option<String> {
    args.get(name)
        .and_then(Value::as_str)
        .map(ToString::to_string)
}

fn bool_arg(args: &Value, name: &str) -> bool {
    args.get(name).and_then(Value::as_bool).unwrap_or(false)
}

fn optional_bool_arg(args: &Value, name: &str) -> Result<Option<bool>> {
    let Some(value) = args.get(name) else {
        return Ok(None);
    };
    value
        .as_bool()
        .map(Some)
        .ok_or_else(|| anyhow!("{name} must be a boolean"))
}

fn int_arg(args: &Value, name: &str) -> Option<i64> {
    args.get(name).and_then(Value::as_i64)
}

fn usize_arg(args: &Value, name: &str) -> Result<Option<usize>> {
    let Some(value) = int_arg(args, name) else {
        return Ok(None);
    };
    if value < 0 {
        bail!("{name} cannot be negative");
    }
    Ok(Some(value as usize))
}

fn string_list_arg(args: &Value, name: &str) -> Result<Vec<String>> {
    let Some(value) = args.get(name) else {
        return Ok(Vec::new());
    };
    if let Some(text) = value.as_str() {
        return Ok(vec![text.to_string()]);
    }
    let Some(items) = value.as_array() else {
        bail!("{name} must be a string or array of strings");
    };
    items
        .iter()
        .map(|item| {
            item.as_str()
                .map(ToString::to_string)
                .ok_or_else(|| anyhow!("{name} must contain only strings"))
        })
        .collect()
}

fn xlsx_edit_payload_arg(args: &Value) -> Option<String> {
    if let Some(edits) = args.get("edits") {
        return Some(edits.to_string());
    }
    let has_target = args.get("cell").is_some() || args.get("range").is_some();
    let has_value = args.get("value").is_some() || args.get("values").is_some();
    if !has_target || !has_value {
        return None;
    }
    let mut edit = serde_json::Map::new();
    for key in ["sheet", "cell", "range"] {
        if let Some(value) = string_arg(args, key) {
            edit.insert(key.to_string(), Value::String(value));
        }
    }
    if let Some(value) = args.get("value") {
        edit.insert("value".to_string(), value.clone());
    }
    if let Some(values) = args.get("values") {
        edit.insert("values".to_string(), values.clone());
    }
    if let Some(value_type) = string_arg(args, "type").or_else(|| string_arg(args, "value_type")) {
        edit.insert("type".to_string(), Value::String(value_type));
    }
    Some(Value::Array(vec![Value::Object(edit)]).to_string())
}

fn insert_axis_arg(args: &Value) -> Result<InsertAxis> {
    match required_string_arg(args, "axis")?
        .to_ascii_lowercase()
        .as_str()
    {
        "rows" | "row" => Ok(InsertAxis::Rows),
        "columns" | "column" | "cols" | "col" => Ok(InsertAxis::Columns),
        other => bail!("axis must be rows or columns, got {other}"),
    }
}

fn autofit_axis_arg(args: &Value) -> Result<AutoFitAxis> {
    match string_arg(args, "axis")
        .unwrap_or_else(|| "all".to_string())
        .to_ascii_lowercase()
        .as_str()
    {
        "all" => Ok(AutoFitAxis::All),
        "rows" | "row" => Ok(AutoFitAxis::Rows),
        "columns" | "column" | "cols" | "col" => Ok(AutoFitAxis::Columns),
        other => bail!("axis must be all, rows, or columns, got {other}"),
    }
}

fn docx_replace_engine_arg(args: &Value) -> Result<DocxReplaceEngine> {
    match string_arg(args, "engine")
        .unwrap_or_else(|| "direct".to_string())
        .to_ascii_lowercase()
        .as_str()
    {
        "direct" => Ok(DocxReplaceEngine::Direct),
        "com" => Ok(DocxReplaceEngine::Com),
        other => bail!("engine must be direct or com, got {other}"),
    }
}
