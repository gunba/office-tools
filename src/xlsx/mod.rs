use crate::ooxml::{
    PartMap, child, child_mut, children, escaped, parse_xml, read_part,
    relationship_target_for_part, resolve_relationship_target, rewrite_parts, write_new_package,
    write_xml,
};
use anyhow::{Context, Result, anyhow, bail};
use calamine::{Data, Reader, open_workbook_auto};
use chrono::{NaiveDate, NaiveDateTime};
use clap::{Args, Subcommand, ValueEnum};
use regex::Regex;
use serde::{Deserialize, Serialize};
use serde_json::Value;
use std::collections::{BTreeMap, BTreeSet};
use std::fs;
use std::path::{Path, PathBuf};
use xmltree::{Element, XMLNode};

const WORKBOOK: &str = "xl/workbook.xml";
const WORKBOOK_RELS: &str = "xl/_rels/workbook.xml.rels";
const CONTENT_TYPES: &str = "[Content_Types].xml";
const CALC_CHAIN_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";
const WORKSHEET_REL_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const WORKSHEET_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const WORKBOOK_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";

#[derive(Debug, Args)]
pub struct XlsxArgs {
    #[command(subcommand)]
    command: XlsxCommand,
}

#[derive(Debug, Subcommand)]
pub enum XlsxCommand {
    /// Read a worksheet range as markdown or JSON.
    Read(ReadArgs),
    /// List cells with addresses as JSON-friendly records.
    Cells(CellsArgs),
    /// List workbook sheets in order.
    ListSheets(ListSheetsArgs),
    /// Search visible cell values.
    Search(SearchArgs),
    /// Inspect workbook OOXML structure, formulas, validation, and formatting metadata.
    Inspect(InspectArgs),
    /// List workbook and worksheet relationships without opening or saving the workbook.
    Relationships(RelationshipsArgs),
    /// List worksheet formulas without opening or saving the workbook.
    Formulas(FormulasArgs),
    /// List worksheet table names and ranges without opening or saving the workbook.
    Tables(TablesArgs),
    /// List worksheet data validation rules without opening or saving the workbook.
    Validations(ValidationsArgs),
    /// List worksheet conditional formatting rules without opening or saving the workbook.
    ConditionalFormatting(ConditionalFormattingArgs),
    /// List worksheet hyperlinks without opening or saving the workbook.
    Hyperlinks(HyperlinksArgs),
    /// List worksheet comments/notes without opening or saving the workbook.
    Comments(CommentsArgs),
    /// List workbook defined names without opening or saving the workbook.
    DefinedNames(DefinedNamesArgs),
    /// List worksheet merged ranges without opening or saving the workbook.
    MergedRanges(MergedRangesArgs),
    /// List worksheet auto-filter ranges without opening or saving the workbook.
    AutoFilters(AutoFiltersArgs),
    /// List worksheet protection settings without opening or saving the workbook.
    Protections(ProtectionsArgs),
    /// Create a new XLSX workbook from a JSON spec.
    Create(CreateArgs),
    /// Edit cells through the direct OOXML writer.
    Edit(EditArgs),
    /// Add a blank worksheet without rewriting unrelated parts.
    AddSheet(AddSheetArgs),
    /// Move a worksheet in workbook order.
    MoveSheet(MoveSheetArgs),
    /// Rename a worksheet through Excel COM on Windows.
    RenameSheet(RenameSheetArgs),
    /// Run Excel COM validation/recalculation on Windows.
    Validate(ValidateArgs),
    /// Insert rows or columns through Excel COM on Windows.
    Insert(InsertArgs),
    /// Autofit rows or columns through Excel COM on Windows.
    #[command(name = "autofit", alias = "auto-fit")]
    AutoFit(AutoFitArgs),
    /// Apply simple range formatting through Excel COM on Windows.
    Format(FormatArgs),
    /// Copy sheets through Excel COM on Windows.
    CopySheets(CopySheetsArgs),
}

#[derive(Debug, Args)]
pub struct ReadArgs {
    pub file: PathBuf,
    #[arg(long)]
    pub sheet: Option<String>,
    #[arg(long)]
    pub range: Option<String>,
    #[arg(long)]
    pub show_empty_rows: bool,
    #[arg(long)]
    pub show_empty_cols: bool,
    #[arg(long)]
    pub json: bool,
}

#[derive(Debug, Args)]
pub struct CellsArgs {
    pub file: PathBuf,
    #[arg(long)]
    pub sheet: Option<String>,
    #[arg(long)]
    pub range: Option<String>,
    #[arg(long)]
    pub include_empty: bool,
    #[arg(long)]
    pub json: bool,
}

#[derive(Debug, Args)]
pub struct ListSheetsArgs {
    pub file: PathBuf,
    #[arg(long)]
    pub json: bool,
}

#[derive(Debug, Args)]
pub struct SearchArgs {
    pub file: PathBuf,
    #[arg(short, long)]
    pub query: String,
    #[arg(long)]
    pub sheet: Option<String>,
    #[arg(long)]
    pub json: bool,
}

#[derive(Debug, Args)]
pub struct InspectArgs {
    pub file: PathBuf,
    #[arg(long)]
    pub sheet: Option<String>,
}

#[derive(Debug, Args)]
pub struct RelationshipsArgs {
    pub file: PathBuf,
    #[arg(long)]
    pub sheet: Option<String>,
    #[arg(long)]
    pub json: bool,
}

#[derive(Debug, Args)]
pub struct FormulasArgs {
    pub file: PathBuf,
    #[arg(long)]
    pub sheet: Option<String>,
    #[arg(long)]
    pub json: bool,
}

#[derive(Debug, Args)]
pub struct TablesArgs {
    pub file: PathBuf,
    #[arg(long)]
    pub sheet: Option<String>,
    #[arg(long)]
    pub json: bool,
}

#[derive(Debug, Args)]
pub struct ValidationsArgs {
    pub file: PathBuf,
    #[arg(long)]
    pub sheet: Option<String>,
    #[arg(long)]
    pub json: bool,
}

#[derive(Debug, Args)]
pub struct ConditionalFormattingArgs {
    pub file: PathBuf,
    #[arg(long)]
    pub sheet: Option<String>,
    #[arg(long)]
    pub json: bool,
}

#[derive(Debug, Args)]
pub struct HyperlinksArgs {
    pub file: PathBuf,
    #[arg(long)]
    pub sheet: Option<String>,
    #[arg(long)]
    pub json: bool,
}

#[derive(Debug, Args)]
pub struct CommentsArgs {
    pub file: PathBuf,
    #[arg(long)]
    pub sheet: Option<String>,
    #[arg(long)]
    pub json: bool,
}

#[derive(Debug, Args)]
pub struct DefinedNamesArgs {
    pub file: PathBuf,
    #[arg(long)]
    pub name: Option<String>,
    #[arg(long)]
    pub json: bool,
}

#[derive(Debug, Args)]
pub struct MergedRangesArgs {
    pub file: PathBuf,
    #[arg(long)]
    pub sheet: Option<String>,
    #[arg(long)]
    pub json: bool,
}

#[derive(Debug, Args)]
pub struct AutoFiltersArgs {
    pub file: PathBuf,
    #[arg(long)]
    pub sheet: Option<String>,
    #[arg(long)]
    pub json: bool,
}

#[derive(Debug, Args)]
pub struct ProtectionsArgs {
    pub file: PathBuf,
    #[arg(long)]
    pub sheet: Option<String>,
    #[arg(long)]
    pub json: bool,
}

#[derive(Debug, Args)]
pub struct CreateArgs {
    pub spec: PathBuf,
    #[arg(short, long)]
    pub output: PathBuf,
    #[arg(long)]
    pub overwrite: bool,
    #[arg(long)]
    pub json: bool,
}

#[derive(Debug, Args)]
pub struct EditArgs {
    pub file: PathBuf,
    #[arg(long)]
    pub sheet: Option<String>,
    #[arg(long)]
    pub cell: Option<String>,
    #[arg(long)]
    pub range: Option<String>,
    #[arg(long)]
    pub value: Option<String>,
    #[arg(long)]
    pub clear: bool,
    #[arg(long)]
    pub edits: Option<String>,
    #[arg(long)]
    pub edits_file: Option<PathBuf>,
    #[arg(long = "value-type")]
    pub value_type: Option<ValueType>,
    #[arg(long)]
    pub allow_protected: bool,
    #[arg(long)]
    pub json: bool,
}

#[derive(Debug, Args)]
pub struct AddSheetArgs {
    pub file: PathBuf,
    pub name: String,
    #[arg(long)]
    pub before: Option<String>,
    #[arg(long)]
    pub after: Option<String>,
    #[arg(long)]
    pub json: bool,
}

#[derive(Debug, Args)]
pub struct MoveSheetArgs {
    pub file: PathBuf,
    pub name: String,
    #[arg(long)]
    pub before: Option<String>,
    #[arg(long)]
    pub after: Option<String>,
    #[arg(long)]
    pub json: bool,
}

#[derive(Debug, Args)]
pub struct RenameSheetArgs {
    pub file: PathBuf,
    #[arg(long)]
    pub sheet: String,
    #[arg(long)]
    pub name: String,
    #[arg(long)]
    pub keep_open: bool,
}

#[derive(Debug, Args)]
pub struct ValidateArgs {
    pub file: PathBuf,
    #[arg(long)]
    pub full_calc: bool,
    #[arg(long)]
    pub save: bool,
    #[arg(long)]
    pub check_errors: bool,
    #[arg(long, default_value_t = 50)]
    pub max_errors: usize,
    #[arg(long)]
    pub keep_open: bool,
}

#[derive(Debug, Args)]
pub struct InsertArgs {
    pub file: PathBuf,
    #[arg(long)]
    pub sheet: String,
    #[arg(long)]
    pub axis: InsertAxis,
    #[arg(long)]
    pub range: String,
    #[arg(long)]
    pub keep_open: bool,
}

#[derive(Debug, Clone, Copy, ValueEnum)]
pub enum InsertAxis {
    Rows,
    Columns,
}

#[derive(Debug, Args)]
pub struct AutoFitArgs {
    pub file: PathBuf,
    #[arg(long)]
    pub sheet: String,
    #[arg(long)]
    pub range: Option<String>,
    #[arg(long, value_enum, default_value_t = AutoFitAxis::All)]
    pub axis: AutoFitAxis,
    #[arg(long)]
    pub keep_open: bool,
}

#[derive(Debug, Clone, Copy, ValueEnum)]
pub enum AutoFitAxis {
    All,
    Rows,
    Columns,
}

#[derive(Debug, Args)]
pub struct FormatArgs {
    pub file: PathBuf,
    #[arg(long)]
    pub sheet: String,
    #[arg(long)]
    pub range: String,
    #[arg(long)]
    pub number_format: Option<String>,
    #[arg(long)]
    pub bold: Option<bool>,
    #[arg(long)]
    pub italic: Option<bool>,
    #[arg(long)]
    pub wrap_text: Option<bool>,
    #[arg(long)]
    pub fill_color: Option<String>,
    #[arg(long)]
    pub font_color: Option<String>,
    #[arg(long)]
    pub keep_open: bool,
}

#[derive(Debug, Args)]
pub struct CopySheetsArgs {
    pub file: PathBuf,
    #[arg(long)]
    pub source: PathBuf,
    #[arg(long, required = true)]
    pub sheets: Vec<String>,
    #[arg(long)]
    pub after: Option<String>,
    #[arg(long)]
    pub replace: bool,
    #[arg(long)]
    pub keep_open: bool,
}

#[derive(Debug, Clone, Copy, ValueEnum, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum ValueType {
    Text,
    String,
    Date,
    Datetime,
    Int,
    Integer,
    Float,
    Number,
    Numeric,
    Bool,
    Boolean,
}

#[derive(Debug, Clone, Deserialize)]
pub struct CellEdit {
    pub sheet: Option<String>,
    pub cell: Option<String>,
    #[serde(rename = "range")]
    pub range_ref: Option<String>,
    pub value: Option<Value>,
    pub values: Option<Value>,
    #[serde(rename = "type")]
    pub value_type: Option<ValueType>,
}

#[derive(Debug, Deserialize)]
pub struct CreateWorkbookSpec {
    #[serde(default)]
    pub date1904: bool,
    #[serde(default)]
    pub sheets: Vec<CreateSheetSpec>,
}

#[derive(Debug, Deserialize)]
pub struct CreateSheetSpec {
    pub name: String,
    #[serde(default)]
    pub rows: Vec<Vec<Value>>,
    #[serde(default)]
    pub cells: Vec<CellEdit>,
}

#[derive(Debug, Serialize)]
pub struct ReadResult {
    pub sheet: String,
    pub dimensions: String,
    pub rows: Vec<Vec<String>>,
}

#[derive(Debug, Serialize)]
pub struct SheetInfo {
    pub index: usize,
    pub name: String,
    pub sheet_id: Option<String>,
    pub state: Option<String>,
}

#[derive(Debug, Serialize)]
pub struct SearchHit {
    pub sheet: String,
    pub cell: String,
    pub value: String,
}

#[derive(Debug, Serialize)]
pub struct CellRecord {
    pub sheet: String,
    pub cell: String,
    pub value: String,
}

#[derive(Debug, Serialize)]
pub struct WorkbookInspection {
    pub file: PathBuf,
    pub date1904: bool,
    pub sheets: Vec<SheetInspection>,
    pub defined_names: Vec<DefinedNameInspection>,
    pub calc_properties: BTreeMap<String, String>,
    pub workbook_relationships: Vec<RelationshipInspection>,
}

#[derive(Debug, Serialize)]
pub struct SheetInspection {
    pub name: String,
    pub part: String,
    pub dimension: Option<String>,
    pub rows: usize,
    pub cells: usize,
    pub formulas: usize,
    pub merged_ranges: Vec<String>,
    pub conditional_formatting_ranges: Vec<String>,
    pub data_validation_ranges: Vec<String>,
    pub hyperlinks: usize,
    pub has_auto_filter: bool,
    pub has_sheet_protection: bool,
    pub table_relationships: Vec<String>,
}

#[derive(Debug, Serialize)]
pub struct DefinedNameInspection {
    pub name: String,
    pub local_sheet_id: Option<String>,
    pub sheet: Option<String>,
    pub value: String,
}

#[derive(Debug, Serialize)]
pub struct MergedRangeInfo {
    pub sheet: String,
    pub range: String,
}

#[derive(Debug, Serialize)]
pub struct AutoFilterInfo {
    pub sheet: String,
    pub range: Option<String>,
}

#[derive(Debug, Serialize)]
pub struct ProtectionInfo {
    pub sheet: String,
    pub attributes: BTreeMap<String, String>,
}

#[derive(Debug, Serialize)]
pub struct RelationshipInspection {
    pub id: String,
    pub relationship_type: String,
    pub target: String,
    pub target_mode: Option<String>,
}

#[derive(Debug, Serialize)]
pub struct RelationshipInfo {
    pub scope: String,
    pub sheet: Option<String>,
    pub part: String,
    pub id: String,
    pub relationship_type: String,
    pub target: String,
    pub target_mode: Option<String>,
}

#[derive(Debug, Serialize)]
pub struct FormulaInfo {
    pub sheet: String,
    pub cell: String,
    pub formula: String,
}

#[derive(Debug, Serialize)]
pub struct TableInfo {
    pub sheet: String,
    pub part: String,
    pub name: Option<String>,
    pub display_name: Option<String>,
    pub range: Option<String>,
}

#[derive(Debug, Serialize)]
pub struct DataValidationInfo {
    pub sheet: String,
    pub range: Option<String>,
    pub validation_type: Option<String>,
    pub operator: Option<String>,
    pub allow_blank: Option<String>,
    pub formula1: Option<String>,
    pub formula2: Option<String>,
}

#[derive(Debug, Serialize)]
pub struct ConditionalFormattingInfo {
    pub sheet: String,
    pub range: Option<String>,
    pub rule_type: Option<String>,
    pub priority: Option<String>,
    pub operator: Option<String>,
    pub formulas: Vec<String>,
}

#[derive(Debug, Serialize)]
pub struct HyperlinkInfo {
    pub sheet: String,
    pub reference: Option<String>,
    pub target: Option<String>,
    pub location: Option<String>,
    pub display: Option<String>,
    pub tooltip: Option<String>,
}

#[derive(Debug, Serialize)]
pub struct CommentInfo {
    pub sheet: String,
    pub cell: Option<String>,
    pub author: Option<String>,
    pub text: String,
    pub part: String,
}

#[derive(Debug, Serialize)]
pub struct EditResult {
    pub file: PathBuf,
    pub target_count: usize,
    pub sheets: BTreeMap<String, usize>,
    pub formula_recalc_marked: bool,
}

#[derive(Debug, Serialize)]
pub struct CreateResult {
    pub file: PathBuf,
    pub sheets: Vec<SheetInfo>,
    pub cells_written: usize,
    pub formula_recalc_marked: bool,
}

impl XlsxArgs {
    pub fn run(self) -> Result<()> {
        match self.command {
            XlsxCommand::Read(args) => {
                let result = read(&args)?;
                if args.json {
                    print_json(&result)?;
                } else {
                    print_read_result(&result);
                }
                Ok(())
            }
            XlsxCommand::Cells(args) => {
                let result = cells(&args)?;
                if args.json {
                    print_json(&result)?;
                } else if result.is_empty() {
                    println!("  No cells found.");
                } else {
                    for item in result {
                        println!("  [{}] {}: {}", item.sheet, item.cell, item.value);
                    }
                }
                Ok(())
            }
            XlsxCommand::ListSheets(args) => {
                let sheets = list_sheets(&args.file)?;
                if args.json {
                    print_json(&sheets)?;
                } else {
                    for sheet in sheets {
                        println!("  {}. {}", sheet.index, sheet.name);
                    }
                }
                Ok(())
            }
            XlsxCommand::Search(args) => {
                let hits = search(&args)?;
                if args.json {
                    print_json(&hits)?;
                } else if hits.is_empty() {
                    println!("  No matches for '{}'", args.query);
                } else {
                    for hit in &hits {
                        println!("  [{}] {}: {}", hit.sheet, hit.cell, hit.value);
                    }
                    println!("\n  {} match(es) found.", hits.len());
                }
                Ok(())
            }
            XlsxCommand::Inspect(args) => {
                let result = inspect(&args)?;
                print_json(&result)?;
                Ok(())
            }
            XlsxCommand::Relationships(args) => {
                let result = relationships(&args)?;
                if args.json {
                    print_json(&result)?;
                } else if result.is_empty() {
                    println!("  No relationships found.");
                } else {
                    for item in result {
                        let sheet = item
                            .sheet
                            .as_deref()
                            .map(|name| format!(" ({name})"))
                            .unwrap_or_default();
                        println!(
                            "  [{}{sheet}] {} -> {}",
                            item.scope, item.relationship_type, item.target
                        );
                    }
                }
                Ok(())
            }
            XlsxCommand::Formulas(args) => {
                let result = formulas(&args)?;
                if args.json {
                    print_json(&result)?;
                } else if result.is_empty() {
                    println!("  No formulas found.");
                } else {
                    for item in result {
                        println!("  [{}] {}: ={}", item.sheet, item.cell, item.formula);
                    }
                }
                Ok(())
            }
            XlsxCommand::Tables(args) => {
                let result = tables(&args)?;
                if args.json {
                    print_json(&result)?;
                } else if result.is_empty() {
                    println!("  No tables found.");
                } else {
                    for item in result {
                        let name = item
                            .display_name
                            .as_deref()
                            .or(item.name.as_deref())
                            .unwrap_or("(unnamed)");
                        let range = item.range.as_deref().unwrap_or("(unknown range)");
                        println!("  [{}] {name}: {range}", item.sheet);
                    }
                }
                Ok(())
            }
            XlsxCommand::Validations(args) => {
                let result = validations(&args)?;
                if args.json {
                    print_json(&result)?;
                } else if result.is_empty() {
                    println!("  No data validations found.");
                } else {
                    for item in result {
                        let range = item.range.as_deref().unwrap_or("(unknown range)");
                        let validation_type =
                            item.validation_type.as_deref().unwrap_or("(unknown type)");
                        println!("  [{}] {range}: {validation_type}", item.sheet);
                    }
                }
                Ok(())
            }
            XlsxCommand::ConditionalFormatting(args) => {
                let result = conditional_formatting(&args)?;
                if args.json {
                    print_json(&result)?;
                } else if result.is_empty() {
                    println!("  No conditional formatting rules found.");
                } else {
                    for item in result {
                        let range = item.range.as_deref().unwrap_or("(unknown range)");
                        let rule_type = item.rule_type.as_deref().unwrap_or("(unknown type)");
                        println!("  [{}] {range}: {rule_type}", item.sheet);
                    }
                }
                Ok(())
            }
            XlsxCommand::Hyperlinks(args) => {
                let result = hyperlinks(&args)?;
                if args.json {
                    print_json(&result)?;
                } else if result.is_empty() {
                    println!("  No hyperlinks found.");
                } else {
                    for item in result {
                        let reference = item.reference.as_deref().unwrap_or("(unknown ref)");
                        let target = item
                            .target
                            .as_deref()
                            .or(item.location.as_deref())
                            .unwrap_or("(unknown target)");
                        println!("  [{}] {reference}: {target}", item.sheet);
                    }
                }
                Ok(())
            }
            XlsxCommand::Comments(args) => {
                let result = comments(&args)?;
                if args.json {
                    print_json(&result)?;
                } else if result.is_empty() {
                    println!("  No comments found.");
                } else {
                    for item in result {
                        let cell = item.cell.as_deref().unwrap_or("(unknown cell)");
                        let author = item.author.as_deref().unwrap_or("(unknown author)");
                        println!("  [{}] {cell} ({author}): {}", item.sheet, item.text);
                    }
                }
                Ok(())
            }
            XlsxCommand::DefinedNames(args) => {
                let result = defined_names(&args)?;
                if args.json {
                    print_json(&result)?;
                } else if result.is_empty() {
                    println!("  No defined names found.");
                } else {
                    for item in result {
                        let scope = item.sheet.as_deref().unwrap_or("workbook");
                        println!("  [{scope}] {}: {}", item.name, item.value);
                    }
                }
                Ok(())
            }
            XlsxCommand::MergedRanges(args) => {
                let result = merged_ranges(&args)?;
                if args.json {
                    print_json(&result)?;
                } else if result.is_empty() {
                    println!("  No merged ranges found.");
                } else {
                    for item in result {
                        println!("  [{}] {}", item.sheet, item.range);
                    }
                }
                Ok(())
            }
            XlsxCommand::AutoFilters(args) => {
                let result = auto_filters(&args)?;
                if args.json {
                    print_json(&result)?;
                } else if result.is_empty() {
                    println!("  No auto-filters found.");
                } else {
                    for item in result {
                        let range = item.range.as_deref().unwrap_or("(unknown range)");
                        println!("  [{}] {range}", item.sheet);
                    }
                }
                Ok(())
            }
            XlsxCommand::Protections(args) => {
                let result = protections(&args)?;
                if args.json {
                    print_json(&result)?;
                } else if result.is_empty() {
                    println!("  No sheet protection found.");
                } else {
                    for item in result {
                        println!("  [{}] {:?}", item.sheet, item.attributes);
                    }
                }
                Ok(())
            }
            XlsxCommand::Create(args) => {
                let result = create(&args)?;
                if args.json {
                    print_json(&result)?;
                } else {
                    println!("Saved via direct OOXML engine: {}", result.file.display());
                    println!("Sheets: {}", result.sheets.len());
                    println!("Cells written: {}", result.cells_written);
                    if result.formula_recalc_marked {
                        println!("Formula cells marked workbook for recalculation on Excel open.");
                    }
                }
                Ok(())
            }
            XlsxCommand::Edit(args) => {
                let result = edit(&args)?;
                if args.json {
                    print_json(&result)?;
                } else {
                    for (sheet, count) in &result.sheets {
                        println!("  Edited {count} target(s) on {sheet}");
                    }
                    println!("\nSaved via direct OOXML engine: {}", result.file.display());
                    if result.formula_recalc_marked {
                        println!(
                            "Formula edits marked workbook for recalculation on next Excel open."
                        );
                    }
                    println!("Total target range(s): {}", result.target_count);
                }
                Ok(())
            }
            XlsxCommand::AddSheet(args) => {
                let part = add_sheet(
                    &args.file,
                    &args.name,
                    args.before.as_deref(),
                    args.after.as_deref(),
                )?;
                if args.json {
                    print_json(
                        &serde_json::json!({ "file": args.file, "name": args.name, "part": part }),
                    )?;
                } else {
                    println!("  Added sheet {} ({part})", args.name);
                    println!("\nSaved via direct OOXML engine: {}", args.file.display());
                }
                Ok(())
            }
            XlsxCommand::MoveSheet(args) => {
                move_sheet(
                    &args.file,
                    &args.name,
                    args.before.as_deref(),
                    args.after.as_deref(),
                )?;
                if args.json {
                    print_json(&serde_json::json!({ "file": args.file, "name": args.name }))?;
                } else {
                    println!("  Moved sheet {}", args.name);
                    println!("\nSaved via direct OOXML engine: {}", args.file.display());
                }
                Ok(())
            }
            XlsxCommand::RenameSheet(args) => {
                validate_sheet_name(&args.name)?;
                println!("{}", crate::wincom::excel_rename_sheet(&args)?);
                Ok(())
            }
            XlsxCommand::Validate(args) => {
                println!("{}", crate::wincom::excel_validate(&args)?);
                Ok(())
            }
            XlsxCommand::Insert(args) => {
                println!("{}", crate::wincom::excel_insert(&args)?);
                Ok(())
            }
            XlsxCommand::AutoFit(args) => {
                println!("{}", crate::wincom::excel_autofit(&args)?);
                Ok(())
            }
            XlsxCommand::Format(args) => {
                println!("{}", crate::wincom::excel_format(&args)?);
                Ok(())
            }
            XlsxCommand::CopySheets(args) => {
                println!("{}", crate::wincom::excel_copy_sheets(&args)?);
                Ok(())
            }
        }
    }
}

pub fn read(args: &ReadArgs) -> Result<ReadResult> {
    let mut workbook = open_workbook_auto(&args.file)
        .with_context(|| format!("open workbook: {}", args.file.display()))?;
    let sheet_names = workbook.sheet_names().to_vec();
    let sheet = args
        .sheet
        .clone()
        .or_else(|| sheet_names.first().cloned())
        .ok_or_else(|| anyhow!("workbook has no sheets"))?;
    let range = workbook
        .worksheet_range(&sheet)
        .with_context(|| format!("read worksheet: {sheet}"))?;
    let rows = read_range_values(&range, args.range.as_deref())?;
    let rows = compact_display_rows(rows, args.show_empty_rows, args.show_empty_cols);
    let dimensions = args
        .range
        .clone()
        .unwrap_or_else(|| dimensions_for_rows(&rows));
    Ok(ReadResult {
        sheet,
        dimensions,
        rows: rows
            .into_iter()
            .map(|row| {
                row.into_iter()
                    .map(|value| format_cell_value(&value))
                    .collect()
            })
            .collect(),
    })
}

pub fn cells(args: &CellsArgs) -> Result<Vec<CellRecord>> {
    let mut workbook = open_workbook_auto(&args.file)
        .with_context(|| format!("open workbook: {}", args.file.display()))?;
    let sheet_names = workbook.sheet_names().to_vec();
    let sheet = args
        .sheet
        .clone()
        .or_else(|| sheet_names.first().cloned())
        .ok_or_else(|| anyhow!("workbook has no sheets"))?;
    let range = workbook
        .worksheet_range(&sheet)
        .with_context(|| format!("read worksheet: {sheet}"))?;
    let (start_row, start_col, end_row, end_col) = if let Some(range_ref) = args.range.as_deref() {
        let bounds = range_bounds(range_ref)?;
        (
            bounds.min_row.saturating_sub(1),
            bounds.min_col.saturating_sub(1),
            bounds.max_row.saturating_sub(1),
            bounds.max_col.saturating_sub(1),
        )
    } else {
        let height = range.height();
        let width = range.width();
        if height == 0 || width == 0 {
            return Ok(Vec::new());
        }
        (0, 0, height - 1, width - 1)
    };
    let mut out = Vec::new();
    for row_idx in start_row..=end_row {
        for col_idx in start_col..=end_col {
            let value = range
                .get((row_idx, col_idx))
                .cloned()
                .unwrap_or(Data::Empty);
            if !args.include_empty && is_blank_value(&value) {
                continue;
            }
            out.push(CellRecord {
                sheet: sheet.clone(),
                cell: cell_ref(col_idx + 1, row_idx + 1),
                value: format_cell_value(&value),
            });
        }
    }
    Ok(out)
}

pub fn list_sheets(path: &Path) -> Result<Vec<SheetInfo>> {
    let package = XlsxPackage::open(path)?;
    let Some(sheets) = child(&package.workbook, "sheets") else {
        return Ok(Vec::new());
    };
    Ok(children(sheets, "sheet")
        .enumerate()
        .map(|(index, sheet)| SheetInfo {
            index: index + 1,
            name: attr(sheet, "name").unwrap_or_default().to_string(),
            sheet_id: attr(sheet, "sheetId").map(ToString::to_string),
            state: attr(sheet, "state").map(ToString::to_string),
        })
        .collect())
}

pub fn search(args: &SearchArgs) -> Result<Vec<SearchHit>> {
    let mut workbook = open_workbook_auto(&args.file)
        .with_context(|| format!("open workbook: {}", args.file.display()))?;
    let query = args.query.to_lowercase();
    let sheet_names = if let Some(sheet) = &args.sheet {
        vec![sheet.clone()]
    } else {
        workbook.sheet_names().to_vec()
    };
    let mut hits = Vec::new();
    for sheet in sheet_names {
        let range = workbook
            .worksheet_range(&sheet)
            .with_context(|| format!("read worksheet: {sheet}"))?;
        for (row_idx, row) in range.rows().enumerate() {
            for (col_idx, cell) in row.iter().enumerate() {
                let value = format_cell_value(cell);
                if !value.is_empty() && value.to_lowercase().contains(&query) {
                    hits.push(SearchHit {
                        sheet: sheet.clone(),
                        cell: cell_ref(col_idx + 1, row_idx + 1),
                        value,
                    });
                }
            }
        }
    }
    Ok(hits)
}

pub fn inspect(args: &InspectArgs) -> Result<WorkbookInspection> {
    let package = XlsxPackage::open(&args.file)?;
    let mut sheets = Vec::new();
    let ordered_names = sheet_names_in_order(&package.workbook)?;
    for name in ordered_names {
        if args.sheet.as_ref().is_some_and(|wanted| wanted != &name) {
            continue;
        }
        let Some(target) = package.sheets.get(&name) else {
            continue;
        };
        let worksheet = package.read_xml(&target.part)?;
        sheets.push(inspect_sheet(&args.file, &name, &target.part, &worksheet)?);
    }
    let defined_names = collect_defined_names(&package.workbook)?;
    let calc_properties = child(&package.workbook, "calcPr")
        .map(|calc_pr| calc_pr.attributes.clone().into_iter().collect())
        .unwrap_or_default();
    let workbook_relationships = children(&package.rels, "Relationship")
        .map(|rel| RelationshipInspection {
            id: attr(rel, "Id").unwrap_or_default().to_string(),
            relationship_type: attr(rel, "Type").unwrap_or_default().to_string(),
            target: attr(rel, "Target").unwrap_or_default().to_string(),
            target_mode: attr(rel, "TargetMode").map(ToString::to_string),
        })
        .collect();
    Ok(WorkbookInspection {
        file: args.file.clone(),
        date1904: workbook_uses_1904_dates(&package.workbook),
        sheets,
        defined_names,
        calc_properties,
        workbook_relationships,
    })
}

pub fn relationships(args: &RelationshipsArgs) -> Result<Vec<RelationshipInfo>> {
    let package = XlsxPackage::open(&args.file)?;
    let mut out = Vec::new();
    if args.sheet.is_none() {
        out.extend(
            children(&package.rels, "Relationship").map(|rel| RelationshipInfo {
                scope: "workbook".to_string(),
                sheet: None,
                part: WORKBOOK.to_string(),
                id: attr(rel, "Id").unwrap_or_default().to_string(),
                relationship_type: attr(rel, "Type").unwrap_or_default().to_string(),
                target: attr(rel, "Target").unwrap_or_default().to_string(),
                target_mode: attr(rel, "TargetMode").map(ToString::to_string),
            }),
        );
    }
    for sheet_name in sheet_names_in_order(&package.workbook)? {
        if args
            .sheet
            .as_ref()
            .is_some_and(|wanted| wanted != &sheet_name)
        {
            continue;
        }
        let Some(target) = package.sheets.get(&sheet_name) else {
            continue;
        };
        for rel in sheet_relationships(&args.file, &target.part)? {
            out.push(RelationshipInfo {
                scope: "worksheet".to_string(),
                sheet: Some(sheet_name.clone()),
                part: target.part.clone(),
                id: rel.id,
                relationship_type: rel.relationship_type,
                target: rel.target,
                target_mode: rel.target_mode,
            });
        }
    }
    Ok(out)
}

pub fn defined_names(args: &DefinedNamesArgs) -> Result<Vec<DefinedNameInspection>> {
    let package = XlsxPackage::open(&args.file)?;
    let mut names = collect_defined_names(&package.workbook)?;
    if let Some(name) = &args.name {
        names.retain(|item| item.name == *name);
    }
    Ok(names)
}

fn collect_defined_names(workbook: &Element) -> Result<Vec<DefinedNameInspection>> {
    let sheet_names = sheet_names_in_order(workbook)?;
    let Some(defined_names) = child(workbook, "definedNames") else {
        return Ok(Vec::new());
    };
    Ok(children(defined_names, "definedName")
        .map(|defined_name| {
            let local_sheet_id = attr(defined_name, "localSheetId").map(ToString::to_string);
            let sheet = local_sheet_id
                .as_deref()
                .and_then(|id| id.parse::<usize>().ok())
                .and_then(|index| sheet_names.get(index).cloned());
            DefinedNameInspection {
                name: attr(defined_name, "name").unwrap_or_default().to_string(),
                local_sheet_id,
                sheet,
                value: element_text_all(defined_name),
            }
        })
        .collect())
}

pub fn merged_ranges(args: &MergedRangesArgs) -> Result<Vec<MergedRangeInfo>> {
    let package = XlsxPackage::open(&args.file)?;
    let mut out = Vec::new();
    for sheet_name in sheet_names_in_order(&package.workbook)? {
        if args
            .sheet
            .as_ref()
            .is_some_and(|wanted| wanted != &sheet_name)
        {
            continue;
        }
        let Some(target) = package.sheets.get(&sheet_name) else {
            continue;
        };
        let worksheet = package.read_xml(&target.part)?;
        for range in collect_merged_ranges(&worksheet) {
            out.push(MergedRangeInfo {
                sheet: sheet_name.clone(),
                range,
            });
        }
    }
    Ok(out)
}

pub fn auto_filters(args: &AutoFiltersArgs) -> Result<Vec<AutoFilterInfo>> {
    let package = XlsxPackage::open(&args.file)?;
    let mut out = Vec::new();
    for sheet_name in sheet_names_in_order(&package.workbook)? {
        if args
            .sheet
            .as_ref()
            .is_some_and(|wanted| wanted != &sheet_name)
        {
            continue;
        }
        let Some(target) = package.sheets.get(&sheet_name) else {
            continue;
        };
        let worksheet = package.read_xml(&target.part)?;
        if let Some(auto_filter) = child(&worksheet, "autoFilter") {
            out.push(AutoFilterInfo {
                sheet: sheet_name,
                range: attr(auto_filter, "ref").map(ToString::to_string),
            });
        }
    }
    Ok(out)
}

pub fn protections(args: &ProtectionsArgs) -> Result<Vec<ProtectionInfo>> {
    let package = XlsxPackage::open(&args.file)?;
    let mut out = Vec::new();
    for sheet_name in sheet_names_in_order(&package.workbook)? {
        if args
            .sheet
            .as_ref()
            .is_some_and(|wanted| wanted != &sheet_name)
        {
            continue;
        }
        let Some(target) = package.sheets.get(&sheet_name) else {
            continue;
        };
        let worksheet = package.read_xml(&target.part)?;
        if let Some(protection) = child(&worksheet, "sheetProtection") {
            out.push(ProtectionInfo {
                sheet: sheet_name,
                attributes: protection.attributes.clone().into_iter().collect(),
            });
        }
    }
    Ok(out)
}

pub fn formulas(args: &FormulasArgs) -> Result<Vec<FormulaInfo>> {
    let package = XlsxPackage::open(&args.file)?;
    let mut out = Vec::new();
    for sheet_name in sheet_names_in_order(&package.workbook)? {
        if args
            .sheet
            .as_ref()
            .is_some_and(|wanted| wanted != &sheet_name)
        {
            continue;
        }
        let Some(target) = package.sheets.get(&sheet_name) else {
            continue;
        };
        let worksheet = package.read_xml(&target.part)?;
        collect_sheet_formulas(&sheet_name, &worksheet, &mut out);
    }
    Ok(out)
}

fn collect_sheet_formulas(sheet_name: &str, worksheet: &Element, out: &mut Vec<FormulaInfo>) {
    let Some(sheet_data) = child(worksheet, "sheetData") else {
        return;
    };
    for row in children(sheet_data, "row") {
        for cell in children(row, "c") {
            let Some(formula) = child(cell, "f") else {
                continue;
            };
            let formula_text = element_text_all(formula);
            if formula_text.trim().is_empty() {
                continue;
            }
            out.push(FormulaInfo {
                sheet: sheet_name.to_string(),
                cell: attr(cell, "r").unwrap_or_default().to_string(),
                formula: formula_text,
            });
        }
    }
}

pub fn tables(args: &TablesArgs) -> Result<Vec<TableInfo>> {
    let package = XlsxPackage::open(&args.file)?;
    let mut out = Vec::new();
    for sheet_name in sheet_names_in_order(&package.workbook)? {
        if args
            .sheet
            .as_ref()
            .is_some_and(|wanted| wanted != &sheet_name)
        {
            continue;
        }
        let Some(target) = package.sheets.get(&sheet_name) else {
            continue;
        };
        for table_part in sheet_table_parts(args.file.as_path(), &target.part)? {
            let table = package.read_xml(&table_part)?;
            out.push(TableInfo {
                sheet: sheet_name.clone(),
                part: table_part,
                name: attr(&table, "name").map(ToString::to_string),
                display_name: attr(&table, "displayName").map(ToString::to_string),
                range: attr(&table, "ref").map(ToString::to_string),
            });
        }
    }
    Ok(out)
}

pub fn validations(args: &ValidationsArgs) -> Result<Vec<DataValidationInfo>> {
    let package = XlsxPackage::open(&args.file)?;
    let mut out = Vec::new();
    for sheet_name in sheet_names_in_order(&package.workbook)? {
        if args
            .sheet
            .as_ref()
            .is_some_and(|wanted| wanted != &sheet_name)
        {
            continue;
        }
        let Some(target) = package.sheets.get(&sheet_name) else {
            continue;
        };
        let worksheet = package.read_xml(&target.part)?;
        collect_sheet_validations(&sheet_name, &worksheet, &mut out);
    }
    Ok(out)
}

pub fn conditional_formatting(
    args: &ConditionalFormattingArgs,
) -> Result<Vec<ConditionalFormattingInfo>> {
    let package = XlsxPackage::open(&args.file)?;
    let mut out = Vec::new();
    for sheet_name in sheet_names_in_order(&package.workbook)? {
        if args
            .sheet
            .as_ref()
            .is_some_and(|wanted| wanted != &sheet_name)
        {
            continue;
        }
        let Some(target) = package.sheets.get(&sheet_name) else {
            continue;
        };
        let worksheet = package.read_xml(&target.part)?;
        collect_sheet_conditional_formatting(&sheet_name, &worksheet, &mut out);
    }
    Ok(out)
}

pub fn hyperlinks(args: &HyperlinksArgs) -> Result<Vec<HyperlinkInfo>> {
    let package = XlsxPackage::open(&args.file)?;
    let mut out = Vec::new();
    for sheet_name in sheet_names_in_order(&package.workbook)? {
        if args
            .sheet
            .as_ref()
            .is_some_and(|wanted| wanted != &sheet_name)
        {
            continue;
        }
        let Some(target) = package.sheets.get(&sheet_name) else {
            continue;
        };
        let worksheet = package.read_xml(&target.part)?;
        let relationships = sheet_relationship_targets(args.file.as_path(), &target.part)?;
        collect_sheet_hyperlinks(&sheet_name, &worksheet, &relationships, &mut out);
    }
    Ok(out)
}

pub fn comments(args: &CommentsArgs) -> Result<Vec<CommentInfo>> {
    let package = XlsxPackage::open(&args.file)?;
    let mut out = Vec::new();
    for sheet_name in sheet_names_in_order(&package.workbook)? {
        if args
            .sheet
            .as_ref()
            .is_some_and(|wanted| wanted != &sheet_name)
        {
            continue;
        }
        let Some(target) = package.sheets.get(&sheet_name) else {
            continue;
        };
        let relationships = sheet_relationship_targets(args.file.as_path(), &target.part)?;
        for comments_part in relationships
            .values()
            .filter(|rel| rel.relationship_type.ends_with("/comments"))
            .map(|rel| rel.target.clone())
        {
            let comments_xml = package.read_xml(&comments_part)?;
            collect_comments_part(&sheet_name, &comments_part, &comments_xml, &mut out);
        }
    }
    Ok(out)
}

fn collect_comments_part(
    sheet_name: &str,
    comments_part: &str,
    comments_xml: &Element,
    out: &mut Vec<CommentInfo>,
) {
    let authors = child(comments_xml, "authors")
        .map(|authors| {
            children(authors, "author")
                .map(element_text_all)
                .collect::<Vec<_>>()
        })
        .unwrap_or_default();
    let Some(comment_list) = child(comments_xml, "commentList") else {
        return;
    };
    for comment in children(comment_list, "comment") {
        let author = attr(comment, "authorId")
            .and_then(|id| id.parse::<usize>().ok())
            .and_then(|idx| authors.get(idx).cloned());
        let text = child(comment, "text")
            .map(element_text_all)
            .unwrap_or_default();
        out.push(CommentInfo {
            sheet: sheet_name.to_string(),
            cell: attr(comment, "ref").map(ToString::to_string),
            author,
            text,
            part: comments_part.to_string(),
        });
    }
}

fn collect_sheet_hyperlinks(
    sheet_name: &str,
    worksheet: &Element,
    relationships: &BTreeMap<String, SheetRelationshipTarget>,
    out: &mut Vec<HyperlinkInfo>,
) {
    let Some(hyperlinks) = child(worksheet, "hyperlinks") else {
        return;
    };
    for hyperlink in children(hyperlinks, "hyperlink") {
        let rid = attr(hyperlink, "r:id").or_else(|| attr(hyperlink, "id"));
        let target = rid.and_then(|id| relationships.get(id).map(|rel| rel.target.clone()));
        out.push(HyperlinkInfo {
            sheet: sheet_name.to_string(),
            reference: attr(hyperlink, "ref").map(ToString::to_string),
            target,
            location: attr(hyperlink, "location").map(ToString::to_string),
            display: attr(hyperlink, "display").map(ToString::to_string),
            tooltip: attr(hyperlink, "tooltip").map(ToString::to_string),
        });
    }
}

fn collect_sheet_conditional_formatting(
    sheet_name: &str,
    worksheet: &Element,
    out: &mut Vec<ConditionalFormattingInfo>,
) {
    for conditional_formatting in children(worksheet, "conditionalFormatting") {
        let range = attr(conditional_formatting, "sqref").map(ToString::to_string);
        for rule in children(conditional_formatting, "cfRule") {
            out.push(ConditionalFormattingInfo {
                sheet: sheet_name.to_string(),
                range: range.clone(),
                rule_type: attr(rule, "type").map(ToString::to_string),
                priority: attr(rule, "priority").map(ToString::to_string),
                operator: attr(rule, "operator").map(ToString::to_string),
                formulas: children(rule, "formula").map(element_text_all).collect(),
            });
        }
    }
}

fn collect_sheet_validations(
    sheet_name: &str,
    worksheet: &Element,
    out: &mut Vec<DataValidationInfo>,
) {
    let Some(data_validations) = child(worksheet, "dataValidations") else {
        return;
    };
    for data_validation in children(data_validations, "dataValidation") {
        out.push(DataValidationInfo {
            sheet: sheet_name.to_string(),
            range: attr(data_validation, "sqref").map(ToString::to_string),
            validation_type: attr(data_validation, "type").map(ToString::to_string),
            operator: attr(data_validation, "operator").map(ToString::to_string),
            allow_blank: attr(data_validation, "allowBlank").map(ToString::to_string),
            formula1: child(data_validation, "formula1").map(element_text_all),
            formula2: child(data_validation, "formula2").map(element_text_all),
        });
    }
}

#[derive(Debug, Clone)]
struct SheetRelationshipTarget {
    relationship_type: String,
    target: String,
}

fn sheet_relationship_targets(
    path: &Path,
    sheet_part: &str,
) -> Result<BTreeMap<String, SheetRelationshipTarget>> {
    let (dir, file) = sheet_part
        .rsplit_once('/')
        .unwrap_or(("xl/worksheets", sheet_part));
    let rels_part = format!("{dir}/_rels/{file}.rels");
    let Ok(data) = read_part(path, &rels_part) else {
        return Ok(BTreeMap::new());
    };
    let rels = parse_xml(&data)?;
    let mut out = BTreeMap::new();
    for rel in children(&rels, "Relationship") {
        let Some(id) = attr(rel, "Id") else {
            continue;
        };
        let relationship_type = attr(rel, "Type").unwrap_or_default().to_string();
        let Some(target) = attr(rel, "Target") else {
            continue;
        };
        let target = if attr(rel, "TargetMode") == Some("External") {
            target.to_string()
        } else {
            resolve_relationship_target(sheet_part, target)
        };
        out.insert(
            id.to_string(),
            SheetRelationshipTarget {
                relationship_type,
                target,
            },
        );
    }
    Ok(out)
}

fn sheet_table_parts(path: &Path, sheet_part: &str) -> Result<Vec<String>> {
    let (dir, file) = sheet_part
        .rsplit_once('/')
        .unwrap_or(("xl/worksheets", sheet_part));
    let rels_part = format!("{dir}/_rels/{file}.rels");
    let Ok(data) = read_part(path, &rels_part) else {
        return Ok(Vec::new());
    };
    let rels = parse_xml(&data)?;
    Ok(children(&rels, "Relationship")
        .filter(|rel| {
            attr(rel, "Type").is_some_and(|relationship_type| relationship_type.ends_with("/table"))
        })
        .filter_map(|rel| attr(rel, "Target"))
        .map(|target| resolve_relationship_target(sheet_part, target))
        .collect())
}

pub fn create(args: &CreateArgs) -> Result<CreateResult> {
    let text = fs::read_to_string(&args.spec)
        .with_context(|| format!("read XLSX create spec: {}", args.spec.display()))?;
    let spec: CreateWorkbookSpec = serde_json::from_str(&text).context("parse XLSX create spec")?;
    create_workbook(&spec, &args.output, args.overwrite)
}

pub fn create_workbook(
    spec: &CreateWorkbookSpec,
    output: &Path,
    overwrite: bool,
) -> Result<CreateResult> {
    if output.exists() && !overwrite {
        bail!(
            "output already exists: {}; pass --overwrite to replace it",
            output.display()
        );
    }
    let sheets = normalized_create_sheets(spec)?;
    let mut parts = PartMap::new();
    parts.insert(
        CONTENT_TYPES.to_string(),
        create_content_types(sheets.len()),
    );
    parts.insert("_rels/.rels".to_string(), create_root_rels());
    parts.insert(
        WORKBOOK_RELS.to_string(),
        create_workbook_rels(sheets.len()),
    );

    let mut cells_written = 0usize;
    let mut formula_written = false;
    for (idx, sheet) in sheets.iter().enumerate() {
        let (xml, written, has_formula) = create_worksheet_xml(sheet, spec.date1904)
            .with_context(|| format!("create worksheet XML for '{}'", sheet.name))?;
        parts.insert(format!("xl/worksheets/sheet{}.xml", idx + 1), xml);
        cells_written += written;
        formula_written |= has_formula;
    }
    parts.insert(
        WORKBOOK.to_string(),
        create_workbook_xml(&sheets, spec.date1904, formula_written),
    );
    write_new_package(output, &parts)?;
    Ok(CreateResult {
        file: output.to_path_buf(),
        sheets: sheets
            .iter()
            .enumerate()
            .map(|(idx, sheet)| SheetInfo {
                index: idx + 1,
                name: sheet.name.clone(),
                sheet_id: Some((idx + 1).to_string()),
                state: None,
            })
            .collect(),
        cells_written,
        formula_recalc_marked: formula_written,
    })
}

fn normalized_create_sheets(spec: &CreateWorkbookSpec) -> Result<Vec<CreateSheetSpec>> {
    let sheets = if spec.sheets.is_empty() {
        vec![CreateSheetSpec {
            name: "Sheet1".to_string(),
            rows: Vec::new(),
            cells: Vec::new(),
        }]
    } else {
        spec.sheets
            .iter()
            .map(|sheet| CreateSheetSpec {
                name: sheet.name.clone(),
                rows: sheet.rows.clone(),
                cells: sheet.cells.clone(),
            })
            .collect::<Vec<_>>()
    };
    let mut names = BTreeSet::new();
    for sheet in &sheets {
        validate_sheet_name(&sheet.name)?;
        let key = sheet.name.to_ascii_lowercase();
        if !names.insert(key) {
            bail!("duplicate sheet name: {}", sheet.name);
        }
    }
    Ok(sheets)
}

fn create_worksheet_xml(sheet: &CreateSheetSpec, date1904: bool) -> Result<(Vec<u8>, usize, bool)> {
    let mut worksheet = parse_xml(&minimal_worksheet_xml()?)?;
    let mut cells_written = 0usize;
    let mut formula_written = false;

    for (row_idx, row) in sheet.rows.iter().enumerate() {
        for (col_idx, value) in row.iter().enumerate() {
            formula_written |= set_worksheet_cell(
                &mut worksheet,
                row_idx + 1,
                col_idx + 1,
                value,
                None,
                date1904,
            )?;
            cells_written += 1;
        }
    }

    for edit in &sheet.cells {
        if edit.sheet.as_ref().is_some_and(|name| name != &sheet.name) {
            bail!(
                "cell edit sheet '{}' does not match enclosing sheet '{}'",
                edit.sheet.as_deref().unwrap_or_default(),
                sheet.name
            );
        }
        let target = edit
            .cell
            .as_ref()
            .or(edit.range_ref.as_ref())
            .ok_or_else(|| anyhow!("cell edit in '{}' is missing cell/range", sheet.name))?;
        let value = edit
            .values
            .clone()
            .or_else(|| edit.value.clone())
            .unwrap_or(Value::Null);
        for (row_idx, col_idx, item) in iter_target_cells(target, &value)? {
            formula_written |= set_worksheet_cell(
                &mut worksheet,
                row_idx,
                col_idx,
                &item,
                edit.value_type,
                date1904,
            )?;
            cells_written += 1;
        }
    }
    update_worksheet_dimension(&mut worksheet)?;
    Ok((write_xml(&worksheet)?, cells_written, formula_written))
}

fn set_worksheet_cell(
    worksheet: &mut Element,
    row_idx: usize,
    col_idx: usize,
    value: &Value,
    value_type: Option<ValueType>,
    date1904: bool,
) -> Result<bool> {
    let formula_written = matches!(value, Value::String(text) if text.starts_with('='));
    let sheet_data_idx = ensure_sheet_data(worksheet);
    let sheet_data = element_child_mut_at(worksheet, sheet_data_idx)?;
    let row_idx_in_parent = ensure_row(sheet_data, row_idx);
    let row = element_child_mut_at(sheet_data, row_idx_in_parent)?;
    let cell_idx = ensure_cell(row, row_idx, col_idx)?;
    let cell = element_child_mut_at(row, cell_idx)?;
    set_cell_value(cell, value, value_type, date1904)?;
    Ok(formula_written)
}

fn create_content_types(sheet_count: usize) -> Vec<u8> {
    let mut xml = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>"#,
    );
    xml.push_str(&format!(
        r#"<Override PartName="/xl/workbook.xml" ContentType="{WORKBOOK_CONTENT_TYPE}"/>"#
    ));
    for idx in 1..=sheet_count {
        xml.push_str(&format!(
            r#"<Override PartName="/xl/worksheets/sheet{idx}.xml" ContentType="{WORKSHEET_CONTENT_TYPE}"/>"#
        ));
    }
    xml.push_str("</Types>");
    xml.into_bytes()
}

fn create_root_rels() -> Vec<u8> {
    br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
        .to_vec()
}

fn create_workbook_rels(sheet_count: usize) -> Vec<u8> {
    let mut xml = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    );
    for idx in 1..=sheet_count {
        xml.push_str(&format!(
            r#"<Relationship Id="rId{idx}" Type="{WORKSHEET_REL_TYPE}" Target="worksheets/sheet{idx}.xml"/>"#
        ));
    }
    xml.push_str("</Relationships>");
    xml.into_bytes()
}

fn create_workbook_xml(
    sheets: &[CreateSheetSpec],
    date1904: bool,
    formula_written: bool,
) -> Vec<u8> {
    let mut xml = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
    );
    if date1904 {
        xml.push_str(r#"<workbookPr date1904="1"/>"#);
    }
    xml.push_str("<sheets>");
    for (idx, sheet) in sheets.iter().enumerate() {
        let id = idx + 1;
        xml.push_str(&format!(
            r#"<sheet name="{}" sheetId="{id}" r:id="rId{id}"/>"#,
            escaped(&sheet.name)
        ));
    }
    xml.push_str("</sheets>");
    if formula_written {
        xml.push_str(r#"<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>"#);
    }
    xml.push_str("</workbook>");
    xml.into_bytes()
}

fn inspect_sheet(
    path: &Path,
    name: &str,
    part: &str,
    worksheet: &Element,
) -> Result<SheetInspection> {
    let dimension = child(worksheet, "dimension")
        .and_then(|dimension| attr(dimension, "ref").map(ToString::to_string));
    let sheet_data = child(worksheet, "sheetData");
    let rows = sheet_data
        .map(|data| children(data, "row").count())
        .unwrap_or(0);
    let mut cells = 0;
    let mut formulas = 0;
    if let Some(sheet_data) = sheet_data {
        for row in children(sheet_data, "row") {
            for cell in children(row, "c") {
                cells += 1;
                if child(cell, "f").is_some() {
                    formulas += 1;
                }
            }
        }
    }
    let merged_ranges = collect_merged_ranges(worksheet);
    let conditional_formatting_ranges = children(worksheet, "conditionalFormatting")
        .filter_map(|cf| attr(cf, "sqref").map(ToString::to_string))
        .collect();
    let data_validation_ranges = child(worksheet, "dataValidations")
        .map(|validations| {
            children(validations, "dataValidation")
                .filter_map(|dv| attr(dv, "sqref").map(ToString::to_string))
                .collect()
        })
        .unwrap_or_default();
    let hyperlinks = child(worksheet, "hyperlinks")
        .map(|links| children(links, "hyperlink").count())
        .unwrap_or(0);
    let table_relationships = sheet_relationships(path, part)?
        .into_iter()
        .filter(|rel| rel.relationship_type.ends_with("/table"))
        .map(|rel| rel.target)
        .collect();
    Ok(SheetInspection {
        name: name.to_string(),
        part: part.to_string(),
        dimension,
        rows,
        cells,
        formulas,
        merged_ranges,
        conditional_formatting_ranges,
        data_validation_ranges,
        hyperlinks,
        has_auto_filter: child(worksheet, "autoFilter").is_some(),
        has_sheet_protection: child(worksheet, "sheetProtection").is_some(),
        table_relationships,
    })
}

fn collect_merged_ranges(worksheet: &Element) -> Vec<String> {
    child(worksheet, "mergeCells")
        .map(|merge_cells| {
            children(merge_cells, "mergeCell")
                .filter_map(|merge| attr(merge, "ref").map(ToString::to_string))
                .collect()
        })
        .unwrap_or_default()
}

fn sheet_relationships(path: &Path, sheet_part: &str) -> Result<Vec<RelationshipInspection>> {
    let (dir, file) = sheet_part
        .rsplit_once('/')
        .unwrap_or(("xl/worksheets", sheet_part));
    let rels_part = format!("{dir}/_rels/{file}.rels");
    let Ok(data) = read_part(path, &rels_part) else {
        return Ok(Vec::new());
    };
    let rels = parse_xml(&data)?;
    Ok(children(&rels, "Relationship")
        .map(|rel| RelationshipInspection {
            id: attr(rel, "Id").unwrap_or_default().to_string(),
            relationship_type: attr(rel, "Type").unwrap_or_default().to_string(),
            target: attr(rel, "Target").unwrap_or_default().to_string(),
            target_mode: attr(rel, "TargetMode").map(ToString::to_string),
        })
        .collect())
}

pub fn edit(args: &EditArgs) -> Result<EditResult> {
    if !args.file.exists() {
        bail!("file not found: {}", args.file.display());
    }
    let edits = load_edits_from_args(args)?;
    direct_edit_workbook(
        &args.file,
        &edits,
        args.sheet.as_deref(),
        args.allow_protected,
    )
}

fn load_edits_from_args(args: &EditArgs) -> Result<Vec<CellEdit>> {
    if args.edits.is_some() && args.edits_file.is_some() {
        bail!("provide --edits or --edits-file, not both");
    }
    if args.clear && args.value.is_some() {
        bail!("provide --clear or --value, not both");
    }
    if let Some(path) = &args.edits_file {
        let text = fs::read_to_string(path)
            .with_context(|| format!("read edits file: {}", path.display()))?;
        return serde_json::from_str(&text)
            .with_context(|| format!("parse edits JSON: {}", path.display()));
    }
    if let Some(text) = &args.edits {
        return serde_json::from_str(text).context("parse --edits JSON");
    }
    let target = args.cell.clone().or_else(|| args.range.clone());
    if let Some(target) = target
        && (args.clear || args.value.is_some())
    {
        let mut edit = CellEdit {
            sheet: args.sheet.clone(),
            cell: args.cell.clone(),
            range_ref: args.range.clone(),
            value: if args.clear {
                Some(Value::Null)
            } else {
                args.value.clone().map(Value::String)
            },
            values: None,
            value_type: args.value_type,
        };
        if args.cell.is_none() && args.range.is_none() {
            edit.cell = Some(target);
        }
        return Ok(vec![edit]);
    }
    bail!("provide --cell/--value, --range/--value, --clear, --edits JSON, or --edits-file");
}

fn direct_edit_workbook(
    path: &Path,
    edits: &[CellEdit],
    default_sheet: Option<&str>,
    allow_protected: bool,
) -> Result<EditResult> {
    let mut package = XlsxPackage::open(path)?;
    let date1904 = workbook_uses_1904_dates(&package.workbook);
    let mut grouped: BTreeMap<String, Vec<(String, Value, Option<ValueType>)>> = BTreeMap::new();
    for edit in edits {
        let sheet = edit
            .sheet
            .as_deref()
            .or(default_sheet)
            .ok_or_else(|| anyhow!("--sheet is required unless each edit includes a sheet"))?;
        let target = edit
            .cell
            .as_ref()
            .or(edit.range_ref.as_ref())
            .ok_or_else(|| anyhow!("edit is missing cell/range: {edit:?}"))?;
        if !package.sheets.contains_key(sheet) {
            bail!("sheet '{sheet}' not found");
        }
        let value = edit
            .values
            .clone()
            .or_else(|| edit.value.clone())
            .unwrap_or(Value::Null);
        grouped.entry(sheet.to_string()).or_default().push((
            target.clone(),
            value,
            edit.value_type,
        ));
    }

    let mut updates = PartMap::new();
    let mut formula_written = false;
    for (sheet, sheet_edits) in &grouped {
        let part = package
            .sheets
            .get(sheet)
            .expect("sheet key already checked")
            .part
            .clone();
        let mut worksheet = package.read_xml(&part)?;
        if !allow_protected && child(&worksheet, "sheetProtection").is_some() {
            bail!(
                "sheet '{sheet}' is protected; use Excel COM for locked-cell preflight or pass --allow-protected"
            );
        }
        let sheet_data_idx = ensure_sheet_data(&mut worksheet);
        let sheet_data = element_child_mut_at(&mut worksheet, sheet_data_idx)?;
        for (target, value, value_type) in sheet_edits {
            for (row_idx, col_idx, item) in iter_target_cells(target, value)? {
                if matches!(&item, Value::String(text) if text.starts_with('=')) {
                    formula_written = true;
                }
                let row_idx_in_parent = ensure_row(sheet_data, row_idx);
                let row = element_child_mut_at(sheet_data, row_idx_in_parent)?;
                let cell_idx = ensure_cell(row, row_idx, col_idx)?;
                let cell = element_child_mut_at(row, cell_idx)?;
                set_cell_value(cell, &item, *value_type, date1904)?;
            }
        }
        update_worksheet_dimension(&mut worksheet)?;
        updates.insert(part, write_xml(&worksheet)?);
    }

    let removed_calc_chain_parts = if formula_written {
        set_workbook_recalc_flags(&mut package.workbook);
        let removed = remove_calc_chain_relationships(&mut package.rels);
        if !removed.is_empty() {
            let mut content_types = package.read_xml(CONTENT_TYPES)?;
            for part in &removed {
                remove_content_type_override(&mut content_types, part);
            }
            updates.insert(CONTENT_TYPES.to_string(), write_xml(&content_types)?);
            updates.insert(WORKBOOK_RELS.to_string(), write_xml(&package.rels)?);
        }
        updates.insert(WORKBOOK.to_string(), write_xml(&package.workbook)?);
        removed
    } else {
        Vec::new()
    };

    rewrite_parts(path, &updates, removed_calc_chain_parts)?;
    Ok(EditResult {
        file: path.to_path_buf(),
        target_count: grouped.values().map(Vec::len).sum(),
        sheets: grouped
            .into_iter()
            .map(|(sheet, edits)| (sheet, edits.len()))
            .collect(),
        formula_recalc_marked: formula_written,
    })
}

pub fn add_sheet(
    path: &Path,
    name: &str,
    before: Option<&str>,
    after: Option<&str>,
) -> Result<String> {
    if before.is_some() && after.is_some() {
        bail!("provide --before or --after, not both");
    }
    validate_sheet_name(name)?;
    let mut package = XlsxPackage::open(path)?;
    if package
        .sheets
        .keys()
        .any(|existing| existing.eq_ignore_ascii_case(name))
    {
        bail!("sheet '{name}' already exists");
    }
    let old_order = sheet_names_in_order(&package.workbook)?;
    let part = next_sheet_part(path)?;
    let rid = next_relationship_id(&package.rels);
    let rel_target = relationship_target_for_part(WORKBOOK, &part);
    let mut rel = Element::new("Relationship");
    rel.attributes.insert("Id".to_string(), rid.clone());
    rel.attributes
        .insert("Type".to_string(), WORKSHEET_REL_TYPE.to_string());
    rel.attributes.insert("Target".to_string(), rel_target);
    package.rels.children.push(XMLNode::Element(rel));

    let mut sheet = Element::new("sheet");
    sheet
        .attributes
        .insert("name".to_string(), name.to_string());
    sheet.attributes.insert(
        "sheetId".to_string(),
        next_sheet_id(&package.workbook)?.to_string(),
    );
    sheet.attributes.insert("r:id".to_string(), rid);
    insert_sheet_element(&mut package.workbook, sheet, before, after)?;
    let new_order = sheet_names_in_order(&package.workbook)?;
    update_sheet_index_references(&mut package.workbook, &old_order, &new_order)?;

    let mut content_types = package.read_xml(CONTENT_TYPES)?;
    ensure_content_type_override(&mut content_types, &part, WORKSHEET_CONTENT_TYPE);
    let mut updates = PartMap::new();
    updates.insert(WORKBOOK.to_string(), write_xml(&package.workbook)?);
    updates.insert(WORKBOOK_RELS.to_string(), write_xml(&package.rels)?);
    updates.insert(CONTENT_TYPES.to_string(), write_xml(&content_types)?);
    updates.insert(part.clone(), minimal_worksheet_xml()?);
    rewrite_parts(path, &updates, Vec::<String>::new())?;
    Ok(part)
}

pub fn move_sheet(
    path: &Path,
    name: &str,
    before: Option<&str>,
    after: Option<&str>,
) -> Result<()> {
    if before.is_some() == after.is_some() {
        bail!("provide exactly one of --before or --after");
    }
    if before == Some(name) || after == Some(name) {
        return Ok(());
    }
    let mut package = XlsxPackage::open(path)?;
    let old_order = sheet_names_in_order(&package.workbook)?;
    let sheets_idx = child_index(&package.workbook, "sheets")
        .ok_or_else(|| anyhow!("workbook has no sheets"))?;
    let sheets = element_child_mut_at(&mut package.workbook, sheets_idx)?;
    let sheet_pos = sheets
        .children
        .iter()
        .position(|node| matches!(node, XMLNode::Element(el) if el.name == "sheet" && attr(el, "name") == Some(name)));
    let Some(sheet_pos) = sheet_pos else {
        bail!("sheet '{name}' not found");
    };
    let node = sheets.children.remove(sheet_pos);
    if let XMLNode::Element(sheet) = node {
        insert_sheet_element_into_sheets(sheets, sheet, before, after)?;
    } else {
        unreachable!("position only matches sheet element");
    }
    let new_order = sheet_names_in_order(&package.workbook)?;
    update_sheet_index_references(&mut package.workbook, &old_order, &new_order)?;
    let mut updates = PartMap::new();
    updates.insert(WORKBOOK.to_string(), write_xml(&package.workbook)?);
    rewrite_parts(path, &updates, Vec::<String>::new())?;
    Ok(())
}

#[derive(Debug, Clone)]
struct SheetTarget {
    part: String,
}

struct XlsxPackage<'a> {
    path: &'a Path,
    workbook: Element,
    rels: Element,
    sheets: BTreeMap<String, SheetTarget>,
}

impl<'a> XlsxPackage<'a> {
    fn open(path: &'a Path) -> Result<Self> {
        let workbook = parse_xml(&read_part(path, WORKBOOK)?)?;
        let rels = parse_xml(&read_part(path, WORKBOOK_RELS)?)?;
        let mut rid_to_target = BTreeMap::new();
        for rel in children(&rels, "Relationship") {
            if let (Some(rid), Some(target)) = (attr(rel, "Id"), attr(rel, "Target")) {
                rid_to_target.insert(
                    rid.to_string(),
                    resolve_relationship_target(WORKBOOK, target),
                );
            }
        }
        let sheets_el =
            child(&workbook, "sheets").ok_or_else(|| anyhow!("workbook has no sheets"))?;
        let mut sheets = BTreeMap::new();
        for sheet in children(sheets_el, "sheet") {
            let Some(name) = attr(sheet, "name") else {
                continue;
            };
            let Some(rid) = attr(sheet, "r:id").or_else(|| attr(sheet, "id")) else {
                continue;
            };
            if let Some(part) = rid_to_target.get(rid) {
                sheets.insert(name.to_string(), SheetTarget { part: part.clone() });
            }
        }
        Ok(Self {
            path,
            workbook,
            rels,
            sheets,
        })
    }

    fn read_xml(&self, name: &str) -> Result<Element> {
        parse_xml(&read_part(self.path, name)?).with_context(|| format!("parse XML part: {name}"))
    }
}

fn read_range_values(
    range: &calamine::Range<Data>,
    range_ref: Option<&str>,
) -> Result<Vec<Vec<Data>>> {
    let (start_row, start_col, end_row, end_col) = if let Some(range_ref) = range_ref {
        let bounds = range_bounds(range_ref)?;
        (
            bounds.min_row.saturating_sub(1),
            bounds.min_col.saturating_sub(1),
            bounds.max_row.saturating_sub(1),
            bounds.max_col.saturating_sub(1),
        )
    } else {
        let height = range.height();
        let width = range.width();
        if height == 0 || width == 0 {
            return Ok(Vec::new());
        }
        (0, 0, height - 1, width - 1)
    };
    let mut rows = Vec::new();
    for row_idx in start_row..=end_row {
        let mut row = Vec::new();
        for col_idx in start_col..=end_col {
            row.push(
                range
                    .get((row_idx, col_idx))
                    .cloned()
                    .unwrap_or(Data::Empty),
            );
        }
        rows.push(row);
    }
    Ok(rows)
}

fn compact_display_rows(
    rows: Vec<Vec<Data>>,
    show_empty_rows: bool,
    show_empty_cols: bool,
) -> Vec<Vec<Data>> {
    let filtered = if show_empty_rows {
        rows
    } else {
        rows.into_iter()
            .filter(|row| row.iter().any(|value| !is_blank_value(value)))
            .collect::<Vec<_>>()
    };
    if filtered.is_empty() || show_empty_cols {
        return filtered;
    }
    let max_cols = filtered.iter().map(Vec::len).max().unwrap_or(0);
    let keep_cols = (0..max_cols)
        .filter(|col| {
            filtered
                .iter()
                .any(|row| row.get(*col).is_some_and(|value| !is_blank_value(value)))
        })
        .collect::<Vec<_>>();
    filtered
        .into_iter()
        .map(|row| {
            keep_cols
                .iter()
                .map(|col| row.get(*col).cloned().unwrap_or(Data::Empty))
                .collect()
        })
        .collect()
}

fn print_read_result(result: &ReadResult) {
    println!("Sheet: {}", result.sheet);
    println!("Dimensions: {}", result.dimensions);
    println!();
    if result.rows.is_empty() {
        println!("(empty sheet)");
    } else {
        print_markdown_table(&result.rows);
    }
}

fn print_markdown_table(rows: &[Vec<String>]) {
    if rows.is_empty() {
        println!("(no data)");
        return;
    }
    let max_cols = rows.iter().map(Vec::len).max().unwrap_or(0);
    let mut widths = vec![3usize; max_cols];
    for row in rows {
        for (idx, value) in row.iter().enumerate() {
            widths[idx] = widths[idx].max(value.len()).min(50);
        }
    }
    let mut normalized = rows.to_vec();
    for row in &mut normalized {
        row.resize(max_cols, String::new());
    }
    let header = &normalized[0];
    println!(
        "| {} |",
        header
            .iter()
            .enumerate()
            .map(|(idx, value)| pad_cell(value, widths[idx]))
            .collect::<Vec<_>>()
            .join(" | ")
    );
    println!(
        "| {} |",
        widths
            .iter()
            .map(|width| "-".repeat(*width))
            .collect::<Vec<_>>()
            .join(" | ")
    );
    for row in normalized.iter().skip(1) {
        println!(
            "| {} |",
            row.iter()
                .enumerate()
                .map(|(idx, value)| pad_cell(value, widths[idx]))
                .collect::<Vec<_>>()
                .join(" | ")
        );
    }
}

fn pad_cell(value: &str, width: usize) -> String {
    let truncated = if value.len() > width {
        &value[..width]
    } else {
        value
    };
    format!("{truncated:<width$}")
}

fn format_cell_value(value: &Data) -> String {
    match value {
        Data::Empty => String::new(),
        Data::String(s) => s.clone(),
        Data::Float(v) => {
            if v.fract() == 0.0 {
                format!("{v:.0}")
            } else {
                format!("{v:.2}")
            }
        }
        Data::Int(v) => v.to_string(),
        Data::Bool(v) => v.to_string(),
        Data::DateTime(v) => v.to_string(),
        Data::DateTimeIso(v) => v.clone(),
        Data::DurationIso(v) => v.clone(),
        Data::Error(v) => v.to_string(),
    }
}

fn is_blank_value(value: &Data) -> bool {
    matches!(value, Data::Empty) || matches!(value, Data::String(s) if s.is_empty())
}

fn dimensions_for_rows(rows: &[Vec<Data>]) -> String {
    if rows.is_empty() {
        "A1".to_string()
    } else {
        let rows_len = rows.len();
        let cols_len = rows.iter().map(Vec::len).max().unwrap_or(1);
        let end = cell_ref(cols_len, rows_len);
        if end == "A1" {
            "A1".to_string()
        } else {
            format!("A1:{end}")
        }
    }
}

fn print_json<T: Serialize>(value: &T) -> Result<()> {
    println!("{}", serde_json::to_string_pretty(value)?);
    Ok(())
}

#[derive(Debug, Clone, Copy)]
struct RangeBounds {
    min_col: usize,
    min_row: usize,
    max_col: usize,
    max_row: usize,
}

fn range_bounds(input: &str) -> Result<RangeBounds> {
    let (start, end) = input.split_once(':').unwrap_or((input, input));
    let (min_col, min_row) = split_cell_ref_or_axis(start)?;
    let (max_col, max_row) = split_cell_ref_or_axis(end)?;
    Ok(RangeBounds {
        min_col: min_col.min(max_col),
        min_row: min_row.min(max_row),
        max_col: min_col.max(max_col),
        max_row: min_row.max(max_row),
    })
}

fn split_cell_ref_or_axis(input: &str) -> Result<(usize, usize)> {
    if input.chars().all(|ch| ch.is_ascii_alphabetic()) {
        return Ok((column_index(input)?, 1));
    }
    if input.chars().all(|ch| ch.is_ascii_digit()) {
        return Ok((1, input.parse::<usize>()?));
    }
    split_cell_ref(input)
}

fn split_cell_ref(input: &str) -> Result<(usize, usize)> {
    let cleaned = input.replace('$', "");
    let re = Regex::new(r"^([A-Za-z]{1,3})([1-9][0-9]*)$").expect("valid regex");
    let caps = re
        .captures(&cleaned)
        .ok_or_else(|| anyhow!("invalid cell reference: {input}"))?;
    let col = column_index(&caps[1])?;
    let row = caps[2].parse::<usize>()?;
    Ok((col, row))
}

fn column_index(letters: &str) -> Result<usize> {
    let mut result = 0usize;
    for ch in letters.chars() {
        if !ch.is_ascii_alphabetic() {
            bail!("invalid column reference: {letters}");
        }
        result = result * 26 + (ch.to_ascii_uppercase() as u8 - b'A' + 1) as usize;
    }
    Ok(result)
}

fn column_letter(mut index: usize) -> String {
    let mut chars = Vec::new();
    while index > 0 {
        index -= 1;
        chars.push((b'A' + (index % 26) as u8) as char);
        index /= 26;
    }
    chars.iter().rev().collect()
}

fn cell_ref(col: usize, row: usize) -> String {
    format!("{}{}", column_letter(col), row)
}

fn cell_sort_key(reference: &str) -> Result<(usize, usize)> {
    let (col, row) = split_cell_ref(reference)?;
    Ok((row, col))
}

fn iter_target_cells(target: &str, value: &Value) -> Result<Vec<(usize, usize, Value)>> {
    let bounds = range_bounds(target)?;
    let mut out = Vec::new();
    match value {
        Value::Array(values) if values.first().is_some_and(Value::is_array) => {
            for (row_offset, row_value) in values.iter().enumerate() {
                let Some(row) = row_value.as_array() else {
                    bail!("matrix value mixes row arrays and scalar values");
                };
                for (col_offset, item) in row.iter().enumerate() {
                    out.push((
                        bounds.min_row + row_offset,
                        bounds.min_col + col_offset,
                        item.clone(),
                    ));
                }
            }
        }
        Value::Array(values) => {
            for (col_offset, item) in values.iter().enumerate() {
                out.push((bounds.min_row, bounds.min_col + col_offset, item.clone()));
            }
        }
        scalar => {
            for row in bounds.min_row..=bounds.max_row {
                for col in bounds.min_col..=bounds.max_col {
                    out.push((row, col, scalar.clone()));
                }
            }
        }
    }
    Ok(out)
}

fn ensure_sheet_data(worksheet: &mut Element) -> usize {
    if let Some(idx) = child_index(worksheet, "sheetData") {
        return idx;
    }
    let insert_index = worksheet
        .children
        .iter()
        .position(|node| {
            matches!(
                node,
                XMLNode::Element(el)
                    if matches!(
                        el.name.as_str(),
                        "sheetProtection" | "protectedRanges" | "autoFilter" | "sortState" | "dataConsolidate"
                    )
            )
        })
        .unwrap_or(worksheet.children.len());
    worksheet
        .children
        .insert(insert_index, XMLNode::Element(Element::new("sheetData")));
    insert_index
}

fn ensure_row(sheet_data: &mut Element, row_idx: usize) -> usize {
    for (idx, node) in sheet_data.children.iter().enumerate() {
        let XMLNode::Element(row) = node else {
            continue;
        };
        if row.name != "row" {
            continue;
        }
        let current = attr(row, "r")
            .and_then(|v| v.parse::<usize>().ok())
            .unwrap_or(0);
        if current == row_idx {
            return idx;
        }
        if current > row_idx {
            let mut row = Element::new("row");
            row.attributes.insert("r".to_string(), row_idx.to_string());
            sheet_data.children.insert(idx, XMLNode::Element(row));
            return idx;
        }
    }
    let mut row = Element::new("row");
    row.attributes.insert("r".to_string(), row_idx.to_string());
    sheet_data.children.push(XMLNode::Element(row));
    sheet_data.children.len() - 1
}

fn ensure_cell(row: &mut Element, row_idx: usize, col_idx: usize) -> Result<usize> {
    let reference = cell_ref(col_idx, row_idx);
    for (idx, node) in row.children.iter().enumerate() {
        let XMLNode::Element(cell) = node else {
            continue;
        };
        if cell.name != "c" {
            continue;
        }
        let Some(current_ref) = attr(cell, "r") else {
            continue;
        };
        if current_ref == reference {
            return Ok(idx);
        }
        if cell_sort_key(current_ref)? > (row_idx, col_idx) {
            let mut cell = Element::new("c");
            cell.attributes.insert("r".to_string(), reference);
            row.children.insert(idx, XMLNode::Element(cell));
            return Ok(idx);
        }
    }
    let mut cell = Element::new("c");
    cell.attributes.insert("r".to_string(), reference);
    row.children.push(XMLNode::Element(cell));
    Ok(row.children.len() - 1)
}

fn set_cell_value(
    cell: &mut Element,
    value: &Value,
    value_type: Option<ValueType>,
    date1904: bool,
) -> Result<()> {
    clear_cell_contents(cell);
    if matches!(value, Value::Null) {
        cell.attributes.remove("t");
        return Ok(());
    }
    let coerced = coerce_value(value, value_type, date1904)?;
    match coerced {
        CoercedValue::Formula(formula) => {
            cell.attributes.remove("t");
            let mut f = Element::new("f");
            f.children
                .push(XMLNode::Text(formula.trim_start_matches('=').to_string()));
            cell.children.push(XMLNode::Element(f));
        }
        CoercedValue::Bool(value) => {
            cell.attributes.insert("t".to_string(), "b".to_string());
            push_text_child(cell, "v", if value { "1" } else { "0" });
        }
        CoercedValue::Number(value) => {
            cell.attributes.remove("t");
            push_text_child(cell, "v", &value);
        }
        CoercedValue::String(value) => {
            cell.attributes
                .insert("t".to_string(), "inlineStr".to_string());
            let mut inline = Element::new("is");
            let mut text = Element::new("t");
            if value.starts_with(' ') || value.ends_with(' ') || value.contains('\n') {
                text.attributes
                    .insert("xml:space".to_string(), "preserve".to_string());
            }
            text.children.push(XMLNode::Text(value));
            inline.children.push(XMLNode::Element(text));
            cell.children.push(XMLNode::Element(inline));
        }
    }
    Ok(())
}

enum CoercedValue {
    Formula(String),
    Bool(bool),
    Number(String),
    String(String),
}

fn coerce_value(
    value: &Value,
    value_type: Option<ValueType>,
    date1904: bool,
) -> Result<CoercedValue> {
    if let Some(value_type) = value_type {
        let text = value_to_plain_text(value);
        return match value_type {
            ValueType::Text | ValueType::String => Ok(CoercedValue::String(text)),
            ValueType::Date | ValueType::Datetime => {
                let serial = parse_excel_datetime_serial(&text, date1904)?;
                Ok(CoercedValue::Number(format_number(serial)))
            }
            ValueType::Int | ValueType::Integer => {
                Ok(CoercedValue::Number(text.parse::<i64>()?.to_string()))
            }
            ValueType::Float | ValueType::Number | ValueType::Numeric => {
                Ok(CoercedValue::Number(format_number(text.parse::<f64>()?)))
            }
            ValueType::Bool | ValueType::Boolean => Ok(CoercedValue::Bool(parse_bool(&text)?)),
        };
    }
    match value {
        Value::Bool(value) => Ok(CoercedValue::Bool(*value)),
        Value::Number(value) => Ok(CoercedValue::Number(value.to_string())),
        Value::String(text) if text.starts_with('=') => Ok(CoercedValue::Formula(text.clone())),
        Value::String(text) => {
            if let Ok(int) = text.parse::<i64>() {
                return Ok(CoercedValue::Number(int.to_string()));
            }
            if let Ok(float) = text.parse::<f64>() {
                return Ok(CoercedValue::Number(format_number(float)));
            }
            Ok(CoercedValue::String(text.clone()))
        }
        other => Ok(CoercedValue::String(value_to_plain_text(other))),
    }
}

fn value_to_plain_text(value: &Value) -> String {
    match value {
        Value::String(text) => text.clone(),
        Value::Null => String::new(),
        other => other.to_string(),
    }
}

fn parse_bool(text: &str) -> Result<bool> {
    match text.trim().to_ascii_lowercase().as_str() {
        "true" | "1" | "yes" => Ok(true),
        "false" | "0" | "no" => Ok(false),
        _ => bail!("invalid boolean value: {text}"),
    }
}

fn parse_excel_datetime_serial(text: &str, date1904: bool) -> Result<f64> {
    let dt = if let Ok(dt) = NaiveDateTime::parse_from_str(text, "%Y-%m-%dT%H:%M:%S") {
        dt
    } else if let Ok(dt) = NaiveDateTime::parse_from_str(text, "%Y-%m-%d %H:%M:%S") {
        dt
    } else if let Ok(dt) = NaiveDateTime::parse_from_str(text, "%Y-%m-%dT%H:%M") {
        dt
    } else if let Ok(dt) = NaiveDateTime::parse_from_str(text, "%Y-%m-%d %H:%M") {
        dt
    } else {
        NaiveDate::parse_from_str(text, "%Y-%m-%d")?
            .and_hms_opt(0, 0, 0)
            .ok_or_else(|| anyhow!("invalid date value: {text}"))?
    };
    let base = if date1904 {
        NaiveDate::from_ymd_opt(1904, 1, 1).unwrap()
    } else {
        NaiveDate::from_ymd_opt(1899, 12, 30).unwrap()
    }
    .and_hms_opt(0, 0, 0)
    .unwrap();
    let duration = dt.signed_duration_since(base);
    Ok(duration.num_seconds() as f64 / 86_400.0)
}

fn format_number(value: f64) -> String {
    if value.fract() == 0.0 {
        format!("{value:.0}")
    } else {
        value.to_string()
    }
}

fn clear_cell_contents(cell: &mut Element) {
    cell.children.retain(|node| {
        !matches!(node, XMLNode::Element(el) if matches!(el.name.as_str(), "f" | "v" | "is"))
    });
    cell.attributes.remove("t");
}

fn push_text_child(parent: &mut Element, name: &str, text: &str) {
    let mut el = Element::new(name);
    el.children.push(XMLNode::Text(text.to_string()));
    parent.children.push(XMLNode::Element(el));
}

fn update_worksheet_dimension(worksheet: &mut Element) -> Result<()> {
    let Some(sheet_data) = child(worksheet, "sheetData") else {
        return Ok(());
    };
    let mut refs = Vec::new();
    for row in children(sheet_data, "row") {
        for cell in children(row, "c") {
            if let Some(reference) = attr(cell, "r") {
                refs.push(reference.to_string());
            }
        }
    }
    let dim_ref = if refs.is_empty() {
        "A1".to_string()
    } else {
        let mut min_col = usize::MAX;
        let mut min_row = usize::MAX;
        let mut max_col = 0;
        let mut max_row = 0;
        for reference in refs {
            let (col, row) = split_cell_ref(&reference)?;
            min_col = min_col.min(col);
            min_row = min_row.min(row);
            max_col = max_col.max(col);
            max_row = max_row.max(row);
        }
        let start = cell_ref(min_col, min_row);
        let end = cell_ref(max_col, max_row);
        if start == end {
            start
        } else {
            format!("{start}:{end}")
        }
    };
    if let Some(idx) = child_index(worksheet, "dimension") {
        let dimension = element_child_mut_at(worksheet, idx)?;
        dimension.attributes.insert("ref".to_string(), dim_ref);
    } else {
        let mut dimension = Element::new("dimension");
        dimension.attributes.insert("ref".to_string(), dim_ref);
        worksheet.children.insert(0, XMLNode::Element(dimension));
    }
    Ok(())
}

fn workbook_uses_1904_dates(workbook: &Element) -> bool {
    child(workbook, "workbookPr")
        .and_then(|el| attr(el, "date1904"))
        .is_some_and(|value| matches!(value, "1" | "true" | "True"))
}

fn set_workbook_recalc_flags(workbook: &mut Element) {
    let calc_pr = if let Some(idx) = child_index(workbook, "calcPr") {
        element_child_mut_at(workbook, idx).expect("existing calcPr is element")
    } else {
        workbook
            .children
            .push(XMLNode::Element(Element::new("calcPr")));
        element_child_mut_at(workbook, workbook.children.len() - 1).expect("new calcPr is element")
    };
    calc_pr
        .attributes
        .insert("calcMode".to_string(), "auto".to_string());
    calc_pr
        .attributes
        .insert("fullCalcOnLoad".to_string(), "1".to_string());
    calc_pr
        .attributes
        .insert("forceFullCalc".to_string(), "1".to_string());
}

fn remove_calc_chain_relationships(rels: &mut Element) -> Vec<String> {
    let mut removed = Vec::new();
    rels.children.retain(|node| {
        let XMLNode::Element(rel) = node else {
            return true;
        };
        if attr(rel, "Type") == Some(CALC_CHAIN_REL_TYPE) {
            if let Some(target) = attr(rel, "Target") {
                removed.push(resolve_relationship_target(WORKBOOK, target));
            }
            false
        } else {
            true
        }
    });
    removed
}

fn validate_sheet_name(name: &str) -> Result<()> {
    if name.trim().is_empty() {
        bail!("sheet name cannot be blank");
    }
    if name.chars().count() > 31 {
        bail!("sheet name cannot be longer than 31 characters");
    }
    let bad = name
        .chars()
        .filter(|ch| matches!(ch, '[' | ']' | ':' | '*' | '?' | '/' | '\\'))
        .collect::<BTreeSet<_>>();
    if !bad.is_empty() {
        bail!(
            "sheet name contains invalid character(s): {}",
            bad.into_iter().collect::<String>()
        );
    }
    if name.starts_with('\'') || name.ends_with('\'') {
        bail!("sheet name cannot start or end with an apostrophe");
    }
    Ok(())
}

fn next_sheet_part(path: &Path) -> Result<String> {
    let existing = crate::ooxml::list_parts(path)?
        .into_iter()
        .map(|part| part.name)
        .collect::<BTreeSet<_>>();
    let mut index = 1;
    loop {
        let candidate = format!("xl/worksheets/sheet{index}.xml");
        if !existing.contains(&candidate) {
            return Ok(candidate);
        }
        index += 1;
    }
}

fn next_relationship_id(rels: &Element) -> String {
    let used = children(rels, "Relationship")
        .filter_map(|rel| attr(rel, "Id").map(ToString::to_string))
        .collect::<BTreeSet<_>>();
    let mut max = 0;
    for id in &used {
        if let Some(number) = id
            .strip_prefix("rId")
            .and_then(|text| text.parse::<usize>().ok())
        {
            max = max.max(number);
        }
    }
    let mut candidate = max + 1;
    while used.contains(&format!("rId{candidate}")) {
        candidate += 1;
    }
    format!("rId{candidate}")
}

fn next_sheet_id(workbook: &Element) -> Result<usize> {
    let sheets = child(workbook, "sheets").ok_or_else(|| anyhow!("workbook has no sheets"))?;
    let mut max_id = 0;
    for sheet in children(sheets, "sheet") {
        if let Some(id) = attr(sheet, "sheetId").and_then(|id| id.parse::<usize>().ok()) {
            max_id = max_id.max(id);
        }
    }
    Ok(max_id + 1)
}

fn sheet_names_in_order(workbook: &Element) -> Result<Vec<String>> {
    let sheets = child(workbook, "sheets").ok_or_else(|| anyhow!("workbook has no sheets"))?;
    Ok(children(sheets, "sheet")
        .filter_map(|sheet| attr(sheet, "name").map(ToString::to_string))
        .collect())
}

fn insert_sheet_element(
    workbook: &mut Element,
    sheet: Element,
    before: Option<&str>,
    after: Option<&str>,
) -> Result<()> {
    let sheets_idx =
        child_index(workbook, "sheets").ok_or_else(|| anyhow!("workbook has no sheets"))?;
    let sheets = element_child_mut_at(workbook, sheets_idx)?;
    insert_sheet_element_into_sheets(sheets, sheet, before, after)
}

fn insert_sheet_element_into_sheets(
    sheets: &mut Element,
    sheet: Element,
    before: Option<&str>,
    after: Option<&str>,
) -> Result<()> {
    if let Some(before) = before {
        let idx = sheets
            .children
            .iter()
            .position(|node| matches!(node, XMLNode::Element(el) if el.name == "sheet" && attr(el, "name") == Some(before)))
            .ok_or_else(|| anyhow!("insertion sheet '{before}' not found"))?;
        sheets.children.insert(idx, XMLNode::Element(sheet));
        return Ok(());
    }
    if let Some(after) = after {
        let idx = sheets
            .children
            .iter()
            .position(|node| matches!(node, XMLNode::Element(el) if el.name == "sheet" && attr(el, "name") == Some(after)))
            .ok_or_else(|| anyhow!("insertion sheet '{after}' not found"))?;
        sheets.children.insert(idx + 1, XMLNode::Element(sheet));
        return Ok(());
    }
    sheets.children.push(XMLNode::Element(sheet));
    Ok(())
}

fn update_sheet_index_references(
    workbook: &mut Element,
    old_order: &[String],
    new_order: &[String],
) -> Result<()> {
    if old_order == new_order {
        return Ok(());
    }
    let new_indexes = new_order
        .iter()
        .enumerate()
        .map(|(idx, name)| (name.as_str(), idx))
        .collect::<BTreeMap<_, _>>();
    let old_to_new = old_order
        .iter()
        .enumerate()
        .filter_map(|(idx, name)| {
            new_indexes
                .get(name.as_str())
                .copied()
                .map(|new| (idx, new))
        })
        .collect::<BTreeMap<_, _>>();
    if let Some(defined_names) = child_mut(workbook, "definedNames") {
        for defined_name in children_mut_named(defined_names, "definedName") {
            let Some(local_sheet_id) =
                attr(defined_name, "localSheetId").and_then(|value| value.parse::<usize>().ok())
            else {
                continue;
            };
            if let Some(new_id) = old_to_new.get(&local_sheet_id) {
                defined_name
                    .attributes
                    .insert("localSheetId".to_string(), new_id.to_string());
            }
        }
    }
    if let Some(book_views) = child_mut(workbook, "bookViews") {
        for view in children_mut_named(book_views, "workbookView") {
            for attr_name in ["activeTab", "firstSheet"] {
                let Some(old_value) =
                    attr(view, attr_name).and_then(|value| value.parse::<usize>().ok())
                else {
                    continue;
                };
                if let Some(new_value) = old_to_new.get(&old_value) {
                    view.attributes
                        .insert(attr_name.to_string(), new_value.to_string());
                }
            }
        }
    }
    Ok(())
}

fn ensure_content_type_override(content_types: &mut Element, part: &str, content_type: &str) {
    let part_name = format!("/{}", part.trim_start_matches('/'));
    for node in &mut content_types.children {
        let XMLNode::Element(el) = node else {
            continue;
        };
        if el.name == "Override" && attr(el, "PartName") == Some(part_name.as_str()) {
            el.attributes
                .insert("ContentType".to_string(), content_type.to_string());
            return;
        }
    }
    let mut override_el = Element::new("Override");
    override_el
        .attributes
        .insert("PartName".to_string(), part_name);
    override_el
        .attributes
        .insert("ContentType".to_string(), content_type.to_string());
    content_types.children.push(XMLNode::Element(override_el));
}

fn remove_content_type_override(content_types: &mut Element, part: &str) {
    let part_name = format!("/{}", part.trim_start_matches('/'));
    content_types.children.retain(|node| {
        !matches!(node, XMLNode::Element(el) if el.name == "Override" && attr(el, "PartName") == Some(part_name.as_str()))
    });
}

fn minimal_worksheet_xml() -> Result<Vec<u8>> {
    parse_xml(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <sheetData/>
</worksheet>"#,
    )
    .and_then(|xml| write_xml(&xml))
}

fn child_index(element: &Element, name: &str) -> Option<usize> {
    element
        .children
        .iter()
        .position(|node| matches!(node, XMLNode::Element(el) if el.name == name))
}

fn element_child_mut_at(parent: &mut Element, idx: usize) -> Result<&mut Element> {
    match parent.children.get_mut(idx) {
        Some(XMLNode::Element(el)) => Ok(el),
        _ => bail!("expected XML element at child index {idx}"),
    }
}

fn attr<'a>(element: &'a Element, name: &str) -> Option<&'a str> {
    element
        .attributes
        .get(name)
        .or_else(|| {
            name.split_once(':')
                .and_then(|(_, local)| element.attributes.get(local))
        })
        .map(String::as_str)
}

fn element_text_all(element: &Element) -> String {
    let mut out = String::new();
    for child in &element.children {
        match child {
            XMLNode::Text(text) | XMLNode::CData(text) => out.push_str(text),
            XMLNode::Element(el) => out.push_str(&element_text_all(el)),
            _ => {}
        }
    }
    out
}

fn children_mut_named<'a>(
    element: &'a mut Element,
    name: &'a str,
) -> impl Iterator<Item = &'a mut Element> + 'a {
    element
        .children
        .iter_mut()
        .filter_map(move |node| match node {
            XMLNode::Element(el) if el.name == name => Some(el),
            _ => None,
        })
}

pub fn extract_sheet_text(path: &Path, sheet: Option<&str>) -> Result<String> {
    let result = read(&ReadArgs {
        file: path.to_path_buf(),
        sheet: sheet.map(ToString::to_string),
        range: None,
        show_empty_rows: false,
        show_empty_cols: false,
        json: false,
    })?;
    let mut out = String::new();
    for row in result.rows {
        let line = row
            .into_iter()
            .filter(|cell| !cell.is_empty())
            .collect::<Vec<_>>()
            .join("\t");
        if !line.is_empty() {
            out.push_str(&line);
            out.push('\n');
        }
    }
    Ok(out)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ooxml::{list_parts, read_part_to_string, write_new_package};

    #[test]
    fn direct_edit_preserves_validation_and_conditional_formatting() -> Result<()> {
        let dir = tempfile::tempdir()?;
        let path = dir.path().join("book.xlsx");
        write_new_package(&path, &minimal_xlsx_parts())?;

        let result = edit(&EditArgs {
            file: path.clone(),
            sheet: Some("Sheet1".to_string()),
            cell: Some("A1".to_string()),
            range: None,
            value: Some("replacement".to_string()),
            clear: false,
            edits: None,
            edits_file: None,
            value_type: None,
            allow_protected: false,
            json: true,
        })?;

        assert_eq!(result.target_count, 1);
        let sheet = read_part_to_string(&path, "xl/worksheets/sheet1.xml")?;
        assert!(sheet.contains("conditionalFormatting"));
        assert!(sheet.contains("dataValidations"));
        assert!(sheet.contains("replacement"));
        Ok(())
    }

    #[test]
    fn formula_edit_removes_calc_chain_and_marks_recalc() -> Result<()> {
        let dir = tempfile::tempdir()?;
        let path = dir.path().join("book.xlsx");
        write_new_package(&path, &minimal_xlsx_parts())?;

        let result = edit(&EditArgs {
            file: path.clone(),
            sheet: Some("Sheet1".to_string()),
            cell: Some("B2".to_string()),
            range: None,
            value: Some("=SUM(A1:A2)".to_string()),
            clear: false,
            edits: None,
            edits_file: None,
            value_type: None,
            allow_protected: false,
            json: true,
        })?;

        assert!(result.formula_recalc_marked);
        let workbook = read_part_to_string(&path, "xl/workbook.xml")?;
        assert!(workbook.contains("fullCalcOnLoad"));
        let rels = read_part_to_string(&path, "xl/_rels/workbook.xml.rels")?;
        assert!(!rels.contains("calcChain"));
        let parts = list_parts(&path)?;
        assert!(!parts.iter().any(|part| part.name == "xl/calcChain.xml"));
        Ok(())
    }

    #[test]
    fn add_and_move_sheet_updates_workbook_order() -> Result<()> {
        let dir = tempfile::tempdir()?;
        let path = dir.path().join("book.xlsx");
        write_new_package(&path, &minimal_xlsx_parts())?;

        add_sheet(&path, "Support", None, Some("Sheet1"))?;
        move_sheet(&path, "Support", Some("Sheet1"), None)?;

        let workbook = read_part_to_string(&path, "xl/workbook.xml")?;
        let support_pos = workbook.find("Support").expect("Support sheet");
        let sheet1_pos = workbook.find("Sheet1").expect("Sheet1 sheet");
        assert!(support_pos < sheet1_pos);
        Ok(())
    }

    #[test]
    fn inspect_reports_validation_and_conditional_formatting() -> Result<()> {
        let dir = tempfile::tempdir()?;
        let path = dir.path().join("book.xlsx");
        write_new_package(&path, &minimal_xlsx_parts())?;

        let result = inspect(&InspectArgs {
            file: path,
            sheet: Some("Sheet1".to_string()),
        })?;

        assert_eq!(result.sheets.len(), 1);
        assert_eq!(
            result.sheets[0].conditional_formatting_ranges,
            vec!["A1:A10"]
        );
        assert_eq!(result.sheets[0].data_validation_ranges, vec!["B1:B10"]);
        Ok(())
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
<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>
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
</workbook>"#
                .to_vec(),
        );
        parts.insert(
            "xl/_rels/workbook.xml.rels".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/>
</Relationships>"#
                .to_vec(),
        );
        parts.insert(
            "xl/worksheets/sheet1.xml".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<dimension ref="A1:B2"/>
<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>old</t></is></c></row></sheetData>
<conditionalFormatting sqref="A1:A10"><cfRule type="cellIs" priority="1" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>
<dataValidations count="1"><dataValidation type="whole" allowBlank="1" sqref="B1:B10"><formula1>1</formula1><formula2>10</formula2></dataValidation></dataValidations>
</worksheet>"#
                .to_vec(),
        );
        parts.insert(
            "xl/calcChain.xml".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#.to_vec(),
        );
        parts
    }
}
