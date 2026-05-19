use crate::ooxml::{
    PartMap, child, children, escaped, parse_xml, read_part, rewrite_parts, write_new_package,
};
use anyhow::{Context, Result, anyhow, bail};
use clap::{Args, Subcommand, ValueEnum};
use serde::{Deserialize, Serialize};
use std::fs;
use std::path::{Path, PathBuf};
use xmltree::{Element, XMLNode};

#[derive(Debug, Args)]
pub struct DocxArgs {
    #[command(subcommand)]
    command: DocxCommand,
}

#[derive(Debug, Subcommand)]
pub enum DocxCommand {
    /// Read document body as markdown.
    Read(DocxReadArgs),
    /// Find and replace text in Word XML parts or through Word COM.
    Replace(DocxReplaceArgs),
    /// Compose a new generic .docx from a JSON document spec.
    Compose(DocxComposeArgs),
}

#[derive(Debug, Args)]
pub struct DocxReadArgs {
    pub file: PathBuf,
    #[arg(long)]
    pub json: bool,
}

#[derive(Debug, Args)]
pub struct DocxReplaceArgs {
    pub file: PathBuf,
    #[arg(long)]
    pub find: Option<String>,
    #[arg(long)]
    pub replace: Option<String>,
    #[arg(long)]
    pub replacements: Option<String>,
    #[arg(long)]
    pub replacements_file: Option<PathBuf>,
    #[arg(short, long)]
    pub output: Option<PathBuf>,
    #[arg(long, value_enum, default_value_t = DocxReplaceEngine::Direct)]
    pub engine: DocxReplaceEngine,
}

#[derive(Debug, Clone, Copy, ValueEnum)]
pub enum DocxReplaceEngine {
    Direct,
    Com,
}

#[derive(Debug, Args)]
pub struct DocxComposeArgs {
    pub spec: PathBuf,
    #[arg(short, long)]
    pub output: PathBuf,
}

#[derive(Debug, Clone, Deserialize, Serialize)]
pub struct Replacement {
    pub find: String,
    pub replace: String,
}

#[derive(Debug, Serialize)]
pub struct ReplaceResult {
    pub file: PathBuf,
    pub counts: Vec<ReplacementCount>,
}

#[derive(Debug, Serialize)]
pub struct ReplacementCount {
    pub find: String,
    pub count: usize,
}

#[derive(Debug, Deserialize)]
pub struct ComposeSpec {
    #[serde(default)]
    pub meta: ComposeMeta,
    #[serde(default)]
    pub brand: BrandSpec,
    #[serde(default)]
    pub blocks: Vec<DocxBlock>,
    #[serde(default)]
    pub template: Option<PathBuf>,
}

#[derive(Debug, Default, Deserialize)]
pub struct ComposeMeta {
    pub title: Option<String>,
    pub subject: Option<String>,
    pub creator: Option<String>,
    pub footer_text: Option<String>,
}

#[derive(Debug, Clone, Deserialize)]
pub struct BrandSpec {
    #[serde(default = "default_font")]
    pub font_family: String,
    #[serde(default = "default_body_color")]
    pub body_color: String,
    #[serde(default = "default_accent_color")]
    pub accent_color: String,
    #[serde(default = "default_heading_color")]
    pub heading_color: String,
    #[serde(default = "default_muted_color")]
    pub muted_color: String,
    #[serde(default = "default_table_header_fill")]
    pub table_header_fill: String,
    #[serde(default = "default_alt_row_fill")]
    pub alt_row_fill: String,
}

impl Default for BrandSpec {
    fn default() -> Self {
        Self {
            font_family: default_font(),
            body_color: default_body_color(),
            accent_color: default_accent_color(),
            heading_color: default_heading_color(),
            muted_color: default_muted_color(),
            table_header_fill: default_table_header_fill(),
            alt_row_fill: default_alt_row_fill(),
        }
    }
}

#[derive(Debug, Deserialize)]
#[serde(tag = "type", rename_all = "snake_case")]
pub enum DocxBlock {
    TitlePage {
        client_name: Option<String>,
        title: String,
        subtitle: Option<String>,
        date_text: Option<String>,
        prepared_by: Option<String>,
    },
    Heading {
        text: String,
        #[serde(default = "default_heading_level")]
        level: u8,
        #[serde(default)]
        page_break: bool,
    },
    Body {
        text: String,
        #[serde(default)]
        bold: bool,
        #[serde(default)]
        footnote: Option<String>,
    },
    BodyRich {
        segments: Vec<RichSegment>,
    },
    Bullet {
        text: String,
        #[serde(default = "default_list_level")]
        level: u8,
        bold_prefix: Option<String>,
        #[serde(default)]
        footnote: Option<String>,
        /// Optional paragraph style override. Defaults to `Bullet{level}` (template's branded
        /// bullets, often square). Set to e.g. "ListBullet" for Word's default round bullets.
        #[serde(default)]
        style: Option<String>,
    },
    Numbered {
        text: String,
        #[serde(default = "default_number")]
        number: usize,
        bold_prefix: Option<String>,
        #[serde(default)]
        footnote: Option<String>,
    },
    Quote {
        text: String,
    },
    QuoteBlock {
        lines: Vec<String>,
    },
    Table {
        headers: Vec<String>,
        rows: Vec<Vec<CellContent>>,
        #[serde(default)]
        alt_row_shading: bool,
        #[serde(default)]
        first_col_bold: bool,
        #[serde(default)]
        col_widths_twips: Option<Vec<usize>>,
    },
    BorderlessTable {
        rows: Vec<Vec<CellContent>>,
        #[serde(default)]
        label_col: usize,
        #[serde(default)]
        col_widths_twips: Option<Vec<usize>>,
        /// Optional border style: "none" (default), "rows" (horizontal lines between rows),
        /// or "all" (full grid). Used to lay out memo header tables with subtle row separators.
        #[serde(default)]
        borders: Option<String>,
    },
    Divider,
    Spacer {
        #[serde(default = "default_spacer_twips")]
        height_twips: usize,
    },
    PageBreak,
}

#[derive(Debug, Deserialize)]
pub struct RichSegment {
    pub text: String,
    #[serde(default)]
    pub bold: bool,
    #[serde(default)]
    pub italic: bool,
    pub color: Option<String>,
}

/// A single table cell. Accepts either a plain string or a richer object with
/// per-cell styling. With `untagged` serde, JSON strings deserialize as `Text`
/// and JSON objects with a `text` field deserialize as `Rich`, so existing
/// specs that pass string cells keep working unchanged.
#[derive(Debug, Deserialize)]
#[serde(untagged)]
pub enum CellContent {
    Text(String),
    Rich(RichCell),
}

#[derive(Debug, Deserialize)]
pub struct RichCell {
    pub text: String,
    #[serde(default)]
    pub bold: bool,
    #[serde(default)]
    pub italic: bool,
    pub color: Option<String>,
}

impl CellContent {
    pub fn as_str(&self) -> &str {
        match self {
            CellContent::Text(s) => s.as_str(),
            CellContent::Rich(r) => r.text.as_str(),
        }
    }

    pub fn rich(&self) -> Option<&RichCell> {
        match self {
            CellContent::Rich(r) => Some(r),
            CellContent::Text(_) => None,
        }
    }
}

impl DocxArgs {
    pub fn run(self) -> Result<()> {
        match self.command {
            DocxCommand::Read(args) => {
                let markdown = read_docx_markdown(&args.file)?;
                if args.json {
                    println!(
                        "{}",
                        serde_json::to_string_pretty(&serde_json::json!({ "markdown": markdown }))?
                    );
                } else {
                    println!("{markdown}");
                }
                Ok(())
            }
            DocxCommand::Replace(args) => {
                let result = replace(&args)?;
                println!("{}", serde_json::to_string_pretty(&result)?);
                Ok(())
            }
            DocxCommand::Compose(args) => {
                compose_from_file(&args.spec, &args.output)?;
                println!("Saved: {}", args.output.display());
                Ok(())
            }
        }
    }
}

pub fn read_docx_markdown(path: &Path) -> Result<String> {
    let document = parse_xml(&read_part(path, "word/document.xml")?)
        .with_context(|| format!("parse word/document.xml from {}", path.display()))?;
    let body = child(&document, "body").ok_or_else(|| anyhow!("word/document.xml has no body"))?;
    let mut out = Vec::new();
    for node in &body.children {
        let XMLNode::Element(el) = node else {
            continue;
        };
        match el.name.as_str() {
            "p" => {
                let text = paragraph_text(el);
                if text.trim().is_empty() {
                    out.push(String::new());
                    continue;
                }
                if let Some(style) = paragraph_style(el) {
                    if let Some(level) = style
                        .strip_prefix("Heading")
                        .and_then(|s| s.parse::<usize>().ok())
                    {
                        out.push(format!("{} {}", "#".repeat(level.max(1)), text));
                    } else if style.contains("ListBullet") || style.contains("Bullet") {
                        out.push(format!("- {text}"));
                    } else if style.contains("ListNumber") || style.contains("Number") {
                        out.push(format!("1. {text}"));
                    } else {
                        out.push(text);
                    }
                } else {
                    out.push(text);
                }
            }
            "tbl" => {
                let rows = table_rows(el);
                if rows.is_empty() {
                    continue;
                }
                let max_cols = rows.iter().map(Vec::len).max().unwrap_or(0);
                out.push(String::new());
                out.push(format!(
                    "| {} |",
                    normalized_row(&rows[0], max_cols).join(" | ")
                ));
                out.push(format!("| {} |", vec!["---"; max_cols].join(" | ")));
                for row in rows.iter().skip(1) {
                    out.push(format!("| {} |", normalized_row(row, max_cols).join(" | ")));
                }
                out.push(String::new());
            }
            _ => {}
        }
    }
    let mut result = out.join("\n");
    while result.contains("\n\n\n") {
        result = result.replace("\n\n\n", "\n\n");
    }
    Ok(result.trim().to_string())
}

pub fn replace(args: &DocxReplaceArgs) -> Result<ReplaceResult> {
    let replacements = load_replacements(args)?;
    match args.engine {
        DocxReplaceEngine::Com => crate::wincom::word_replace(args, &replacements),
        DocxReplaceEngine::Direct => direct_replace(args, &replacements),
    }
}

fn direct_replace(args: &DocxReplaceArgs, replacements: &[Replacement]) -> Result<ReplaceResult> {
    if !args.file.exists() {
        bail!("file not found: {}", args.file.display());
    }
    let target = args.output.as_ref().unwrap_or(&args.file);
    if let Some(output) = &args.output {
        fs::copy(&args.file, output)
            .with_context(|| format!("copy {} to {}", args.file.display(), output.display()))?;
    }
    let candidate_parts = crate::ooxml::list_parts(target)?
        .into_iter()
        .map(|part| part.name)
        .filter(|name| {
            name == "word/document.xml"
                || name.starts_with("word/header")
                || name.starts_with("word/footer")
                || matches!(
                    name.as_str(),
                    "word/footnotes.xml" | "word/endnotes.xml" | "word/comments.xml"
                )
        })
        .collect::<Vec<_>>();
    let mut counts = replacements
        .iter()
        .map(|rep| ReplacementCount {
            find: rep.find.clone(),
            count: 0,
        })
        .collect::<Vec<_>>();
    let mut updates = PartMap::new();
    for part in candidate_parts {
        let raw = read_part(target, &part)?;
        let original = std::str::from_utf8(&raw)
            .with_context(|| format!("part is not valid UTF-8: {part}"))?
            .to_string();
        let (rewritten, touched) =
            replace_in_text_elements(&original, replacements, &mut counts);
        if touched {
            updates.insert(part, rewritten.into_bytes());
        }
    }
    rewrite_parts(target, &updates, Vec::<String>::new())?;
    Ok(ReplaceResult {
        file: target.to_path_buf(),
        counts,
    })
}

fn load_replacements(args: &DocxReplaceArgs) -> Result<Vec<Replacement>> {
    if args.replacements.is_some() && args.replacements_file.is_some() {
        bail!("provide --replacements or --replacements-file, not both");
    }
    if let Some(path) = &args.replacements_file {
        let text = fs::read_to_string(path)
            .with_context(|| format!("read replacements file: {}", path.display()))?;
        return serde_json::from_str(&text).context("parse replacements JSON file");
    }
    if let Some(text) = &args.replacements {
        return serde_json::from_str(text).context("parse --replacements JSON");
    }
    if let (Some(find), Some(replace)) = (&args.find, &args.replace) {
        return Ok(vec![Replacement {
            find: find.clone(),
            replace: replace.clone(),
        }]);
    }
    bail!("provide --find/--replace, --replacements JSON, or --replacements-file");
}

/// Direct text-node replacement that operates on raw OOXML bytes.
///
/// Walks the source string and finds every `<w:t ...>...</w:t>` element,
/// rewriting only the text content between the open and close tags. The
/// surrounding XML — namespace declarations, attribute order, `xml:space`
/// attributes, self-closing tags elsewhere, document declaration, and the
/// rest of the bytes — is preserved verbatim.
///
/// This is the safe alternative to a full `xmltree` parse/serialize round-trip
/// (which loses or rearranges OOXML quirks that Word relies on, producing
/// "the file is corrupt; would you like to recover?" prompts in Word).
fn replace_in_text_elements(
    source: &str,
    replacements: &[Replacement],
    counts: &mut [ReplacementCount],
) -> (String, bool) {
    let mut out = String::with_capacity(source.len());
    let mut touched = false;
    let mut cursor = 0;
    let bytes = source.as_bytes();

    while cursor < bytes.len() {
        // Find next "<w:t" — must be at a tag boundary with the next char
        // being whitespace, '>' or '/' so we don't match e.g. "<w:tbl".
        let Some(rel_start) = source[cursor..].find("<w:t") else {
            out.push_str(&source[cursor..]);
            break;
        };
        let tag_start = cursor + rel_start;
        let after_w_t = tag_start + 4;
        if after_w_t >= bytes.len() {
            out.push_str(&source[cursor..]);
            break;
        }
        let boundary = bytes[after_w_t];
        if !matches!(boundary, b' ' | b'\t' | b'\r' | b'\n' | b'>' | b'/') {
            // Not actually a <w:t> — could be <w:tbl>, <w:tc>, <w:tr>, etc.
            // Copy through the false-start and continue scanning.
            out.push_str(&source[cursor..after_w_t]);
            cursor = after_w_t;
            continue;
        }
        // Find the matching '>' that closes the start tag.
        let Some(open_end_rel) = source[after_w_t..].find('>') else {
            out.push_str(&source[cursor..]);
            break;
        };
        let open_end = after_w_t + open_end_rel;
        let open_tag_end = open_end + 1;

        // Self-closing `<w:t/>` — no content to replace.
        if open_end > 0 && bytes[open_end - 1] == b'/' {
            out.push_str(&source[cursor..open_tag_end]);
            cursor = open_tag_end;
            continue;
        }
        // Find the closing `</w:t>`.
        let Some(close_rel) = source[open_tag_end..].find("</w:t>") else {
            out.push_str(&source[cursor..]);
            break;
        };
        let close_start = open_tag_end + close_rel;
        let close_end = close_start + "</w:t>".len();

        // Emit everything up to and including the open tag.
        out.push_str(&source[cursor..open_tag_end]);

        // Modify only the inner text. The inner text is XML-escaped; we
        // operate on the decoded form so that find/replace strings match
        // user intent (e.g. find "&" works even when stored as "&amp;").
        let inner = &source[open_tag_end..close_start];
        let decoded = xml_unescape(inner);
        let mut new_decoded = decoded.clone();
        let mut local_touched = false;
        for (idx, rep) in replacements.iter().enumerate() {
            let n = new_decoded.matches(&rep.find).count();
            if n > 0 {
                new_decoded = new_decoded.replace(&rep.find, &rep.replace);
                counts[idx].count += n;
                local_touched = true;
            }
        }
        if local_touched {
            touched = true;
            out.push_str(&xml_escape(&new_decoded));
        } else {
            out.push_str(inner);
        }
        out.push_str("</w:t>");
        cursor = close_end;
    }
    (out, touched)
}

/// XML-escape characters that have special meaning in element text content.
fn xml_escape(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for ch in s.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            other => out.push(other),
        }
    }
    out
}

/// Reverse of `xml_escape` for the common entities Word produces in text nodes.
fn xml_unescape(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    let bytes = s.as_bytes();
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] == b'&' {
            if let Some(end) = s[i..].find(';') {
                let entity = &s[i..i + end + 1];
                match entity {
                    "&amp;" => {
                        out.push('&');
                        i += entity.len();
                        continue;
                    }
                    "&lt;" => {
                        out.push('<');
                        i += entity.len();
                        continue;
                    }
                    "&gt;" => {
                        out.push('>');
                        i += entity.len();
                        continue;
                    }
                    "&quot;" => {
                        out.push('"');
                        i += entity.len();
                        continue;
                    }
                    "&apos;" => {
                        out.push('\'');
                        i += entity.len();
                        continue;
                    }
                    other if other.starts_with("&#") => {
                        let body = &other[2..other.len() - 1];
                        let code = if let Some(hex) = body.strip_prefix('x').or_else(|| body.strip_prefix('X')) {
                            u32::from_str_radix(hex, 16).ok()
                        } else {
                            body.parse::<u32>().ok()
                        };
                        if let Some(c) = code.and_then(char::from_u32) {
                            out.push(c);
                            i += entity.len();
                            continue;
                        }
                    }
                    _ => {}
                }
            }
        }
        // Default: copy this byte through. Safe because `&` not followed by a
        // recognized entity is left intact, and all other bytes are valid UTF-8
        // boundaries inside a `&str`.
        let ch_len = source_char_len(&s[i..]);
        out.push_str(&s[i..i + ch_len]);
        i += ch_len;
    }
    out
}

fn source_char_len(s: &str) -> usize {
    s.chars().next().map(|c| c.len_utf8()).unwrap_or(1)
}

pub fn compose_from_file(spec_path: &Path, output: &Path) -> Result<()> {
    let text = fs::read_to_string(spec_path)
        .with_context(|| format!("read compose spec: {}", spec_path.display()))?;
    let spec: ComposeSpec = serde_json::from_str(&text).context("parse docx compose spec")?;
    compose_docx(&spec, output)
}

pub fn compose_docx(spec: &ComposeSpec, output: &Path) -> Result<()> {
    if let Some(template) = spec.template.clone() {
        return compose_into_template(spec, &template, output);
    }

    let mut body = String::new();
    let mut footnotes = Vec::new();
    for block in &spec.blocks {
        render_block(block, &spec.brand, &mut body, &mut footnotes, false);
    }
    if body.trim().is_empty()
        && let Some(title) = &spec.meta.title
    {
        render_block(
            &DocxBlock::Heading {
                text: title.clone(),
                level: 1,
                page_break: false,
            },
            &spec.brand,
            &mut body,
            &mut footnotes,
            false,
        );
    }

    let has_footer = spec
        .meta
        .footer_text
        .as_deref()
        .is_some_and(|text| !text.is_empty());
    let has_footnotes = !footnotes.is_empty();
    let section = if has_footer {
        r#"<w:sectPr><w:footerReference w:type="default" r:id="rIdFooter1"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>"#
    } else {
        r#"<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>"#
    };
    let document = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>{body}{section}</w:body></w:document>"#
    );

    let mut parts = base_docx_parts(spec, has_footer, has_footnotes);
    parts.insert("word/document.xml".to_string(), document.into_bytes());
    if has_footer {
        let footer_text = escaped(spec.meta.footer_text.as_deref().unwrap_or_default());
        parts.insert(
            "word/footer1.xml".to_string(),
            format!(
                r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r>{}<w:t>{footer_text}</w:t></w:r></w:p></w:ftr>"#,
                run_props(&spec.brand, "14", Some(&spec.brand.muted_color), false, false)
            )
            .into_bytes(),
        );
    }
    if has_footnotes {
        parts.insert(
            "word/footnotes.xml".to_string(),
            footnotes_xml(&footnotes, &spec.brand).into_bytes(),
        );
    }
    write_new_package(output, &parts)
}

fn compose_into_template(spec: &ComposeSpec, template: &Path, output: &Path) -> Result<()> {
    if !template.is_file() {
        bail!("template not found: {}", template.display());
    }
    let mut body = String::new();
    let mut footnotes = Vec::new();
    for block in &spec.blocks {
        render_block(block, &spec.brand, &mut body, &mut footnotes, true);
    }
    if !footnotes.is_empty() {
        bail!(
            "footnote blocks are not yet supported in template mode; the template's footnotes part would need to be merged"
        );
    }

    let template_doc_bytes = read_part(template, "word/document.xml")
        .with_context(|| format!("read word/document.xml from {}", template.display()))?;
    let template_doc = String::from_utf8(template_doc_bytes)
        .context("template word/document.xml is not valid UTF-8")?;

    let doc_open = extract_document_open_tag(&template_doc)
        .ok_or_else(|| anyhow!("template word/document.xml missing <w:document>"))?;
    let sect_pr = extract_last_sect_pr(&template_doc).unwrap_or_default();

    let new_doc = format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>{doc_open}<w:body>{body}{sect_pr}</w:body></w:document>"#
    );

    if let Some(parent) = output.parent()
        && !parent.as_os_str().is_empty()
    {
        fs::create_dir_all(parent)?;
    }
    fs::copy(template, output).with_context(|| {
        format!(
            "copy template {} to {}",
            template.display(),
            output.display()
        )
    })?;

    let mut updates = PartMap::new();
    updates.insert("word/document.xml".to_string(), new_doc.into_bytes());
    rewrite_parts(output, &updates, std::iter::empty::<&str>())?;
    Ok(())
}

fn extract_document_open_tag(xml: &str) -> Option<String> {
    let start = xml.find("<w:document")?;
    let after = &xml[start..];
    let mut depth = 0usize;
    for (i, c) in after.char_indices() {
        match c {
            '<' => depth += 1,
            '>' if depth == 1 => return Some(after[..=i].to_string()),
            '>' => depth = depth.saturating_sub(1),
            _ => {}
        }
    }
    None
}

fn extract_last_sect_pr(xml: &str) -> Option<String> {
    let body_start = xml.find("<w:body")?;
    let body_end = xml.find("</w:body>")?;
    if body_end <= body_start {
        return None;
    }
    let body = &xml[body_start..body_end];
    let sect_start = body.rfind("<w:sectPr")?;
    let after = &body[sect_start..];
    if let Some(close_rel) = after.find("</w:sectPr>") {
        let end = close_rel + "</w:sectPr>".len();
        Some(after[..end].to_string())
    } else {
        let tag_end = after.find('>')?;
        if after[..tag_end].ends_with('/') {
            Some(after[..=tag_end].to_string())
        } else {
            None
        }
    }
}

fn render_block(
    block: &DocxBlock,
    brand: &BrandSpec,
    body: &mut String,
    footnotes: &mut Vec<String>,
    style_mode: bool,
) {
    match block {
        DocxBlock::TitlePage {
            client_name,
            title,
            subtitle,
            date_text,
            prepared_by,
        } => {
            body.push_str(&spacer_p(900));
            if let Some(client) = client_name {
                body.push_str(&centered_p(
                    client,
                    brand,
                    "28",
                    Some(&brand.muted_color),
                    false,
                ));
            }
            body.push_str(&centered_p(
                &title.to_ascii_uppercase(),
                brand,
                "44",
                Some(&brand.heading_color),
                true,
            ));
            body.push_str(&divider_p(&brand.accent_color));
            if let Some(subtitle) = subtitle {
                body.push_str(&centered_p(
                    subtitle,
                    brand,
                    "28",
                    Some(&brand.body_color),
                    false,
                ));
            }
            if let Some(date_text) = date_text {
                body.push_str(&centered_p(
                    date_text,
                    brand,
                    "22",
                    Some(&brand.muted_color),
                    false,
                ));
            }
            if let Some(prepared_by) = prepared_by {
                body.push_str(&centered_p(
                    &format!("Prepared by {prepared_by}"),
                    brand,
                    "22",
                    Some(&brand.muted_color),
                    false,
                ));
            }
            body.push_str(&page_break_p());
        }
        DocxBlock::Heading {
            text,
            level,
            page_break,
        } => {
            if style_mode {
                let style = format!("Heading{}", level.clamp(&1, &9));
                let extra = if *level == 1 && !page_break {
                    r#"<w:pageBreakBefore w:val="0"/>"#
                } else if *page_break && *level != 1 {
                    "<w:pageBreakBefore/>"
                } else {
                    ""
                };
                body.push_str(&styled_paragraph(
                    &style,
                    extra,
                    &plain_run(text, false, false),
                ));
                return;
            }
            if *page_break {
                body.push_str(&page_break_p());
            }
            let size = match level {
                1 => "44",
                2 => "32",
                3 => "24",
                _ => "22",
            };
            let style = format!("Heading{}", level.clamp(&1, &4));
            body.push_str(&paragraph(
                Some(&style),
                text,
                brand,
                size,
                Some(&brand.heading_color),
                *level > 1,
                false,
            ));
        }
        DocxBlock::Body {
            text,
            bold,
            footnote,
        } => {
            if style_mode {
                if footnote.is_some() {
                    // Footnotes aren't supported in template mode (see compose_into_template).
                    // Surface the body text without the footnote reference rather than dropping the paragraph.
                }
                body.push_str(&styled_paragraph(
                    "BodyText",
                    "",
                    &plain_run(text, *bold, false),
                ));
                return;
            }
            let mut paragraph_xml = paragraph(
                None,
                text,
                brand,
                "18",
                Some(&brand.body_color),
                *bold,
                false,
            );
            if let Some(note) = footnote {
                insert_before_paragraph_end(&mut paragraph_xml, &footnote_ref(footnotes, note));
            }
            body.push_str(&paragraph_xml);
        }
        DocxBlock::BodyRich { segments } => {
            if style_mode {
                let mut runs = String::new();
                for segment in segments {
                    runs.push_str(&plain_rich_run(
                        &segment.text,
                        segment.bold,
                        segment.italic,
                        segment.color.as_deref(),
                    ));
                }
                body.push_str(&styled_paragraph("BodyText", "", &runs));
                return;
            }
            body.push_str("<w:p>");
            for segment in segments {
                body.push_str(&run(
                    &segment.text,
                    brand,
                    "18",
                    Some(segment.color.as_deref().unwrap_or(&brand.body_color)),
                    segment.bold,
                    segment.italic,
                ));
            }
            body.push_str("</w:p>");
        }
        DocxBlock::Bullet {
            text,
            level,
            bold_prefix,
            footnote,
            style,
        } => {
            if style_mode {
                let resolved_style = style
                    .clone()
                    .unwrap_or_else(|| format!("Bullet{}", level.clamp(&1, &4)));
                let mut runs = String::new();
                if let Some(prefix) = bold_prefix.as_deref() {
                    runs.push_str(&plain_run(prefix, true, false));
                }
                runs.push_str(&plain_run(text, false, false));
                body.push_str(&styled_paragraph(&resolved_style, "", &runs));
                return;
            }
            let indent = 360usize * (*level as usize).max(1);
            let note = footnote
                .as_ref()
                .map(|text| footnote_ref(footnotes, text))
                .unwrap_or_default();
            body.push_str(&format!(
                r#"<w:p><w:pPr><w:ind w:left="{indent}" w:hanging="240"/></w:pPr>{}{}{}</w:p>"#,
                run("-  ", brand, "18", Some(&brand.body_color), false, false),
                bold_prefix
                    .as_deref()
                    .map(|prefix| run(prefix, brand, "18", Some(&brand.body_color), true, false))
                    .unwrap_or_default(),
                run(text, brand, "18", Some(&brand.body_color), false, false) + &note
            ));
        }
        DocxBlock::Numbered {
            text,
            number,
            bold_prefix,
            footnote,
        } => {
            if style_mode {
                let mut runs = String::new();
                if let Some(prefix) = bold_prefix.as_deref() {
                    runs.push_str(&plain_run(prefix, true, false));
                }
                runs.push_str(&plain_run(text, false, false));
                body.push_str(&styled_paragraph("ListNumber", "", &runs));
                return;
            }
            let note = footnote
                .as_ref()
                .map(|text| footnote_ref(footnotes, text))
                .unwrap_or_default();
            body.push_str(&format!(
                r#"<w:p><w:pPr><w:ind w:left="360" w:hanging="240"/></w:pPr>{}{}{}</w:p>"#,
                run(
                    &format!("{number}.  "),
                    brand,
                    "18",
                    Some(&brand.body_color),
                    false,
                    false
                ),
                bold_prefix
                    .as_deref()
                    .map(|prefix| run(prefix, brand, "18", Some(&brand.body_color), true, false))
                    .unwrap_or_default(),
                run(text, brand, "18", Some(&brand.body_color), false, false) + &note
            ));
        }
        DocxBlock::Quote { text } => {
            if style_mode {
                body.push_str(&styled_paragraph(
                    "Quote",
                    "",
                    &plain_run(text, false, true),
                ));
                return;
            }
            body.push_str(&format!(
                r#"<w:p><w:pPr><w:ind w:left="567"/></w:pPr>{}</w:p>"#,
                run(text, brand, "18", Some(&brand.body_color), false, true)
            ));
        }
        DocxBlock::QuoteBlock { lines } => {
            for line in lines {
                body.push_str(&format!(
                    r#"<w:p><w:pPr><w:ind w:left="567"/><w:pBdr><w:left w:val="single" w:sz="12" w:space="8" w:color="{}"/></w:pBdr><w:shd w:fill="F2F5F9" w:val="clear"/></w:pPr>{}</w:p>"#,
                    brand.accent_color,
                    run(line, brand, "18", Some(&brand.heading_color), false, false)
                ));
            }
        }
        DocxBlock::Table {
            headers,
            rows,
            alt_row_shading,
            first_col_bold,
            col_widths_twips,
        } => body.push_str(&table(
            headers,
            rows,
            brand,
            *alt_row_shading,
            *first_col_bold,
            true,
            style_mode,
            col_widths_twips.as_deref(),
        )),
        DocxBlock::BorderlessTable {
            rows,
            label_col,
            col_widths_twips,
            borders,
        } => {
            body.push_str(&borderless_table(
                rows,
                brand,
                *label_col,
                style_mode,
                col_widths_twips.as_deref(),
                borders.as_deref().unwrap_or("none"),
            ));
        }
        DocxBlock::Divider => body.push_str(&divider_p(&brand.accent_color)),
        DocxBlock::Spacer { height_twips } => body.push_str(&spacer_p(*height_twips)),
        DocxBlock::PageBreak => body.push_str(&page_break_p()),
    }
}

fn paragraph(
    style: Option<&str>,
    text: &str,
    brand: &BrandSpec,
    size: &str,
    color: Option<&str>,
    bold: bool,
    italic: bool,
) -> String {
    let style_xml = style
        .map(|style| format!(r#"<w:pPr><w:pStyle w:val="{style}"/></w:pPr>"#))
        .unwrap_or_default();
    format!(
        "<w:p>{style_xml}{}</w:p>",
        run(text, brand, size, color, bold, italic)
    )
}

/// Build a paragraph that references a named style and contains pre-rendered runs.
/// `extra_ppr` is inserted inside `<w:pPr>` after the `<w:pStyle>` reference;
/// callers use it to override pageBreakBefore or similar style-defined behaviour.
fn styled_paragraph(style: &str, extra_ppr: &str, content_runs: &str) -> String {
    format!(r#"<w:p><w:pPr><w:pStyle w:val="{style}"/>{extra_ppr}</w:pPr>{content_runs}</w:p>"#)
}

/// A `<w:r>` with only the bare run properties — no font/color/size, since the
/// surrounding paragraph style is expected to carry typography.
fn plain_run(text: &str, bold: bool, italic: bool) -> String {
    let mut rpr = String::new();
    if bold || italic {
        rpr.push_str("<w:rPr>");
        if bold {
            rpr.push_str("<w:b/>");
        }
        if italic {
            rpr.push_str("<w:i/>");
        }
        rpr.push_str("</w:rPr>");
    }
    format!(
        r#"<w:r>{rpr}<w:t xml:space="preserve">{}</w:t></w:r>"#,
        escaped(text)
    )
}

/// Plain run with an optional explicit color override (still no font/size).
fn plain_rich_run(text: &str, bold: bool, italic: bool, color: Option<&str>) -> String {
    let mut rpr = String::new();
    if bold || italic || color.is_some() {
        rpr.push_str("<w:rPr>");
        if bold {
            rpr.push_str("<w:b/>");
        }
        if italic {
            rpr.push_str("<w:i/>");
        }
        if let Some(c) = color {
            rpr.push_str(&format!(r#"<w:color w:val="{}"/>"#, escaped(c)));
        }
        rpr.push_str("</w:rPr>");
    }
    format!(
        r#"<w:r>{rpr}<w:t xml:space="preserve">{}</w:t></w:r>"#,
        escaped(text)
    )
}

fn centered_p(
    text: &str,
    brand: &BrandSpec,
    size: &str,
    color: Option<&str>,
    bold: bool,
) -> String {
    format!(
        r#"<w:p><w:pPr><w:jc w:val="center"/></w:pPr>{}</w:p>"#,
        run(text, brand, size, color, bold, false)
    )
}

fn divider_p(color: &str) -> String {
    format!(
        r#"<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="8" w:space="1" w:color="{}"/></w:pBdr></w:pPr></w:p>"#,
        escaped(color)
    )
}

fn spacer_p(height_twips: usize) -> String {
    format!(r#"<w:p><w:pPr><w:spacing w:after="{height_twips}"/></w:pPr></w:p>"#)
}

fn page_break_p() -> String {
    r#"<w:p><w:r><w:br w:type="page"/></w:r></w:p>"#.to_string()
}

fn run(
    text: &str,
    brand: &BrandSpec,
    size: &str,
    color: Option<&str>,
    bold: bool,
    italic: bool,
) -> String {
    format!(
        r#"<w:r>{}<w:t xml:space="preserve">{}</w:t></w:r>"#,
        run_props(brand, size, color, bold, italic),
        escaped(text)
    )
}

fn run_props(
    brand: &BrandSpec,
    size: &str,
    color: Option<&str>,
    bold: bool,
    italic: bool,
) -> String {
    let color = color.unwrap_or(&brand.body_color);
    format!(
        r#"<w:rPr><w:rFonts w:ascii="{}" w:hAnsi="{}" w:cs="{}"/>{}{}<w:color w:val="{}"/><w:sz w:val="{}"/><w:szCs w:val="{}"/></w:rPr>"#,
        escaped(&brand.font_family),
        escaped(&brand.font_family),
        escaped(&brand.font_family),
        if bold { "<w:b/>" } else { "" },
        if italic { "<w:i/>" } else { "" },
        escaped(color),
        size,
        size
    )
}

fn footnote_ref(footnotes: &mut Vec<String>, text: &str) -> String {
    footnotes.push(text.to_string());
    let id = footnotes.len();
    format!(
        r#"<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="{id}"/></w:r>"#
    )
}

fn insert_before_paragraph_end(paragraph_xml: &mut String, insert: &str) {
    if let Some(pos) = paragraph_xml.rfind("</w:p>") {
        paragraph_xml.insert_str(pos, insert);
    } else {
        paragraph_xml.push_str(insert);
    }
}

fn footnotes_xml(footnotes: &[String], brand: &BrandSpec) -> String {
    let mut entries = String::new();
    entries.push_str(r#"<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>"#);
    entries.push_str(r#"<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>"#);
    for (idx, text) in footnotes.iter().enumerate() {
        let id = idx + 1;
        entries.push_str(&format!(
            r#"<w:footnote w:id="{id}"><w:p><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>{}</w:p></w:footnote>"#,
            run(
                &format!(" {text}"),
                brand,
                "16",
                Some(&brand.body_color),
                false,
                false
            )
        ));
    }
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{entries}</w:footnotes>"#
    )
}

#[allow(clippy::too_many_arguments)]
fn table(
    headers: &[String],
    rows: &[Vec<CellContent>],
    brand: &BrandSpec,
    alt_rows: bool,
    first_col_bold: bool,
    borders: bool,
    style_mode: bool,
    col_widths: Option<&[usize]>,
) -> String {
    let mut out = String::new();
    out.push_str("<w:tbl>");
    out.push_str(r#"<w:tblPr><w:tblW w:w="0" w:type="auto"/>"#);
    if borders {
        out.push_str(r#"<w:tblBorders><w:top w:val="single" w:sz="4" w:color="D9DEE8"/><w:left w:val="single" w:sz="4" w:color="D9DEE8"/><w:bottom w:val="single" w:sz="4" w:color="D9DEE8"/><w:right w:val="single" w:sz="4" w:color="D9DEE8"/><w:insideH w:val="single" w:sz="4" w:color="D9DEE8"/><w:insideV w:val="single" w:sz="4" w:color="D9DEE8"/></w:tblBorders>"#);
    } else {
        out.push_str(r#"<w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders>"#);
    }
    out.push_str("</w:tblPr>");
    out.push_str(&tbl_grid(col_widths, headers.len()));
    out.push_str("<w:tr>");
    for (col_idx, header) in headers.iter().enumerate() {
        let cell_value = CellContent::Text(header.clone());
        out.push_str(&cell(
            &cell_value,
            brand,
            Some(&brand.table_header_fill),
            Some("FFFFFF"),
            true,
            style_mode,
            col_width_at(col_widths, col_idx),
        ));
    }
    out.push_str("</w:tr>");
    for (row_idx, row) in rows.iter().enumerate() {
        out.push_str("<w:tr>");
        for (col_idx, value) in row.iter().enumerate() {
            let fill = if alt_rows && row_idx % 2 == 1 {
                Some(brand.alt_row_fill.as_str())
            } else {
                None
            };
            out.push_str(&cell(
                value,
                brand,
                fill,
                Some(&brand.body_color),
                first_col_bold && col_idx == 0,
                style_mode,
                col_width_at(col_widths, col_idx),
            ));
        }
        out.push_str("</w:tr>");
    }
    out.push_str("</w:tbl>");
    out
}

fn borderless_table(
    rows: &[Vec<CellContent>],
    brand: &BrandSpec,
    label_col: usize,
    style_mode: bool,
    col_widths: Option<&[usize]>,
    border_style: &str,
) -> String {
    let mut out = String::new();
    let borders_xml = match border_style {
        "rows" => {
            r#"<w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/><w:insideH w:val="single" w:sz="4" w:color="D9DEE8"/><w:insideV w:val="nil"/></w:tblBorders>"#
        }
        "all" => {
            r#"<w:tblBorders><w:top w:val="single" w:sz="4" w:color="D9DEE8"/><w:left w:val="single" w:sz="4" w:color="D9DEE8"/><w:bottom w:val="single" w:sz="4" w:color="D9DEE8"/><w:right w:val="single" w:sz="4" w:color="D9DEE8"/><w:insideH w:val="single" w:sz="4" w:color="D9DEE8"/><w:insideV w:val="single" w:sz="4" w:color="D9DEE8"/></w:tblBorders>"#
        }
        _ => {
            r#"<w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders>"#
        }
    };
    out.push_str(&format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>{borders_xml}</w:tblPr>"#
    ));
    let row_width = rows.iter().map(|r| r.len()).max().unwrap_or(0);
    out.push_str(&tbl_grid(col_widths, row_width));
    for row in rows {
        out.push_str("<w:tr>");
        for (idx, value) in row.iter().enumerate() {
            out.push_str(&cell(
                value,
                brand,
                None,
                Some(&brand.body_color),
                idx == label_col,
                style_mode,
                col_width_at(col_widths, idx),
            ));
        }
        out.push_str("</w:tr>");
    }
    out.push_str("</w:tbl>");
    out
}

fn tbl_grid(col_widths: Option<&[usize]>, expected_cols: usize) -> String {
    let widths = col_widths.unwrap_or(&[]);
    if widths.is_empty() {
        return String::new();
    }
    let mut out = String::from("<w:tblGrid>");
    for w in widths {
        out.push_str(&format!(r#"<w:gridCol w:w="{w}"/>"#));
    }
    // Pad with auto-width columns if the spec under-specified.
    for _ in widths.len()..expected_cols {
        out.push_str(r#"<w:gridCol w:w="0"/>"#);
    }
    out.push_str("</w:tblGrid>");
    out
}

fn col_width_at(col_widths: Option<&[usize]>, idx: usize) -> Option<usize> {
    col_widths.and_then(|w| w.get(idx).copied())
}

#[allow(clippy::too_many_arguments)]
fn cell(
    content: &CellContent,
    brand: &BrandSpec,
    fill: Option<&str>,
    default_color: Option<&str>,
    default_bold: bool,
    style_mode: bool,
    width_twips: Option<usize>,
) -> String {
    let shading = fill
        .map(|fill| format!(r#"<w:shd w:fill="{}" w:val="clear"/>"#, escaped(fill)))
        .unwrap_or_default();
    let width = width_twips
        .map(|w| format!(r#"<w:tcW w:w="{w}" w:type="dxa"/>"#))
        .unwrap_or_default();

    // Resolve per-cell overrides on top of the row-level defaults.
    let rich = content.rich();
    let bold = default_bold || rich.map(|r| r.bold).unwrap_or(false);
    let italic = rich.map(|r| r.italic).unwrap_or(false);
    let color = rich.and_then(|r| r.color.as_deref()).or(default_color);
    let text = content.as_str();

    let inner = if style_mode {
        // Header cell (with fill) — use the template's "Bodytextwhite" style so the cell
        // text is white-on-fill without inline font/color overrides. Body cells fall back
        // to "TableBody". Bold/italic and per-cell color override come through as inline rPr.
        let style = if fill.is_some() {
            "Bodytextwhite"
        } else {
            "TableBody"
        };
        let run_xml = if rich.is_some() {
            plain_rich_run(text, bold, italic, color)
        } else {
            plain_run(text, bold, italic)
        };
        styled_paragraph(style, "", &run_xml)
    } else {
        format!("<w:p>{}</w:p>", run(text, brand, "18", color, bold, italic))
    };
    format!(
        r#"<w:tc><w:tcPr>{width}{shading}<w:tcMar><w:top w:w="40" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/><w:left w:w="80" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tcMar></w:tcPr>{inner}</w:tc>"#
    )
}

fn base_docx_parts(spec: &ComposeSpec, has_footer: bool, has_footnotes: bool) -> PartMap {
    let mut parts = PartMap::new();
    parts.insert(
        "[Content_Types].xml".to_string(),
        content_types_xml(has_footer, has_footnotes).into_bytes(),
    );
    parts.insert("_rels/.rels".to_string(), root_rels_xml().into_bytes());
    parts.insert(
        "word/_rels/document.xml.rels".to_string(),
        document_rels_xml(has_footer, has_footnotes).into_bytes(),
    );
    parts.insert(
        "word/styles.xml".to_string(),
        styles_xml(&spec.brand).into_bytes(),
    );
    parts.insert(
        "docProps/core.xml".to_string(),
        core_xml(&spec.meta).into_bytes(),
    );
    parts.insert("docProps/app.xml".to_string(), app_xml().into_bytes());
    parts
}

fn content_types_xml(has_footer: bool, has_footnotes: bool) -> String {
    let footer_override = if has_footer {
        r#"<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>"#
    } else {
        ""
    };
    let footnotes_override = if has_footnotes {
        r#"<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>"#
    } else {
        ""
    };
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
{}{footnotes_override}<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>"#,
        footer_override
    )
}

fn root_rels_xml() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>"#
        .to_string()
}

fn document_rels_xml(has_footer: bool, has_footnotes: bool) -> String {
    let footer_rel = if has_footer {
        r#"<Relationship Id="rIdFooter1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>"#
    } else {
        ""
    };
    let footnotes_rel = if has_footnotes {
        r#"<Relationship Id="rIdFootnotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>"#
    } else {
        ""
    };
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
{}{footnotes_rel}</Relationships>"#,
        footer_rel
    )
}

fn styles_xml(brand: &BrandSpec) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr>{}</w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:color w:val="{}"/><w:sz w:val="44"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="Heading 2"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:color w:val="{}"/><w:sz w:val="32"/></w:rPr></w:style>
</w:styles>"#,
        run_props(brand, "18", Some(&brand.body_color), false, false),
        escaped(&brand.heading_color),
        escaped(&brand.heading_color)
    )
}

fn core_xml(meta: &ComposeMeta) -> String {
    let title = escaped(meta.title.as_deref().unwrap_or("Document"));
    let subject = escaped(meta.subject.as_deref().unwrap_or(""));
    let creator = escaped(meta.creator.as_deref().unwrap_or("office-tools"));
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>{title}</dc:title><dc:subject>{subject}</dc:subject><dc:creator>{creator}</dc:creator></cp:coreProperties>"#
    )
}

fn app_xml() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>office-tools</Application></Properties>"#.to_string()
}

fn paragraph_text(paragraph: &Element) -> String {
    let mut out = String::new();
    collect_text_nodes(paragraph, &mut out);
    out
}

fn paragraph_style(paragraph: &Element) -> Option<String> {
    let p_pr = child(paragraph, "pPr")?;
    let style = child(p_pr, "pStyle")?;
    style
        .attributes
        .get("w:val")
        .or_else(|| style.attributes.get("val"))
        .cloned()
}

fn table_rows(table: &Element) -> Vec<Vec<String>> {
    children(table, "tr")
        .map(|row| {
            children(row, "tc")
                .map(|cell| {
                    let mut text = String::new();
                    collect_text_nodes(cell, &mut text);
                    text.trim().replace('\n', " ")
                })
                .collect::<Vec<_>>()
        })
        .collect()
}

fn normalized_row(row: &[String], max_cols: usize) -> Vec<String> {
    let mut row = row.to_vec();
    row.resize(max_cols, String::new());
    row
}

fn collect_text_nodes(element: &Element, out: &mut String) {
    for node in &element.children {
        match node {
            XMLNode::Element(el) if el.name == "t" => {
                for child in &el.children {
                    if let XMLNode::Text(text) | XMLNode::CData(text) = child {
                        out.push_str(text);
                    }
                }
            }
            XMLNode::Element(el) => {
                if el.name == "tab" {
                    out.push('\t');
                } else if el.name == "br" {
                    out.push('\n');
                }
                collect_text_nodes(el, out);
            }
            _ => {}
        }
    }
}

fn default_font() -> String {
    "Arial".to_string()
}

fn default_body_color() -> String {
    "333333".to_string()
}

fn default_accent_color() -> String {
    "2F6F9F".to_string()
}

fn default_heading_color() -> String {
    "1F2A44".to_string()
}

fn default_muted_color() -> String {
    "666666".to_string()
}

fn default_table_header_fill() -> String {
    "1F2A44".to_string()
}

fn default_alt_row_fill() -> String {
    "F2F5F9".to_string()
}

fn default_heading_level() -> u8 {
    1
}

fn default_list_level() -> u8 {
    1
}

fn default_number() -> usize {
    1
}

fn default_spacer_twips() -> usize {
    240
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn compose_and_read_generic_docx() -> Result<()> {
        let dir = tempfile::tempdir()?;
        let path = dir.path().join("doc.docx");
        let spec = ComposeSpec {
            meta: ComposeMeta {
                title: Some("Test memo".to_string()),
                subject: None,
                creator: None,
                footer_text: Some("Generic footer".to_string()),
            },
            brand: BrandSpec::default(),
            blocks: vec![
                DocxBlock::Heading {
                    text: "Memo".to_string(),
                    level: 1,
                    page_break: false,
                },
                DocxBlock::Body {
                    text: "Body text".to_string(),
                    bold: false,
                    footnote: Some("Footnote text".to_string()),
                },
                DocxBlock::Table {
                    headers: vec!["A".to_string(), "B".to_string()],
                    rows: vec![vec![
                        CellContent::Text("1".to_string()),
                        CellContent::Text("2".to_string()),
                    ]],
                    alt_row_shading: true,
                    first_col_bold: false,
                    col_widths_twips: None,
                },
            ],
            template: None,
        };
        compose_docx(&spec, &path)?;
        let markdown = read_docx_markdown(&path)?;
        let footnotes = crate::ooxml::read_part_to_string(&path, "word/footnotes.xml")?;
        assert!(markdown.contains("Memo"));
        assert!(markdown.contains("Body text"));
        assert!(markdown.contains("| A | B |"));
        assert!(footnotes.contains("Footnote text"));
        Ok(())
    }

    #[test]
    fn compose_into_template_preserves_template_parts_and_replaces_body() -> Result<()> {
        let dir = tempfile::tempdir()?;
        let template = dir.path().join("template.docx");
        let output = dir.path().join("output.docx");

        // Create a "template" by composing a docx with a sentinel block + footer.
        // The template carries word/footer1.xml; compose_into_template must
        // preserve that part while replacing only the body.
        let template_spec = ComposeSpec {
            meta: ComposeMeta {
                title: None,
                subject: None,
                creator: None,
                footer_text: Some("TEMPLATE FOOTER".to_string()),
            },
            brand: BrandSpec::default(),
            blocks: vec![DocxBlock::Body {
                text: "PLACEHOLDER_BODY_TEXT".to_string(),
                bold: false,
                footnote: None,
            }],
            template: None,
        };
        compose_docx(&template_spec, &template)?;

        // Compose into the template - body should be replaced, footer kept.
        let into_spec = ComposeSpec {
            meta: ComposeMeta::default(),
            brand: BrandSpec::default(),
            blocks: vec![
                DocxBlock::Heading {
                    text: "Composed Heading".to_string(),
                    level: 1,
                    page_break: false,
                },
                DocxBlock::Body {
                    text: "Composed body content".to_string(),
                    bold: false,
                    footnote: None,
                },
            ],
            template: Some(template.clone()),
        };
        compose_docx(&into_spec, &output)?;

        let markdown = read_docx_markdown(&output)?;
        assert!(
            markdown.contains("Composed Heading"),
            "new body missing: {markdown}"
        );
        assert!(
            markdown.contains("Composed body content"),
            "new body missing: {markdown}"
        );
        assert!(
            !markdown.contains("PLACEHOLDER_BODY_TEXT"),
            "template body should be replaced, found placeholder in: {markdown}"
        );

        // The footer part from the template must still be present in the output package.
        let footer = crate::ooxml::read_part_to_string(&output, "word/footer1.xml")?;
        assert!(
            footer.contains("TEMPLATE FOOTER"),
            "template footer should be preserved: {footer}"
        );
        Ok(())
    }

    /// Regression: a `direct_replace` round-trip must preserve every byte of
    /// the document.xml that doesn't intersect a replaced text node. The
    /// previous implementation parsed the part via `xmltree` and re-serialized
    /// it, which lost or rearranged namespace declarations, `xml:space`
    /// attributes, and self-closing-tag style — producing files that Word
    /// would only open via the "Recover" prompt and then render with broken
    /// formatting. The new byte-level implementation rewrites only the inner
    /// text of `<w:t>...</w:t>` elements.
    #[test]
    fn direct_replace_preserves_complex_ooxml_byte_for_byte() -> Result<()> {
        use crate::ooxml::{list_parts, read_part_to_string, write_new_package};

        let dir = tempfile::tempdir()?;
        let path = dir.path().join("complex.docx");

        // Hand-rolled document.xml that exercises the quirks the xmltree
        // round-trip used to mangle: multiple namespace declarations on the
        // root, mc:AlternateContent (Word 2010+ fallback wrapper), xml:space
        // = "preserve" on whitespace-bearing text nodes, and a self-closing
        // `<w:tab/>`.
        let document = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" mc:Ignorable="w14 w15 wp14"><w:body><w:p><w:r><w:t xml:space="preserve">Hello </w:t></w:r><w:r><w:t>SENTINEL</w:t></w:r><w:r><w:t xml:space="preserve"> world.</w:t></w:r></w:p><w:p><w:r><w:tab/><w:t>Tabbed</w:t></w:r></w:p><mc:AlternateContent><mc:Choice Requires="wp14"><w:p><w:r><w:t>SENTINEL alt</w:t></w:r></w:p></mc:Choice><mc:Fallback><w:p><w:r><w:t>plain alt</w:t></w:r></w:p></mc:Fallback></mc:AlternateContent></w:body></w:document>"#.to_vec();

        let mut parts = PartMap::new();
        parts.insert(
            "[Content_Types].xml".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec(),
        );
        parts.insert(
            "_rels/.rels".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec(),
        );
        parts.insert("word/document.xml".to_string(), document.clone());
        write_new_package(&path, &parts)?;

        let before = read_part_to_string(&path, "word/document.xml")?;
        // Run the replacement via the same entry point the MCP tool uses.
        let result = replace(&DocxReplaceArgs {
            file: path.clone(),
            output: None,
            find: Some("SENTINEL".to_string()),
            replace: Some("REPLACED".to_string()),
            replacements: None,
            replacements_file: None,
            engine: DocxReplaceEngine::Direct,
        })?;
        assert_eq!(result.counts.len(), 1);
        assert_eq!(result.counts[0].count, 2);

        let after = read_part_to_string(&path, "word/document.xml")?;
        // The only differences must be the two SENTINEL -> REPLACED swaps
        // inside <w:t> elements. Confirm by reverse-substituting and checking
        // byte equality with the original.
        let restored = after.replace("REPLACED", "SENTINEL");
        assert_eq!(
            restored, before,
            "direct_replace must not change any bytes outside the targeted <w:t> text"
        );

        // Sanity-check the structural quirks still survive verbatim.
        assert!(after.contains(r#"xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006""#));
        assert!(after.contains(r#"xml:space="preserve""#));
        assert!(after.contains("<w:tab/>"));
        assert!(after.contains("<mc:AlternateContent>"));
        assert!(after.contains(r#"mc:Ignorable="w14 w15 wp14""#));

        // The package still contains exactly the parts we wrote.
        let parts_now: Vec<_> = list_parts(&path)?
            .into_iter()
            .map(|p| p.name)
            .collect();
        assert!(parts_now.contains(&"word/document.xml".to_string()));
        assert!(parts_now.contains(&"[Content_Types].xml".to_string()));
        Ok(())
    }

    #[test]
    fn direct_replace_xml_escapes_replacement_text() -> Result<()> {
        use crate::ooxml::{read_part_to_string, write_new_package};

        let dir = tempfile::tempdir()?;
        let path = dir.path().join("escape.docx");

        let document = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>OLD</w:t></w:r></w:p></w:body></w:document>"#.to_vec();

        let mut parts = PartMap::new();
        parts.insert(
            "[Content_Types].xml".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec(),
        );
        parts.insert(
            "_rels/.rels".to_string(),
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec(),
        );
        parts.insert("word/document.xml".to_string(), document);
        write_new_package(&path, &parts)?;

        // Replacement that contains XML metacharacters must be escaped on
        // write so the resulting document.xml stays well-formed.
        replace(&DocxReplaceArgs {
            file: path.clone(),
            output: None,
            find: Some("OLD".to_string()),
            replace: Some("A & B < C > D".to_string()),
            replacements: None,
            replacements_file: None,
            engine: DocxReplaceEngine::Direct,
        })?;

        let xml = read_part_to_string(&path, "word/document.xml")?;
        assert!(xml.contains("A &amp; B &lt; C &gt; D"));
        // Round-trip through the parser must still succeed.
        let _ = parse_xml(xml.as_bytes())?;
        Ok(())
    }
}
