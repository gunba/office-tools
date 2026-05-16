use crate::ooxml::{
    PartMap, child, children, escaped, parse_xml, read_part, resolve_relationship_target,
    write_new_package,
};
use anyhow::{Context, Result, anyhow};
use clap::{Args, Subcommand};
use serde::Deserialize;
use std::collections::BTreeMap;
use std::fs;
use std::path::{Path, PathBuf};
use xmltree::{Element, XMLNode};

#[derive(Debug, Args)]
pub struct PptxArgs {
    #[command(subcommand)]
    command: PptxCommand,
}

#[derive(Debug, Subcommand)]
pub enum PptxCommand {
    /// Dump slide text as markdown.
    Read(PptxReadArgs),
    /// Extract speaker notes.
    Notes(PptxReadArgs),
    /// Build a simple 16:9 deck from JSON.
    Build(PptxBuildArgs),
}

#[derive(Debug, Args)]
pub struct PptxReadArgs {
    pub file: PathBuf,
    #[arg(long)]
    pub json: bool,
}

#[derive(Debug, Args)]
pub struct PptxBuildArgs {
    pub spec: PathBuf,
    #[arg(short, long)]
    pub output: PathBuf,
}

#[derive(Debug, Deserialize)]
pub struct DeckSpec {
    #[serde(default)]
    pub meta: DeckMeta,
    #[serde(default)]
    pub brand: DeckBrand,
    #[serde(default)]
    pub slides: Vec<SlideSpec>,
}

#[derive(Debug, Default, Deserialize)]
pub struct DeckMeta {
    pub title: Option<String>,
    pub subtitle: Option<String>,
    pub author: Option<String>,
    pub date: Option<String>,
    pub footer_text: Option<String>,
}

#[derive(Debug, Clone, Deserialize)]
pub struct DeckBrand {
    #[serde(default = "default_font")]
    pub font_family: String,
    #[serde(default = "default_bg")]
    pub background: String,
    #[serde(default = "default_text")]
    pub text: String,
    #[serde(default = "default_accent")]
    pub accent: String,
    #[serde(default = "default_muted")]
    pub muted: String,
}

impl Default for DeckBrand {
    fn default() -> Self {
        Self {
            font_family: default_font(),
            background: default_bg(),
            text: default_text(),
            accent: default_accent(),
            muted: default_muted(),
        }
    }
}

#[derive(Debug, Clone, Deserialize)]
pub struct SlideSpec {
    #[serde(default = "default_layout")]
    pub layout: String,
    pub title: Option<String>,
    pub subtitle: Option<String>,
    pub number: Option<String>,
    pub body: Option<String>,
    pub bullets: Option<Vec<Bullet>>,
    pub left_title: Option<String>,
    pub left_bullets: Option<Vec<Bullet>>,
    pub right_title: Option<String>,
    pub right_bullets: Option<Vec<Bullet>>,
    pub panels: Option<Vec<Panel>>,
    pub quote: Option<String>,
    pub attribution: Option<String>,
    pub headers: Option<Vec<String>>,
    pub rows: Option<Vec<Vec<String>>>,
    pub placeholder_text: Option<String>,
    pub image_prompt: Option<String>,
    pub notes: Option<String>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(untagged)]
pub enum Bullet {
    Text(String),
    Rich {
        text: String,
        sub: Option<Vec<String>>,
    },
}

#[derive(Debug, Clone, Deserialize)]
pub struct Panel {
    pub title: Option<String>,
    #[serde(default)]
    pub bullets: Vec<Bullet>,
}

impl PptxArgs {
    pub fn run(self) -> Result<()> {
        match self.command {
            PptxCommand::Read(args) => {
                let slides = read_slides(&args.file)?;
                if args.json {
                    println!("{}", serde_json::to_string_pretty(&slides)?);
                } else {
                    for slide in slides {
                        println!("## Slide {}", slide.index);
                        if !slide.text.trim().is_empty() {
                            println!("{}", slide.text.trim());
                        }
                        println!();
                    }
                }
                Ok(())
            }
            PptxCommand::Notes(args) => {
                let notes = read_notes(&args.file)?;
                if args.json {
                    println!("{}", serde_json::to_string_pretty(&notes)?);
                } else {
                    for note in notes {
                        if !note.text.trim().is_empty() {
                            println!("## Slide {}", note.index);
                            println!("{}", note.text.trim());
                            println!();
                        }
                    }
                }
                Ok(())
            }
            PptxCommand::Build(args) => {
                let text = fs::read_to_string(&args.spec)
                    .with_context(|| format!("read deck spec: {}", args.spec.display()))?;
                let spec: DeckSpec = serde_json::from_str(&text).context("parse pptx spec JSON")?;
                build_deck(&spec, &args.output)?;
                println!("Saved: {}", args.output.display());
                Ok(())
            }
        }
    }
}

#[derive(Debug, serde::Serialize)]
pub struct SlideText {
    pub index: usize,
    pub part: String,
    pub text: String,
}

pub fn read_slides(path: &Path) -> Result<Vec<SlideText>> {
    let slide_parts = ordered_slide_parts(path)?;
    let mut slides = Vec::new();
    for (idx, part) in slide_parts.iter().enumerate() {
        let xml = parse_xml(&read_part(path, part)?)?;
        let mut lines = Vec::new();
        collect_text_runs(&xml, &mut lines);
        slides.push(SlideText {
            index: idx + 1,
            part: part.clone(),
            text: lines.join("\n"),
        });
    }
    Ok(slides)
}

pub fn read_notes(path: &Path) -> Result<Vec<SlideText>> {
    let slide_parts = ordered_slide_parts(path)?;
    let mut out = Vec::new();
    for (idx, slide_part) in slide_parts.iter().enumerate() {
        let rels_part = slide_rels_part(slide_part);
        let Ok(rels_xml) = read_part(path, &rels_part) else {
            continue;
        };
        let rels = parse_xml(&rels_xml)?;
        for rel in children(&rels, "Relationship") {
            let rel_type = rel
                .attributes
                .get("Type")
                .map(String::as_str)
                .unwrap_or_default();
            if !rel_type.ends_with("/notesSlide") {
                continue;
            }
            let Some(target) = rel.attributes.get("Target") else {
                continue;
            };
            let notes_part = resolve_relationship_target(slide_part, target);
            let xml = parse_xml(&read_part(path, &notes_part)?)?;
            let mut lines = Vec::new();
            collect_text_runs(&xml, &mut lines);
            out.push(SlideText {
                index: idx + 1,
                part: notes_part,
                text: lines.join("\n"),
            });
        }
    }
    Ok(out)
}

pub fn build_deck(spec: &DeckSpec, output: &Path) -> Result<()> {
    let slides = if spec.slides.is_empty() {
        vec![SlideSpec {
            layout: "title".to_string(),
            title: spec
                .meta
                .title
                .clone()
                .or_else(|| Some("Presentation".to_string())),
            subtitle: spec.meta.subtitle.clone(),
            number: None,
            body: None,
            bullets: None,
            left_title: None,
            left_bullets: None,
            right_title: None,
            right_bullets: None,
            panels: None,
            quote: None,
            attribution: None,
            headers: None,
            rows: None,
            placeholder_text: None,
            image_prompt: None,
            notes: None,
        }]
    } else {
        spec.slides.clone()
    };

    let notes_texts = slides.iter().map(slide_notes_text).collect::<Vec<_>>();
    let mut parts = base_pptx_parts(slides.len(), &notes_texts);
    for (idx, slide) in slides.iter().enumerate() {
        let slide_no = idx + 1;
        parts.insert(
            format!("ppt/slides/slide{slide_no}.xml"),
            render_slide(
                slide,
                &spec.brand,
                slide_no,
                slides.len(),
                spec.meta.footer_text.as_deref(),
            )
            .into_bytes(),
        );
        parts.insert(
            format!("ppt/slides/_rels/slide{slide_no}.xml.rels"),
            slide_rels_xml(slide_no, !notes_texts[idx].trim().is_empty()).into_bytes(),
        );
        if !notes_texts[idx].trim().is_empty() {
            parts.insert(
                format!("ppt/notesSlides/notesSlide{slide_no}.xml"),
                render_notes_slide(&notes_texts[idx]).into_bytes(),
            );
        }
    }
    write_new_package(output, &parts)
}

fn ordered_slide_parts(path: &Path) -> Result<Vec<String>> {
    let presentation = parse_xml(&read_part(path, "ppt/presentation.xml")?)?;
    let rels = parse_xml(&read_part(path, "ppt/_rels/presentation.xml.rels")?)?;
    let mut rel_targets = BTreeMap::new();
    for rel in children(&rels, "Relationship") {
        if let (Some(id), Some(target)) = (rel.attributes.get("Id"), rel.attributes.get("Target")) {
            rel_targets.insert(
                id.clone(),
                resolve_relationship_target("ppt/presentation.xml", target),
            );
        }
    }
    let sld_id_lst = child(&presentation, "sldIdLst")
        .ok_or_else(|| anyhow!("ppt/presentation.xml has no slide list"))?;
    let mut parts = Vec::new();
    for sld_id in children(sld_id_lst, "sldId") {
        let rid = sld_id
            .attributes
            .get("r:id")
            .or_else(|| sld_id.attributes.get("id"));
        if let Some(part) = rid.and_then(|rid| rel_targets.get(rid)) {
            parts.push(part.clone());
        }
    }
    Ok(parts)
}

fn collect_text_runs(element: &Element, lines: &mut Vec<String>) {
    for node in &element.children {
        if let XMLNode::Element(el) = node {
            if el.name == "t" {
                let mut text = String::new();
                for child in &el.children {
                    if let XMLNode::Text(value) | XMLNode::CData(value) = child {
                        text.push_str(value);
                    }
                }
                if !text.trim().is_empty() {
                    lines.push(text);
                }
            }
            collect_text_runs(el, lines);
        }
    }
}

fn slide_rels_part(slide_part: &str) -> String {
    let (dir, file) = slide_part
        .rsplit_once('/')
        .unwrap_or(("ppt/slides", slide_part));
    format!("{dir}/_rels/{file}.rels")
}

fn render_slide(
    slide: &SlideSpec,
    brand: &DeckBrand,
    index: usize,
    total: usize,
    footer: Option<&str>,
) -> String {
    let mut shapes = String::new();
    shapes.push_str(&background_rect(&brand.background));
    match slide.layout.as_str() {
        "title" => {
            shapes.push_str(&textbox(
                700_000,
                1_900_000,
                11_000_000,
                1_200_000,
                slide.title.as_deref().unwrap_or(""),
                4000,
                &brand.text,
                true,
            ));
            if let Some(subtitle) = &slide.subtitle {
                shapes.push_str(&textbox(
                    700_000,
                    3_000_000,
                    11_000_000,
                    700_000,
                    subtitle,
                    2200,
                    &brand.accent,
                    false,
                ));
            }
        }
        "section" => {
            if let Some(number) = &slide.number {
                shapes.push_str(&textbox(
                    700_000,
                    1_800_000,
                    2_000_000,
                    1_000_000,
                    number,
                    4400,
                    &brand.accent,
                    true,
                ));
            }
            shapes.push_str(&textbox(
                700_000,
                3_000_000,
                10_800_000,
                1_000_000,
                slide.title.as_deref().unwrap_or(""),
                3200,
                &brand.text,
                true,
            ));
        }
        "two_column" => {
            shapes.push_str(&title_shape(
                slide.title.as_deref().unwrap_or(""),
                &brand.text,
            ));
            shapes.push_str(&textbox(
                800_000,
                1_600_000,
                5_400_000,
                500_000,
                slide.left_title.as_deref().unwrap_or(""),
                1800,
                &brand.accent,
                true,
            ));
            shapes.push_str(&bullet_box(
                800_000,
                2_200_000,
                5_400_000,
                3_900_000,
                slide.left_bullets.as_deref().unwrap_or(&[]),
                &brand.text,
            ));
            shapes.push_str(&textbox(
                6_800_000,
                1_600_000,
                5_400_000,
                500_000,
                slide.right_title.as_deref().unwrap_or(""),
                1800,
                &brand.accent,
                true,
            ));
            shapes.push_str(&bullet_box(
                6_800_000,
                2_200_000,
                5_400_000,
                3_900_000,
                slide.right_bullets.as_deref().unwrap_or(&[]),
                &brand.text,
            ));
        }
        "three_panel" => {
            shapes.push_str(&title_shape(
                slide.title.as_deref().unwrap_or(""),
                &brand.text,
            ));
            let panels = slide.panels.as_deref().unwrap_or(&[]);
            for idx in 0..3 {
                let panel = panels.get(idx);
                let x = 700_000 + idx as i64 * 4_200_000;
                shapes.push_str(&textbox(
                    x,
                    1_650_000,
                    3_700_000,
                    500_000,
                    panel.and_then(|p| p.title.as_deref()).unwrap_or(""),
                    1600,
                    &brand.accent,
                    true,
                ));
                let empty = Vec::new();
                shapes.push_str(&bullet_box(
                    x,
                    2_250_000,
                    3_700_000,
                    3_600_000,
                    panel
                        .map(|p| p.bullets.as_slice())
                        .unwrap_or(empty.as_slice()),
                    &brand.text,
                ));
            }
        }
        "quote" => {
            shapes.push_str(&textbox(
                1_200_000,
                2_000_000,
                10_600_000,
                1_700_000,
                slide.quote.as_deref().unwrap_or(""),
                2800,
                &brand.text,
                false,
            ));
            if let Some(attr) = &slide.attribution {
                shapes.push_str(&textbox(
                    1_200_000,
                    3_900_000,
                    10_600_000,
                    500_000,
                    attr,
                    1600,
                    &brand.muted,
                    false,
                ));
            }
        }
        "table" => {
            shapes.push_str(&title_shape(
                slide.title.as_deref().unwrap_or(""),
                &brand.text,
            ));
            let mut lines = Vec::new();
            if let Some(headers) = &slide.headers {
                lines.push(headers.join(" | "));
            }
            if let Some(rows) = &slide.rows {
                for row in rows {
                    lines.push(row.join(" | "));
                }
            }
            shapes.push_str(&textbox(
                800_000,
                1_700_000,
                11_500_000,
                4_800_000,
                &lines.join("\n"),
                1400,
                &brand.text,
                false,
            ));
        }
        "image_placeholder" => {
            shapes.push_str(&title_shape(
                slide.title.as_deref().unwrap_or(""),
                &brand.text,
            ));
            let mut text = slide
                .placeholder_text
                .clone()
                .unwrap_or_else(|| "Image placeholder".to_string());
            if let Some(prompt) = &slide.image_prompt {
                text.push('\n');
                text.push_str(prompt);
            }
            shapes.push_str(&textbox(
                1_200_000,
                2_000_000,
                10_500_000,
                2_800_000,
                &text,
                1800,
                &brand.muted,
                false,
            ));
        }
        _ => {
            shapes.push_str(&title_shape(
                slide.title.as_deref().unwrap_or(""),
                &brand.text,
            ));
            if let Some(body) = &slide.body {
                shapes.push_str(&textbox(
                    800_000,
                    1_700_000,
                    11_600_000,
                    1_500_000,
                    body,
                    1800,
                    &brand.text,
                    false,
                ));
            }
            shapes.push_str(&bullet_box(
                800_000,
                2_600_000,
                11_600_000,
                3_500_000,
                slide.bullets.as_deref().unwrap_or(&[]),
                &brand.text,
            ));
        }
    }
    if let Some(footer) = footer {
        shapes.push_str(&textbox(
            800_000,
            6_650_000,
            8_500_000,
            250_000,
            footer,
            900,
            &brand.muted,
            false,
        ));
    }
    shapes.push_str(&textbox(
        11_700_000,
        6_650_000,
        900_000,
        250_000,
        &format!("{index} / {total}"),
        900,
        &brand.muted,
        false,
    ));
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>{shapes}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"#
    )
}

fn title_shape(title: &str, color: &str) -> String {
    textbox(
        650_000, 450_000, 11_600_000, 650_000, title, 2400, color, true,
    )
}

fn background_rect(color: &str) -> String {
    format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="2" name="Background"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="12192000" cy="6858000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="{}"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>"#,
        escaped(color)
    )
}

#[allow(clippy::too_many_arguments)]
fn textbox(
    x: i64,
    y: i64,
    w: i64,
    h: i64,
    text: &str,
    size: usize,
    color: &str,
    bold: bool,
) -> String {
    let paragraphs = text
        .lines()
        .map(|line| text_paragraph(line, size, color, bold))
        .collect::<Vec<_>>()
        .join("");
    format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="{}" name="TextBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FFFFFF"><a:alpha val="0"/></a:srgbClr></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square"/><a:lstStyle/>{paragraphs}</p:txBody></p:sp>"#,
        shape_id()
    )
}

fn bullet_box(x: i64, y: i64, w: i64, h: i64, bullets: &[Bullet], color: &str) -> String {
    let mut lines = Vec::new();
    for bullet in bullets {
        match bullet {
            Bullet::Text(text) => lines.push(format!("- {text}")),
            Bullet::Rich { text, sub } => {
                lines.push(format!("- {text}"));
                if let Some(sub) = sub {
                    for item in sub {
                        lines.push(format!("  - {item}"));
                    }
                }
            }
        }
    }
    textbox(x, y, w, h, &lines.join("\n"), 1500, color, false)
}

fn text_paragraph(text: &str, size: usize, color: &str, bold: bool) -> String {
    format!(
        r#"<a:p><a:r><a:rPr lang="en-US" sz="{size}"{}><a:solidFill><a:srgbClr val="{}"/></a:solidFill></a:rPr><a:t>{}</a:t></a:r></a:p>"#,
        if bold { r#" b="1""# } else { "" },
        escaped(color),
        escaped(text)
    )
}

fn base_pptx_parts(slide_count: usize, notes_texts: &[String]) -> PartMap {
    let mut parts = PartMap::new();
    parts.insert(
        "[Content_Types].xml".to_string(),
        pptx_content_types(slide_count, notes_texts).into_bytes(),
    );
    parts.insert("_rels/.rels".to_string(), pptx_root_rels().into_bytes());
    parts.insert(
        "ppt/presentation.xml".to_string(),
        presentation_xml(slide_count).into_bytes(),
    );
    parts.insert(
        "ppt/_rels/presentation.xml.rels".to_string(),
        presentation_rels_xml(slide_count).into_bytes(),
    );
    parts.insert("docProps/core.xml".to_string(), br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:creator>office-tools</dc:creator></cp:coreProperties>"#.to_vec());
    parts.insert("docProps/app.xml".to_string(), br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>office-tools</Application><PresentationFormat>Widescreen</PresentationFormat></Properties>"#.to_vec());
    parts
}

fn pptx_content_types(slide_count: usize, notes_texts: &[String]) -> String {
    let mut overrides = String::new();
    for idx in 1..=slide_count {
        overrides.push_str(&format!(r#"<Override PartName="/ppt/slides/slide{idx}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>"#));
        if notes_texts
            .get(idx - 1)
            .is_some_and(|text| !text.trim().is_empty())
        {
            overrides.push_str(&format!(r#"<Override PartName="/ppt/notesSlides/notesSlide{idx}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>"#));
        }
    }
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>{overrides}<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>"#
    )
}

fn pptx_root_rels() -> String {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>"#.to_string()
}

fn presentation_xml(slide_count: usize) -> String {
    let mut slide_ids = String::new();
    for idx in 1..=slide_count {
        slide_ids.push_str(&format!(
            r#"<p:sldId id="{}" r:id="rId{}"/>"#,
            255 + idx,
            idx
        ));
    }
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst>{slide_ids}</p:sldIdLst><p:sldSz cx="12192000" cy="6858000" type="screen16x9"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>"#
    )
}

fn presentation_rels_xml(slide_count: usize) -> String {
    let mut rels = String::new();
    for idx in 1..=slide_count {
        rels.push_str(&format!(r#"<Relationship Id="rId{idx}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{idx}.xml"/>"#));
    }
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{rels}</Relationships>"#
    )
}

fn slide_rels_xml(slide_no: usize, has_notes: bool) -> String {
    let notes_rel = if has_notes {
        format!(
            r#"<Relationship Id="rIdNotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide{slide_no}.xml"/>"#
        )
    } else {
        String::new()
    };
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{notes_rel}</Relationships>"#
    )
}

fn slide_notes_text(slide: &SlideSpec) -> String {
    let mut parts = Vec::new();
    if let Some(notes) = &slide.notes {
        parts.push(notes.clone());
    }
    if slide.layout == "image_placeholder"
        && let Some(prompt) = &slide.image_prompt
    {
        parts.push(format!("Image prompt: {prompt}"));
    }
    parts.join("\n")
}

fn render_notes_slide(text: &str) -> String {
    let paragraphs = text
        .lines()
        .map(|line| text_paragraph(line, 1200, "000000", false))
        .collect::<Vec<_>>()
        .join("");
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="685800"/><a:ext cx="5486400" cy="5486400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr wrap="square"/><a:lstStyle/>{paragraphs}</p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>"#
    )
}

fn default_layout() -> String {
    "content".to_string()
}

fn default_font() -> String {
    "Arial".to_string()
}

fn default_bg() -> String {
    "FFFFFF".to_string()
}

fn default_text() -> String {
    "1F2A44".to_string()
}

fn default_accent() -> String {
    "2F6F9F".to_string()
}

fn default_muted() -> String {
    "666666".to_string()
}

fn shape_id() -> usize {
    use std::sync::atomic::{AtomicUsize, Ordering};
    static NEXT: AtomicUsize = AtomicUsize::new(10);
    NEXT.fetch_add(1, Ordering::Relaxed)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn build_and_read_simple_deck() -> Result<()> {
        let dir = tempfile::tempdir()?;
        let path = dir.path().join("deck.pptx");
        let spec = DeckSpec {
            meta: DeckMeta {
                title: Some("Deck".to_string()),
                subtitle: None,
                author: None,
                date: None,
                footer_text: Some("Footer".to_string()),
            },
            brand: DeckBrand::default(),
            slides: vec![SlideSpec {
                layout: "content".to_string(),
                title: Some("Slide title".to_string()),
                subtitle: None,
                number: None,
                body: Some("Body".to_string()),
                bullets: Some(vec![Bullet::Text("Point".to_string())]),
                left_title: None,
                left_bullets: None,
                right_title: None,
                right_bullets: None,
                panels: None,
                quote: None,
                attribution: None,
                headers: None,
                rows: None,
                placeholder_text: None,
                image_prompt: None,
                notes: Some("Speaker note".to_string()),
            }],
        };
        build_deck(&spec, &path)?;
        let slides = read_slides(&path)?;
        let notes = read_notes(&path)?;
        assert_eq!(slides.len(), 1);
        assert!(slides[0].text.contains("Slide title"));
        assert!(slides[0].text.contains("Point"));
        assert_eq!(notes.len(), 1);
        assert!(notes[0].text.contains("Speaker note"));
        Ok(())
    }
}
