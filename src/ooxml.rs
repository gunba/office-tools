use anyhow::{Context, Result, anyhow, bail};
use std::collections::BTreeMap;
use std::fs::{self, File};
use std::io::{Cursor, Read, Write};
use std::path::{Path, PathBuf};
use xmltree::{Element, EmitterConfig, XMLNode};
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

pub type PartMap = BTreeMap<String, Vec<u8>>;

#[derive(Debug, Clone, serde::Serialize)]
pub struct PackagePart {
    pub name: String,
    pub content_type: Option<String>,
    pub compressed_size: u64,
    pub uncompressed_size: u64,
}

pub fn list_parts(path: impl AsRef<Path>) -> Result<Vec<PackagePart>> {
    let mut archive = open_archive(path)?;
    let content_types = read_content_types(&mut archive)?;
    let mut parts = Vec::with_capacity(archive.len());
    for index in 0..archive.len() {
        let file = archive.by_index(index)?;
        let name = file.name().to_string();
        parts.push(PackagePart {
            content_type: content_type_for(&content_types, &name),
            name,
            compressed_size: file.compressed_size(),
            uncompressed_size: file.size(),
        });
    }
    Ok(parts)
}

#[derive(Default)]
struct ContentTypes {
    defaults: BTreeMap<String, String>,
    overrides: BTreeMap<String, String>,
}

fn read_content_types(archive: &mut ZipArchive<File>) -> Result<ContentTypes> {
    let Ok(mut file) = archive.by_name("[Content_Types].xml") else {
        return Ok(ContentTypes::default());
    };
    let mut data = Vec::new();
    file.read_to_end(&mut data)?;
    let root = parse_xml(&data).context("parse [Content_Types].xml")?;
    let mut content_types = ContentTypes::default();
    for default in children(&root, "Default") {
        if let (Some(extension), Some(content_type)) = (
            default.attributes.get("Extension"),
            default.attributes.get("ContentType"),
        ) {
            content_types
                .defaults
                .insert(extension.to_ascii_lowercase(), content_type.clone());
        }
    }
    for override_part in children(&root, "Override") {
        if let (Some(part_name), Some(content_type)) = (
            override_part.attributes.get("PartName"),
            override_part.attributes.get("ContentType"),
        ) {
            content_types.overrides.insert(
                normalize_part_name(part_name).to_string(),
                content_type.clone(),
            );
        }
    }
    Ok(content_types)
}

fn content_type_for(content_types: &ContentTypes, name: &str) -> Option<String> {
    let normalized = normalize_part_name(name);
    content_types
        .overrides
        .get(normalized)
        .cloned()
        .or_else(|| {
            normalized
                .rsplit_once('.')
                .and_then(|(_, extension)| {
                    content_types.defaults.get(&extension.to_ascii_lowercase())
                })
                .cloned()
        })
}

pub fn read_part(path: impl AsRef<Path>, name: &str) -> Result<Vec<u8>> {
    let mut archive = open_archive(path)?;
    let mut file = archive
        .by_name(normalize_part_name(name))
        .with_context(|| format!("package part not found: {name}"))?;
    let mut buf = Vec::new();
    file.read_to_end(&mut buf)?;
    Ok(buf)
}

pub fn read_part_to_string(path: impl AsRef<Path>, name: &str) -> Result<String> {
    let data = read_part(path, name)?;
    String::from_utf8(data).with_context(|| format!("package part is not UTF-8 XML/text: {name}"))
}

pub fn read_xml_part(path: impl AsRef<Path>, name: &str) -> Result<Element> {
    parse_xml(&read_part(path, name)?).with_context(|| format!("failed to parse XML part: {name}"))
}

pub fn parse_xml(data: &[u8]) -> Result<Element> {
    Element::parse(Cursor::new(data)).map_err(|err| anyhow!(err))
}

pub fn write_xml(element: &Element) -> Result<Vec<u8>> {
    let mut buf = Vec::new();
    element
        .write_with_config(
            &mut buf,
            EmitterConfig::new()
                .write_document_declaration(true)
                .perform_indent(false),
        )
        .map_err(|err| anyhow!(err))?;
    Ok(buf)
}

pub fn rewrite_parts(
    path: impl AsRef<Path>,
    updates: &PartMap,
    deletes: impl IntoIterator<Item = impl AsRef<str>>,
) -> Result<()> {
    let path = path.as_ref();
    let delete_set = deletes
        .into_iter()
        .map(|s| normalize_part_name(s.as_ref()).to_string())
        .collect::<std::collections::BTreeSet<_>>();
    let parent = path.parent().unwrap_or_else(|| Path::new("."));
    let mut tmp = tempfile::Builder::new()
        .prefix(
            path.file_name()
                .and_then(|name| name.to_str())
                .unwrap_or("office-tools"),
        )
        .suffix(".tmp")
        .tempfile_in(parent)?;

    {
        let input =
            File::open(path).with_context(|| format!("open package: {}", path.display()))?;
        let mut zin = ZipArchive::new(input)?;
        let mut zout = ZipWriter::new(tmp.as_file_mut());
        let options = SimpleFileOptions::default()
            .compression_method(CompressionMethod::Deflated)
            .large_file(true);
        let mut written = std::collections::BTreeSet::new();

        for index in 0..zin.len() {
            let mut entry = zin.by_index(index)?;
            let name = entry.name().to_string();
            if delete_set.contains(&name) || written.contains(&name) {
                continue;
            }
            let data = if let Some(update) = updates.get(&name) {
                update.clone()
            } else {
                let mut data = Vec::new();
                entry.read_to_end(&mut data)?;
                data
            };
            zout.start_file(&name, options)?;
            zout.write_all(&data)?;
            written.insert(name);
        }

        for (name, data) in updates {
            let normalized = normalize_part_name(name);
            if delete_set.contains(normalized) || written.contains(normalized) {
                continue;
            }
            zout.start_file(normalized, options)?;
            zout.write_all(data)?;
            written.insert(normalized.to_string());
        }
        zout.finish()?;
    }

    tmp.persist(path)
        .map_err(|err| anyhow!("replace package {}: {}", path.display(), err.error))?;
    Ok(())
}

pub fn write_new_package(path: impl AsRef<Path>, parts: &PartMap) -> Result<()> {
    let path = path.as_ref();
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent)?;
    }
    let output =
        File::create(path).with_context(|| format!("create package: {}", path.display()))?;
    let mut writer = ZipWriter::new(output);
    let options = SimpleFileOptions::default()
        .compression_method(CompressionMethod::Deflated)
        .large_file(true);
    for (name, data) in parts {
        writer.start_file(normalize_part_name(name), options)?;
        writer.write_all(data)?;
    }
    writer.finish()?;
    Ok(())
}

pub fn normalize_part_name(name: &str) -> &str {
    name.strip_prefix('/').unwrap_or(name)
}

pub fn resolve_relationship_target(source_part: &str, target: &str) -> String {
    if let Some(stripped) = target.strip_prefix('/') {
        return normalize_posix(stripped);
    }
    let base = source_part
        .rsplit_once('/')
        .map(|(dir, _)| dir)
        .unwrap_or("");
    normalize_posix(&format!("{base}/{target}"))
}

pub fn relationship_target_for_part(source_part: &str, part_name: &str) -> String {
    let base = source_part
        .rsplit_once('/')
        .map(|(dir, _)| dir)
        .unwrap_or("");
    let base_segments = base
        .split('/')
        .filter(|s| !s.is_empty())
        .collect::<Vec<_>>();
    let part_segments = part_name
        .split('/')
        .filter(|s| !s.is_empty())
        .collect::<Vec<_>>();

    let mut common = 0;
    while common < base_segments.len()
        && common < part_segments.len()
        && base_segments[common] == part_segments[common]
    {
        common += 1;
    }

    let mut out = Vec::new();
    for _ in common..base_segments.len() {
        out.push("..".to_string());
    }
    for segment in &part_segments[common..] {
        out.push((*segment).to_string());
    }
    if out.is_empty() {
        ".".to_string()
    } else {
        out.join("/")
    }
}

pub fn normalize_posix(path: &str) -> String {
    let mut stack = Vec::new();
    for segment in path.split('/') {
        match segment {
            "" | "." => {}
            ".." => {
                stack.pop();
            }
            other => stack.push(other),
        }
    }
    stack.join("/")
}

pub fn element_text(element: &Element) -> String {
    let mut out = String::new();
    collect_text(element, &mut out);
    out
}

pub fn collect_text(element: &Element, out: &mut String) {
    for child in &element.children {
        match child {
            XMLNode::Text(text) | XMLNode::CData(text) => out.push_str(text),
            XMLNode::Element(el) => collect_text(el, out),
            _ => {}
        }
    }
}

pub fn child<'a>(element: &'a Element, name: &str) -> Option<&'a Element> {
    element.children.iter().find_map(|node| match node {
        XMLNode::Element(el) if el.name == name => Some(el),
        _ => None,
    })
}

pub fn child_mut<'a>(element: &'a mut Element, name: &str) -> Option<&'a mut Element> {
    element.children.iter_mut().find_map(|node| match node {
        XMLNode::Element(el) if el.name == name => Some(el),
        _ => None,
    })
}

pub fn children<'a>(element: &'a Element, name: &'a str) -> impl Iterator<Item = &'a Element> + 'a {
    element.children.iter().filter_map(move |node| match node {
        XMLNode::Element(el) if el.name == name => Some(el),
        _ => None,
    })
}

pub fn children_mut<'a>(
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

pub fn find_descendants<'a>(element: &'a Element, name: &str, out: &mut Vec<&'a Element>) {
    for child in &element.children {
        if let XMLNode::Element(el) = child {
            if el.name == name {
                out.push(el);
            }
            find_descendants(el, name, out);
        }
    }
}

pub fn find_descendants_mut(element: &mut Element, name: &str, out: &mut Vec<*mut Element>) {
    for child in &mut element.children {
        if let XMLNode::Element(el) = child {
            if el.name == name {
                out.push(el as *mut Element);
            }
            find_descendants_mut(el, name, out);
        }
    }
}

pub fn take_children(element: &mut Element, names: &[&str]) {
    element.children.retain(|node| match node {
        XMLNode::Element(el) => !names.contains(&el.name.as_str()),
        _ => true,
    });
}

pub fn escaped(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    for ch in text.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&apos;"),
            _ => out.push(ch),
        }
    }
    out
}

pub fn safe_filename_fragment(text: &str, max_len: usize) -> String {
    let mut out = String::new();
    for ch in text.chars() {
        if ch.is_ascii_alphanumeric() || matches!(ch, ' ' | '-' | '_') {
            out.push(ch);
        } else {
            out.push('_');
        }
        if out.len() >= max_len {
            break;
        }
    }
    let trimmed = out.trim();
    if trimmed.is_empty() {
        "untitled".to_string()
    } else {
        trimmed.to_string()
    }
}

pub fn package_extension(path: &Path) -> Result<String> {
    path.extension()
        .and_then(|ext| ext.to_str())
        .map(|ext| ext.to_ascii_lowercase())
        .ok_or_else(|| anyhow!("missing file extension"))
}

pub fn ensure_office_extension(path: &Path, allowed: &[&str]) -> Result<()> {
    let ext = path
        .extension()
        .and_then(|ext| ext.to_str())
        .map(|ext| ext.to_ascii_lowercase())
        .ok_or_else(|| anyhow!("missing file extension: {}", path.display()))?;
    if allowed.iter().any(|allowed| ext == *allowed) {
        Ok(())
    } else {
        bail!(
            "unsupported extension .{} for {}; expected one of {:?}",
            ext,
            path.display(),
            allowed
        )
    }
}

pub fn absolute(path: impl AsRef<Path>) -> Result<PathBuf> {
    let path = path.as_ref();
    if path.is_absolute() {
        Ok(path.to_path_buf())
    } else {
        Ok(std::env::current_dir()?.join(path))
    }
}

fn open_archive(path: impl AsRef<Path>) -> Result<ZipArchive<File>> {
    let path = path.as_ref();
    let file = File::open(path).with_context(|| format!("open package: {}", path.display()))?;
    ZipArchive::new(file).with_context(|| format!("read zip package: {}", path.display()))
}
