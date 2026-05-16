use std::fs;
use std::path::{Path, PathBuf};

#[test]
fn repository_contains_no_python_payloads_or_legacy_corpus_files() -> anyhow::Result<()> {
    let root = repo_root();
    let files = repo_files(&root)?;
    let legacy_prefix = concat!("a", "t", "o");
    let offenders = files
        .iter()
        .filter(|path| {
            let name = path
                .file_name()
                .and_then(|name| name.to_str())
                .unwrap_or_default()
                .to_ascii_lowercase();
            let ext = path
                .extension()
                .and_then(|ext| ext.to_str())
                .unwrap_or_default()
                .to_ascii_lowercase();
            matches!(ext.as_str(), "py" | "pyc" | "whl" | "pyd")
                || name.contains("python")
                || name.starts_with(legacy_prefix)
        })
        .collect::<Vec<_>>();
    assert!(
        offenders.is_empty(),
        "unexpected payload files: {offenders:#?}"
    );
    Ok(())
}

#[test]
fn repository_contains_no_private_brand_or_legacy_corpus_text() -> anyhow::Result<()> {
    let root = repo_root();
    let mut offenders = Vec::new();
    let private_brand_upper = concat!("R", "S", "M");
    let private_brand_lower = concat!("r", "s", "m");
    let legacy_corpus_name = concat!("a", "t", "o", "-", "m", "c", "p");
    for path in repo_files(&root)? {
        let Ok(text) = fs::read_to_string(&path) else {
            continue;
        };
        if text.contains(private_brand_upper)
            || text.contains(private_brand_lower)
            || text.contains(legacy_corpus_name)
        {
            offenders.push(path);
        }
    }
    assert!(
        offenders.is_empty(),
        "unexpected text matches: {offenders:#?}"
    );
    Ok(())
}

#[test]
fn outlook_implementation_has_no_send_reply_forward_or_delete_surface() -> anyhow::Result<()> {
    let root = repo_root();
    let outlook = fs::read_to_string(root.join("src/outlook.rs"))?;
    let wincom = fs::read_to_string(root.join("src/wincom.rs"))?;
    let combined = format!("{outlook}\n{wincom}");
    for forbidden in [
        ".Send(",
        ".Reply(",
        ".ReplyAll(",
        ".Forward(",
        "SendUsingAccount",
    ] {
        assert!(
            !combined.contains(forbidden),
            "Outlook implementation contains forbidden operation {forbidden}"
        );
    }
    Ok(())
}

fn repo_root() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
}

fn repo_files(root: &Path) -> anyhow::Result<Vec<PathBuf>> {
    let mut files = Vec::new();
    visit(root, root, &mut files)?;
    Ok(files)
}

fn visit(root: &Path, current: &Path, files: &mut Vec<PathBuf>) -> anyhow::Result<()> {
    for entry in fs::read_dir(current)? {
        let entry = entry?;
        let path = entry.path();
        let relative = path.strip_prefix(root).unwrap_or(&path);
        let first = relative
            .components()
            .next()
            .and_then(|component| match component {
                std::path::Component::Normal(value) => value.to_str(),
                _ => None,
            });
        if matches!(first, Some(".git" | "target")) {
            continue;
        }
        if path.is_dir() {
            visit(root, &path, files)?;
        } else {
            files.push(path);
        }
    }
    Ok(())
}
