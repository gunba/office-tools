use crate::ooxml::{PartMap, list_parts, read_part, rewrite_parts};
use anyhow::{Context, Result, bail};
use base64::Engine;
use base64::engine::general_purpose::STANDARD as BASE64_STANDARD;
use clap::{Args, Subcommand};
use std::fs;
use std::path::PathBuf;

#[derive(Debug, Args)]
pub struct PackageArgs {
    #[command(subcommand)]
    command: PackageCommand,
}

#[derive(Debug, Subcommand)]
pub enum PackageCommand {
    /// List OOXML zip package parts.
    ListParts {
        file: PathBuf,
        #[arg(long)]
        json: bool,
    },
    /// Read a package part to stdout.
    ReadPart {
        file: PathBuf,
        part: String,
        #[arg(short, long)]
        output: Option<PathBuf>,
        #[arg(long)]
        base64: bool,
    },
    /// Replace or create a package part from a file.
    WritePart {
        file: PathBuf,
        part: String,
        #[arg(short, long)]
        input: Option<PathBuf>,
        #[arg(long)]
        base64: Option<String>,
        #[arg(long)]
        text: Option<String>,
    },
    /// Delete one or more package parts.
    DeletePart { file: PathBuf, parts: Vec<String> },
}

impl PackageArgs {
    pub fn run(self) -> Result<()> {
        match self.command {
            PackageCommand::ListParts { file, json } => {
                let parts = list_parts(&file)?;
                if json {
                    println!("{}", serde_json::to_string_pretty(&parts)?);
                } else {
                    for part in parts {
                        let content_type = part.content_type.as_deref().unwrap_or("unknown");
                        println!(
                            "{}\t{}\t{} bytes\t{} compressed",
                            part.name, content_type, part.uncompressed_size, part.compressed_size
                        );
                    }
                }
                Ok(())
            }
            PackageCommand::ReadPart {
                file,
                part,
                output,
                base64,
            } => {
                let data = read_part(&file, &part)?;
                let output_data = if base64 {
                    BASE64_STANDARD.encode(data).into_bytes()
                } else {
                    data
                };
                if let Some(output) = output {
                    fs::write(&output, output_data)
                        .with_context(|| format!("write {}", output.display()))?;
                } else {
                    print!("{}", String::from_utf8_lossy(&output_data));
                }
                Ok(())
            }
            PackageCommand::WritePart {
                file,
                part,
                input,
                base64,
                text,
            } => {
                let sources = usize::from(input.is_some())
                    + usize::from(base64.is_some())
                    + usize::from(text.is_some());
                if sources != 1 {
                    bail!("provide exactly one of --input, --base64, or --text");
                }
                let data = match (input, base64, text) {
                    (Some(input), None, None) => {
                        fs::read(&input).with_context(|| format!("read {}", input.display()))?
                    }
                    (None, Some(base64), None) => BASE64_STANDARD
                        .decode(base64)
                        .context("decode base64 package part")?,
                    (None, None, Some(text)) => text.into_bytes(),
                    _ => unreachable!("source count checked above"),
                };
                let mut updates = PartMap::new();
                updates.insert(part.clone(), data);
                rewrite_parts(&file, &updates, Vec::<String>::new())?;
                println!("Updated {part} in {}", file.display());
                Ok(())
            }
            PackageCommand::DeletePart { file, parts } => {
                if parts.is_empty() {
                    bail!("provide at least one part to delete");
                }
                let updates = PartMap::new();
                rewrite_parts(&file, &updates, parts)?;
                println!("Deleted package part(s) from {}", file.display());
                Ok(())
            }
        }
    }
}
