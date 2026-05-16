use anyhow::Result;
use clap::{Parser, Subcommand};

#[derive(Debug, Parser)]
#[command(name = "office-tools")]
#[command(version, about = "Rust Office document primitives and MCP server")]
struct Cli {
    #[command(subcommand)]
    command: Command,
}

#[derive(Debug, Subcommand)]
enum Command {
    /// Excel / XLSX commands, prioritising direct OOXML preservation.
    Xlsx(office_tools::xlsx::XlsxArgs),
    /// Word / DOCX commands.
    Docx(office_tools::docx::DocxArgs),
    /// PowerPoint / PPTX commands.
    Pptx(office_tools::pptx::PptxArgs),
    /// Read-only Outlook commands over WinCOM.
    Outlook(office_tools::outlook::OutlookArgs),
    /// Raw OOXML zip package primitives for agent scripts.
    Package(office_tools::script::PackageArgs),
    /// Non-destructive Windows Office COM availability check.
    Doctor(office_tools::doctor::DoctorArgs),
    /// Model Context Protocol server.
    Mcp(office_tools::mcp::McpArgs),
}

fn main() -> Result<()> {
    match Cli::parse().command {
        Command::Xlsx(args) => args.run(),
        Command::Docx(args) => args.run(),
        Command::Pptx(args) => args.run(),
        Command::Outlook(args) => args.run(),
        Command::Package(args) => args.run(),
        Command::Doctor(args) => args.run(),
        Command::Mcp(args) => args.run(),
    }
}
