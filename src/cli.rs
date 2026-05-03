//! clap definitions for the `docxai` command-line surface.
//!
//! The PRD pins the CLI to exactly five verbs: `snapshot`, `add`, `set`,
//! `delete`, `styles`. Adding a new verb is a v1.0 breaking change and
//! requires explicit approval (PRD §7.1).

use std::path::PathBuf;

use clap::{Args, Parser, Subcommand};

#[derive(Debug, Parser)]
#[command(
    name = "docxai",
    version,
    about = "Create and modify .docx files for AI agents.",
    long_about = "docxai is a deterministic CLI that exposes five verbs \
                  (snapshot, add, set, delete, styles) for editing .docx \
                  files. Designed for use by AI coding agents."
)]
pub struct Cli {
    #[command(subcommand)]
    pub command: Command,
}

#[derive(Debug, Subcommand)]
pub enum Command {
    /// Print a JSON snapshot of the document body, styles, and refs.
    Snapshot(SnapshotArgs),

    /// Append or insert a new element (paragraph, table, image, equation).
    Add(AddArgs),

    /// Modify an existing element identified by a ref (@p3, @t1.r2.c3, ...).
    Set(SetArgs),

    /// Remove an existing element identified by a ref.
    Delete(DeleteArgs),

    /// List the named styles available in the document's `styles.xml`.
    Styles(StylesArgs),
}

#[derive(Debug, Args)]
pub struct SnapshotArgs {
    /// Path to the .docx file.
    pub file: PathBuf,

    /// Pretty-print the JSON snapshot.
    #[arg(long)]
    pub pretty: bool,

    /// Drill into a specific table and return all its cells.
    #[arg(long, value_name = "REF")]
    pub table: Option<String>,
}

#[derive(Debug, Args)]
pub struct AddArgs {
    /// Path to the .docx file.
    pub file: PathBuf,

    #[command(subcommand)]
    pub kind: AddKind,
}

#[derive(Debug, Subcommand)]
pub enum AddKind {
    /// Add a paragraph with optional named style.
    Paragraph(AddParagraphArgs),

    /// Add a table with given dimensions and optional header row.
    Table(AddTableArgs),

    /// Add an inline image at a given width.
    Image(AddImageArgs),

    /// Add a display equation from LaTeX source.
    Equation(AddEquationArgs),
}

#[derive(Debug, Args)]
pub struct AddParagraphArgs {
    /// Paragraph text in markdown subset (see `docxai --help`).
    #[arg(long)]
    pub text: String,

    /// Named style from `available_styles` (see `docxai styles`).
    #[arg(long)]
    pub style: Option<String>,

    #[command(flatten)]
    pub position: PositionArgs,
}

#[derive(Debug, Args)]
pub struct AddTableArgs {
    #[arg(long)]
    pub rows: u32,

    #[arg(long)]
    pub cols: u32,

    /// Comma-separated header cells, e.g. `--header "Metric,Q3,Q4"`.
    #[arg(long)]
    pub header: Option<String>,

    #[command(flatten)]
    pub position: PositionArgs,
}

#[derive(Debug, Args)]
pub struct AddImageArgs {
    #[arg(long)]
    pub path: PathBuf,

    /// Width with unit suffix (`12cm`, `4.5in`, `300px`).
    #[arg(long)]
    pub width: Option<String>,

    #[arg(long)]
    pub caption: Option<String>,

    #[command(flatten)]
    pub position: PositionArgs,
}

#[derive(Debug, Args)]
pub struct AddEquationArgs {
    #[arg(long)]
    pub latex: String,

    #[command(flatten)]
    pub position: PositionArgs,
}

/// Mutually exclusive insertion anchors for `add`.
#[derive(Debug, Args)]
#[group(multiple = false)]
pub struct PositionArgs {
    #[arg(long, value_name = "REF")]
    pub after: Option<String>,

    #[arg(long, value_name = "REF")]
    pub before: Option<String>,
}

#[derive(Debug, Args)]
pub struct SetArgs {
    pub file: PathBuf,

    /// Ref to mutate (@p3, @t1.r2.c3, @i1, @e1).
    #[arg(value_name = "REF")]
    pub reference: String,

    #[arg(long)]
    pub text: Option<String>,

    #[arg(long)]
    pub style: Option<String>,

    #[arg(long)]
    pub width: Option<String>,

    #[arg(long)]
    pub caption: Option<String>,

    #[arg(long)]
    pub latex: Option<String>,
}

#[derive(Debug, Args)]
pub struct DeleteArgs {
    pub file: PathBuf,

    #[arg(value_name = "REF")]
    pub reference: String,
}

#[derive(Debug, Args)]
pub struct StylesArgs {
    pub file: PathBuf,
}

#[cfg(test)]
mod tests {
    use super::*;
    use clap::CommandFactory;

    #[test]
    fn cli_definition_is_internally_consistent() {
        Cli::command().debug_assert();
    }

    #[test]
    fn five_verbs_exist_at_top_level() {
        let cmd = Cli::command();
        let verbs: Vec<&str> = cmd.get_subcommands().map(|s| s.get_name()).collect();
        assert_eq!(
            verbs,
            vec!["snapshot", "add", "set", "delete", "styles"],
            "PRD §7.1 freezes the verb list to exactly these five"
        );
    }

    #[test]
    fn add_has_four_kinds() {
        let cmd = Cli::command();
        let add = cmd.find_subcommand("add").expect("add subcommand");
        let kinds: Vec<&str> = add.get_subcommands().map(|s| s.get_name()).collect();
        assert_eq!(kinds, vec!["paragraph", "table", "image", "equation"]);
    }

    #[test]
    fn after_and_before_are_mutually_exclusive() {
        let result = Cli::try_parse_from([
            "docxai",
            "add",
            "doc.docx",
            "paragraph",
            "--text",
            "x",
            "--after",
            "@p1",
            "--before",
            "@p2",
        ]);
        assert!(result.is_err(), "--after and --before must conflict");
    }
}
