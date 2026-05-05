//! docxai library entry point.
//!
//! Public modules expose the CLI definitions and error model; the binary in
//! `src/main.rs` only handles process boundaries (argv parsing, exit codes).

pub mod cli;
pub mod doc;
pub mod error;
pub mod markdown;
pub mod mutate;
pub mod refs;
pub mod snapshot;
pub mod styles;

use std::io::{self, Write};

use serde_json::json;

use crate::cli::{AddKind, Cli, Command, SnapshotArgs, StylesArgs};
use crate::doc::Doc;
use crate::error::DocxaiError;
use crate::refs::Ref;

/// Dispatch a parsed [`Cli`] to its handler.
pub fn run(cli: Cli) -> Result<(), DocxaiError> {
    match cli.command {
        Command::Snapshot(args) => run_snapshot(args, &mut io::stdout().lock()),
        Command::Add(args) => run_add(args, &mut io::stdout().lock()),
        Command::Set(args) => run_set(args, &mut io::stdout().lock()),
        Command::Delete(args) => run_delete(args, &mut io::stdout().lock()),
        Command::Styles(args) => run_styles(args, &mut io::stdout().lock()),
    }
}

/// Implementation of `docxai snapshot <FILE> [--pretty] [--table @tN]` (PRD §8).
///
/// Emits a single-line JSON object by default (PRD §10.2). `--pretty`
/// enables indented output. `--table @tN` switches to the drilled view (§8.3).
pub fn run_snapshot(args: SnapshotArgs, writer: &mut dyn Write) -> Result<(), DocxaiError> {
    let doc = Doc::load(&args.file)?;
    if let Some(table_ref) = args.table.as_deref() {
        let snap = snapshot::build_table_snapshot(&doc, table_ref)?;
        write_json(writer, &snap, args.pretty)
    } else {
        let snap = snapshot::build_snapshot(&doc)?;
        write_json(writer, &snap, args.pretty)
    }
}

fn write_json<T: serde::Serialize>(
    writer: &mut dyn Write,
    value: &T,
    pretty: bool,
) -> Result<(), DocxaiError> {
    let result = if pretty {
        serde_json::to_writer_pretty(&mut *writer, value)
    } else {
        serde_json::to_writer(&mut *writer, value)
    };
    result.map_err(|e| DocxaiError::Generic(format!("serialize snapshot: {e}")))?;
    writeln!(writer).map_err(|e| DocxaiError::Generic(format!("write snapshot: {e}")))?;
    Ok(())
}

/// Implementation of `docxai styles <FILE>`.
///
/// Output (PRD #9): single-line JSON `{"styles":[...]}` on stdout.
/// `writer` is injected so tests can capture the output without spawning a process.
pub fn run_styles(args: StylesArgs, writer: &mut dyn Write) -> Result<(), DocxaiError> {
    let doc = Doc::load(&args.file)?;
    let styles = styles::list_paragraph_styles(&doc)?;
    let payload = json!({ "styles": styles });
    serde_json::to_writer(&mut *writer, &payload)
        .map_err(|e| DocxaiError::Generic(format!("serialize styles output: {e}")))?;
    writeln!(writer).map_err(|e| DocxaiError::Generic(format!("write styles output: {e}")))?;
    Ok(())
}

pub fn run_add(args: cli::AddArgs, writer: &mut dyn Write) -> Result<(), DocxaiError> {
    match args.kind {
        AddKind::Paragraph(para) => {
            let mut doc = Doc::load(&args.file)?;
            let result = mutate::add_paragraph(
                &mut doc,
                &para.text,
                para.style.as_deref(),
                para.position.after.as_deref(),
                para.position.before.as_deref(),
            )?;
            writeln!(writer, "{result}")
                .map_err(|e| DocxaiError::Generic(format!("write: {e}")))?;
            Ok(())
        }
        AddKind::Table(_) => Err(DocxaiError::NotImplemented("add table")),
        AddKind::Image(_) => Err(DocxaiError::NotImplemented("add image")),
        AddKind::Equation(_) => Err(DocxaiError::NotImplemented("add equation")),
    }
}

pub fn run_set(args: cli::SetArgs, writer: &mut dyn Write) -> Result<(), DocxaiError> {
    let parsed = Ref::parse(&args.reference)?;
    match &parsed {
        Ref::Paragraph(_) => {
            if args.text.is_none() && args.style.is_none() {
                return Err(DocxaiError::InvalidArgument(
                    "set @pN requires at least one of --text or --style".into(),
                ));
            }
            if args.width.is_some() || args.caption.is_some() || args.latex.is_some() {
                return Err(DocxaiError::InvalidArgument(
                    "unsupported option for ref kind paragraph".to_string(),
                ));
            }
            let mut doc = Doc::load(&args.file)?;
            let result = mutate::set_paragraph(
                &mut doc,
                &args.reference,
                args.text.as_deref(),
                args.style.as_deref(),
            )?;
            writeln!(writer, "{result}")
                .map_err(|e| DocxaiError::Generic(format!("write: {e}")))?;
            Ok(())
        }
        _ => Err(DocxaiError::NotImplemented("set (this ref kind)")),
    }
}

pub fn run_delete(args: cli::DeleteArgs, writer: &mut dyn Write) -> Result<(), DocxaiError> {
    let parsed = Ref::parse(&args.reference)?;
    match &parsed {
        Ref::Paragraph(_) => {
            let mut doc = Doc::load(&args.file)?;
            let result = mutate::delete_paragraph(&mut doc, &args.reference)?;
            writeln!(writer, "{result}")
                .map_err(|e| DocxaiError::Generic(format!("write: {e}")))?;
            Ok(())
        }
        _ => Err(DocxaiError::NotImplemented("delete (this ref kind)")),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::doc::test_fixture::minimal_docx_bytes;
    use std::io::Write;
    use tempfile::NamedTempFile;

    #[test]
    fn run_styles_prints_json_array() {
        let mut tmp = NamedTempFile::new().unwrap();
        tmp.write_all(&minimal_docx_bytes()).unwrap();
        tmp.flush().unwrap();

        let mut buf = Vec::new();
        run_styles(
            StylesArgs {
                file: tmp.path().to_path_buf(),
            },
            &mut buf,
        )
        .expect("styles should succeed on minimal fixture");
        let s = String::from_utf8(buf).unwrap();
        assert_eq!(s.trim_end(), r#"{"styles":["Title","Body"]}"#);
        assert!(s.ends_with('\n'), "output must be newline-terminated");
    }

    #[test]
    fn run_snapshot_emits_single_line_json_by_default() {
        let mut tmp = NamedTempFile::new().unwrap();
        tmp.write_all(&minimal_docx_bytes()).unwrap();
        tmp.flush().unwrap();

        let mut buf = Vec::new();
        run_snapshot(
            crate::cli::SnapshotArgs {
                file: tmp.path().to_path_buf(),
                pretty: false,
                table: None,
            },
            &mut buf,
        )
        .expect("snapshot should succeed on minimal fixture");
        let s = String::from_utf8(buf).unwrap();
        // Newline-terminated, exactly one line.
        assert!(s.ends_with('\n'));
        assert_eq!(s.matches('\n').count(), 1, "expected single-line JSON");
        // Schema landmarks per PRD §8.1.
        assert!(s.contains(r#""version":"1.0""#));
        assert!(s.contains(r#""available_styles":["Title","Body"]"#));
        assert!(s.contains(r#""ref":"@p1""#));
        assert!(s.contains(r#""kind":"paragraph""#));
        assert!(s.contains(r#""text":"Hello""#));
    }

    #[test]
    fn run_snapshot_pretty_indents_output() {
        let mut tmp = NamedTempFile::new().unwrap();
        tmp.write_all(&minimal_docx_bytes()).unwrap();
        tmp.flush().unwrap();

        let mut buf = Vec::new();
        run_snapshot(
            crate::cli::SnapshotArgs {
                file: tmp.path().to_path_buf(),
                pretty: true,
                table: None,
            },
            &mut buf,
        )
        .unwrap();
        let s = String::from_utf8(buf).unwrap();
        assert!(s.matches('\n').count() > 1, "pretty output must span lines");
    }

    #[test]
    fn run_snapshot_table_subview_returns_error_when_missing() {
        let mut tmp = NamedTempFile::new().unwrap();
        tmp.write_all(&minimal_docx_bytes()).unwrap();
        tmp.flush().unwrap();

        let err = run_snapshot(
            crate::cli::SnapshotArgs {
                file: tmp.path().to_path_buf(),
                pretty: false,
                table: Some("@t1".into()),
            },
            &mut Vec::new(),
        )
        .expect_err("minimal fixture has no tables");
        assert_eq!(err.exit_code(), crate::error::ExitCode::InvalidArgument);
    }

    #[test]
    fn run_styles_propagates_load_error() {
        let err = run_styles(
            StylesArgs {
                file: "/nonexistent/does-not-exist.docx".into(),
            },
            &mut Vec::new(),
        )
        .expect_err("missing file must fail");
        assert_eq!(err.exit_code(), crate::error::ExitCode::Generic);
    }
}
