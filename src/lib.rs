//! docxai library entry point.
//!
//! Public modules expose the CLI definitions and error model; the binary in
//! `src/main.rs` only handles process boundaries (argv parsing, exit codes).

pub mod cli;
pub mod error;

use crate::cli::{Cli, Command};
use crate::error::DocxaiError;

/// Dispatch a parsed [`Cli`] to its handler.
///
/// Each subcommand is currently a stub returning [`DocxaiError::NotImplemented`]
/// so the binary remains buildable while individual verbs are filled in.
pub fn run(cli: Cli) -> Result<(), DocxaiError> {
    match cli.command {
        Command::Snapshot(_) => Err(DocxaiError::NotImplemented("snapshot")),
        Command::Add(_) => Err(DocxaiError::NotImplemented("add")),
        Command::Set(_) => Err(DocxaiError::NotImplemented("set")),
        Command::Delete(_) => Err(DocxaiError::NotImplemented("delete")),
        Command::Styles(_) => Err(DocxaiError::NotImplemented("styles")),
    }
}
