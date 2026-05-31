//! `docxai skills` — install bundled agent skills (mirrors `glab skills`).
//!
//! The project's agent skills live under `skills/` in the repository and are
//! embedded into the binary at compile time via [`include_str!`]. Installing
//! writes them into a standard `.agents/skills/` layout that Claude Code,
//! Codex, Gemini CLI and other Agent-Skills-compatible agents discover
//! automatically.
//!
//! This is a meta/tooling verb. It does not touch any `.docx`; the five
//! document verbs (snapshot/add/set/delete/styles) stay frozen per PRD §7.1.

use std::fs;
use std::io::Write;
use std::path::{Path, PathBuf};

use crate::cli::{SkillsArgs, SkillsCommand, SkillsInstallArgs};
use crate::error::DocxaiError;

/// One skill embedded in the binary: its name, one-line description, and the
/// `(relative path, contents)` of every file it ships.
struct EmbeddedSkill {
    name: &'static str,
    description: &'static str,
    files: &'static [(&'static str, &'static str)],
}

/// All skills bundled with this build. Paths are relative to `src/`.
static SKILLS: &[EmbeddedSkill] = &[EmbeddedSkill {
    name: "editing-docx-with-docxai",
    description: "Create and modify Word .docx files with the docxai CLI.",
    files: &[
        (
            "SKILL.md",
            include_str!("../skills/editing-docx-with-docxai/SKILL.md"),
        ),
        (
            "references/reference.md",
            include_str!("../skills/editing-docx-with-docxai/references/reference.md"),
        ),
    ],
}];

/// Dispatch `docxai skills <install|list>`.
pub fn run_skills(args: SkillsArgs, writer: &mut dyn Write) -> Result<(), DocxaiError> {
    match args.command {
        SkillsCommand::List => list(writer),
        SkillsCommand::Install(install) => self::install(&install, writer),
    }
}

fn list(writer: &mut dyn Write) -> Result<(), DocxaiError> {
    for skill in SKILLS {
        writeln!(writer, "{}\t{}", skill.name, skill.description)
            .map_err(|e| DocxaiError::Generic(format!("write: {e}")))?;
    }
    Ok(())
}

fn install(args: &SkillsInstallArgs, writer: &mut dyn Write) -> Result<(), DocxaiError> {
    let selected: Vec<&EmbeddedSkill> = match &args.name {
        None => SKILLS.iter().collect(),
        Some(name) => {
            let found: Vec<&EmbeddedSkill> =
                SKILLS.iter().filter(|s| s.name == name.as_str()).collect();
            if found.is_empty() {
                let available: Vec<&str> = SKILLS.iter().map(|s| s.name).collect();
                return Err(DocxaiError::InvalidArgument(format!(
                    "no bundled skill named '{name}'; available: {}",
                    available.join(", ")
                )));
            }
            found
        }
    };

    let base = target_dir(args)?;

    for skill in selected {
        let skill_dir = base.join(skill.name);
        let mut written = 0usize;
        for (rel, contents) in skill.files {
            let dest = skill_dir.join(rel);
            if dest.exists() && !args.force {
                return Err(DocxaiError::InvalidArgument(format!(
                    "{} already exists; pass --force to overwrite",
                    dest.display()
                )));
            }
            if let Some(parent) = dest.parent() {
                fs::create_dir_all(parent).map_err(|e| {
                    DocxaiError::Generic(format!("create {}: {e}", parent.display()))
                })?;
            }
            fs::write(&dest, contents)
                .map_err(|e| DocxaiError::Generic(format!("write {}: {e}", dest.display())))?;
            written += 1;
        }
        writeln!(
            writer,
            "installed skill '{}' ({written} files) -> {}",
            skill.name,
            skill_dir.display()
        )
        .map_err(|e| DocxaiError::Generic(format!("write: {e}")))?;
    }
    Ok(())
}

/// Resolve where skills are installed: `--path` wins, then `--global`
/// (`~/.agents/skills`), else `<repo-root>/.agents/skills`.
fn target_dir(args: &SkillsInstallArgs) -> Result<PathBuf, DocxaiError> {
    if let Some(path) = &args.path {
        return Ok(path.clone());
    }
    if args.global {
        let home = std::env::var_os("HOME").ok_or_else(|| {
            DocxaiError::Generic("cannot resolve --global: HOME is not set".into())
        })?;
        return Ok(Path::new(&home).join(".agents").join("skills"));
    }
    Ok(repo_root().join(".agents").join("skills"))
}

/// Walk up from the current directory looking for a `.git` entry; fall back to
/// the current directory when not inside a repository.
fn repo_root() -> PathBuf {
    let cwd = std::env::current_dir().unwrap_or_else(|_| PathBuf::from("."));
    let mut dir = cwd.as_path();
    loop {
        if dir.join(".git").exists() {
            return dir.to_path_buf();
        }
        match dir.parent() {
            Some(parent) => dir = parent,
            None => return cwd,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::cli::SkillsInstallArgs;

    fn install_args(path: PathBuf, force: bool) -> SkillsInstallArgs {
        SkillsInstallArgs {
            name: None,
            force,
            global: false,
            path: Some(path),
        }
    }

    #[test]
    fn list_includes_bundled_skill() {
        let mut buf = Vec::new();
        list(&mut buf).unwrap();
        let out = String::from_utf8(buf).unwrap();
        assert!(out.contains("editing-docx-with-docxai"));
    }

    #[test]
    fn install_writes_all_files() {
        let dir = tempfile::tempdir().unwrap();
        let target = dir.path().to_path_buf();
        install(&install_args(target.clone(), false), &mut Vec::new()).unwrap();
        assert!(target.join("editing-docx-with-docxai/SKILL.md").exists());
        assert!(
            target
                .join("editing-docx-with-docxai/references/reference.md")
                .exists()
        );
    }

    #[test]
    fn install_refuses_existing_without_force() {
        let dir = tempfile::tempdir().unwrap();
        let target = dir.path().to_path_buf();
        install(&install_args(target.clone(), false), &mut Vec::new()).unwrap();
        let err = install(&install_args(target.clone(), false), &mut Vec::new())
            .expect_err("second install must refuse without --force");
        assert_eq!(err.exit_code(), crate::error::ExitCode::InvalidArgument);
    }

    #[test]
    fn install_force_overwrites() {
        let dir = tempfile::tempdir().unwrap();
        let target = dir.path().to_path_buf();
        install(&install_args(target.clone(), false), &mut Vec::new()).unwrap();
        install(&install_args(target, true), &mut Vec::new())
            .expect("--force must overwrite existing files");
    }

    #[test]
    fn install_unknown_name_errors() {
        let dir = tempfile::tempdir().unwrap();
        let args = SkillsInstallArgs {
            name: Some("does-not-exist".into()),
            force: false,
            global: false,
            path: Some(dir.path().to_path_buf()),
        };
        let err = install(&args, &mut Vec::new()).expect_err("unknown skill must error");
        assert_eq!(err.exit_code(), crate::error::ExitCode::InvalidArgument);
    }
}
