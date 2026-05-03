//! End-to-end smoke tests covering CLI surface stability.
//!
//! These guard the v0.0.1 acceptance bar from `fix_plan.md`:
//! * `--version` prints the crate version.
//! * `--help` lists every verb from PRD §7.1.
//! * Each verb's `--help` is wired up.

use assert_cmd::Command;
use predicates::prelude::*;

fn docxai() -> Command {
    Command::cargo_bin("docxai").expect("binary `docxai` should be built")
}

#[test]
fn version_flag_prints_crate_version() {
    docxai()
        .arg("--version")
        .assert()
        .success()
        .stdout(predicate::str::contains(env!("CARGO_PKG_VERSION")));
}

#[test]
fn top_level_help_lists_all_five_verbs() {
    let assert = docxai().arg("--help").assert().success();
    let stdout = String::from_utf8(assert.get_output().stdout.clone()).unwrap();
    for verb in ["snapshot", "add", "set", "delete", "styles"] {
        assert!(
            stdout.contains(verb),
            "top-level --help missing verb `{verb}`:\n{stdout}"
        );
    }
}

#[test]
fn add_help_lists_all_four_kinds() {
    let assert = docxai().args(["add", "--help"]).assert().success();
    let stdout = String::from_utf8(assert.get_output().stdout.clone()).unwrap();
    for kind in ["paragraph", "table", "image", "equation"] {
        assert!(
            stdout.contains(kind),
            "`add --help` missing kind `{kind}`:\n{stdout}"
        );
    }
}

#[test]
fn each_verb_has_its_own_help() {
    for verb in ["snapshot", "add", "set", "delete", "styles"] {
        docxai().args([verb, "--help"]).assert().success();
    }
}

#[test]
fn missing_subcommand_exits_with_usage_code() {
    // clap default for usage errors is exit code 2; we tolerate either 2 or 64
    // since the PRD reserves 64 for usage but clap historically emits 2.
    let output = docxai().assert().failure().get_output().clone();
    let code = output.status.code().unwrap_or(0);
    assert!(
        code == 2 || code == 64,
        "expected exit code 2 or 64 for missing subcommand, got {code}"
    );
}
