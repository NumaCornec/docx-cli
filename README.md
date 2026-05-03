# docxai

A deterministic Rust CLI for AI agents to create and modify `.docx` (Word) files.

`docxai` exposes exactly five verbs — `snapshot`, `add`, `set`, `delete`, `styles` —
designed to be discovered and used by coding agents (Claude Code, Cursor, Codex CLI).
There are no LLM calls inside the binary; the agent drives `docxai` over Bash.

## Status

Pre-alpha. The CLI surface is stubbed; verbs return a "not yet implemented"
error while individual milestones land. See `docxai-PRD.md` for the full spec.

## Install

```sh
cargo install --path .
```

## Quick start

```sh
docxai snapshot report.docx              # inspect refs and styles
docxai add report.docx paragraph \
    --text "Findings" --style Heading1   # append a paragraph
docxai set report.docx @p3 --text "..."  # edit an existing paragraph
docxai delete report.docx @p7            # remove an element
docxai styles report.docx                # list available styles
```

Run `docxai --help` and `docxai <verb> --help` for the full reference.

## Build & test

```sh
cargo build
cargo test
cargo clippy -- -D warnings
cargo fmt --check
```

## License

MIT OR Apache-2.0.
