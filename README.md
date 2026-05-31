# docxai

A deterministic Rust CLI for AI agents to create and modify `.docx` (Word) files.

`docxai` exposes exactly five verbs — `snapshot`, `add`, `set`, `delete`, `styles` —
designed to be discovered and used by coding agents (Claude Code, Cursor, Codex CLI).
There are no LLM calls inside the binary; the agent drives `docxai` over Bash.

## Status

Alpha. All five verbs are implemented and tested against synthetic fixtures:

- `snapshot` — JSON body/styles/refs, `--pretty`, `--table @tN` drill-down
- `styles` — list usable paragraph styles
- `add` — `paragraph`, `table`, `image`, `equation` (with `--after`/`--before`)
- `set` — edit paragraph text/style and table cells
- `delete` — remove an element by ref

Not yet verified against real Word/LibreOffice `.docx` output (synthetic
fixtures only). See `docxai-PRD.md` for the full spec and `fix_plan.md` for
milestone status.

## Requirements

- Rust 1.85+ (edition 2024).
- **`pandoc`** — required only for `add equation` and equation round-tripping
  in `snapshot` (LaTeX ↔ OOXML math). All other verbs work without it.
  - macOS: `brew install pandoc` · Linux: `sudo apt install pandoc` ·
    Windows: `winget install pandoc`

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
docxai add report.docx equation \
    --latex "x^2 + y^2 = z^2"            # display equation (needs pandoc)
```

If no `--style` is given, `add paragraph` uses the document's `Body` style when
present, otherwise leaves the paragraph unstyled to inherit the document default.

Run `docxai --help` and `docxai <verb> --help` for the full reference.

## Build & test

```sh
cargo build
cargo test
cargo clippy -- -D warnings
cargo fmt --check
```

## License

Apache-2.0. See [LICENSE](LICENSE).
