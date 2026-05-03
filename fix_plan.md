# docxai — Implementation Plan

Tracker for autonomous development against `docxai-PRD.md`. Updated each loop.

## Critical path

`#1 → #2 → #3 → #4 → #6 → #7 → #8 → #10 → #12 → #13 → #14`

## Milestone status

### M0 — Bootstrap (v0.0.1)

- [x] **#1** Cargo project + layout — `Cargo.toml`, `src/{main,lib,cli,error}.rs`, `rustfmt.toml`, `.gitignore`
- [x] **#6** clap skeleton with 5 verbs (M1 in PRD but landed early) — see `src/cli.rs`
- [ ] **#2** GitHub Actions CI (`.github/workflows/{ci,release}.yml`)
- [ ] **#3** Test fixtures corpus (`tests/fixtures/*.docx`) — needs real Word/LibreOffice files; blocked on environment
- [ ] **#4** Smoke test: roundtrip noop preservation
- [ ] **#5** `cargo-dist` setup (p1)

### M1 — Read-only (v0.1.0)

- [ ] **#7** `Doc::load()` — open zip, parse parts (in progress this loop)
- [ ] **#8** Ref resolver (`@p3`, `@t1.r2.c3`, …)
- [ ] **#9** `styles` command
- [ ] **#10** `snapshot` command (JSON per §8.1)
- [ ] **#11** Word runs → markdown light renderer

### M2 — Paragraph mutations (v0.2.0)

- [ ] **#12** Markdown subset parser (markdown → `<w:r>`)
- [ ] **#13** `Doc::save()` atomic
- [ ] **#14** `add paragraph` (append)
- [ ] **#15** `add paragraph --after/--before`
- [ ] **#16** `set @pN --text`
- [ ] **#17** `set @pN --style`
- [ ] **#18** `delete @pN`

### M3-M7

See `docxai-PRD.md` §15. Tracked there until M2 complete.

## Per-loop notes

### Loop 2026-05-03 (current)

Picked **#7 Doc::load** — unblocks all reads (snapshot, styles) and writes (save, mutations).

Approach:
- New `src/doc/{mod,load,parts}.rs` module
- `Parts` struct stores well-known XML parts (content_types, document, styles, rels, core_props) plus raw blob list for byte-perfect preservation of unknown parts
- Tests build minimal valid docx in tempfile via `zip` crate (no real fixture binaries required yet); fixtures (#3) deferred until #14 lands and we need round-trip cases
- Errors mapped to `DocxaiError::Generic` (PRD §10.1: file inaccessible / format invalid → exit 1)

Out-of-scope this loop: actual XML parsing of content (deferred to #8/#10), atomic save (#13), CI setup (#2).

**Status:** code written, tests NOT_RUN. `cargo` commands are not in `.ralphrc` `ALLOWED_TOOLS`, every loop's first action must request approval (or, better, add `Bash(cargo *)` and `Bash(cargo)` to that list).

## Open decisions

From PRD §15 "Décisions à acter avant kickoff" — defer:
1. License: already MIT OR Apache-2.0 in `Cargo.toml`. ✓
2. GitHub org / release owners / eval framework / equation priority: not needed pre-M2.
