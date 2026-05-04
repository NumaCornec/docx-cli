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

- [x] **#7** `Doc::load()` — open zip, parse parts (`src/doc.rs`)
- [ ] **#8** Ref resolver (`@p3`, `@t1.r2.c3`, …) — **parser landed (`src/refs.rs`); body indexer + `Doc::resolve` deferred until #10 needs it**
- [x] **#9** `styles` command — `src/styles.rs` parser + `run_styles` wired in `lib.rs`
- [x] **#10** `snapshot` command (JSON per §8.1) — `src/snapshot.rs` body indexer + `run_snapshot` wired in `lib.rs`; v0.1 covers paragraphs (with bold/italic markdown), tables (placeholder + drill via `--table @tN`), `metadata` from core.xml, `available_styles`, `preserved_features` for footnotes/comments/tracked-changes/equations
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

### Loop 2026-05-05 b (current)

Picked **#10 snapshot** — critical-path unblocker for the agent loop and consumer of #7 (Doc::load), #9 (styles), and the body indexer the resolver in #8 will reuse.

Approach:
- New `src/snapshot.rs`: streaming `quick-xml` walk over `word/document.xml`. Tracks `tbl_depth` so paragraphs nested in cells do **not** consume top-level `@pN` slots — `ref_numbering_is_sequential_and_per_kind` test pins this invariant.
- `BodyItem` is `#[serde(tag = "kind")]` (`paragraph` | `table`) matching PRD §8.1. Images (#24) and equations (#30) deferred — they appear in `preserved_features` instead so the agent knows the doc has them.
- Paragraph runs concatenate `<w:t>` text; `<w:b/>` → `**…**`, `<w:i/>` → `*…*`. Markdown-special chars get backslash-escaped at render time so subsequent `set --text` round-trips don't double-interpret. Full markdown subset (code, links, math) lives with #11.
- Tables: capture rows/cols + header (first row); cell texts cached in `TableLayout` so `--table @tN` reuses the same walk via `build_table_snapshot`.
- `--table` ref validation goes through the existing `Ref::parse`; non-table ref / out-of-bounds index return `InvalidArgument` (exit 2).
- Output: single-line JSON by default (PRD §10.2), `--pretty` switches to indented. Both newline-terminated.
- `preserved_features` is a sorted Vec built from cheap byte scans (`word/footnotes.xml`, `comments.xml`, `endnotes.xml` in the others map; `<w:ins`/`<w:del`/`<m:oMath` substring sweep on document.xml). Good enough for v0.1; precise detection lands with #19/#24/#30.

Out-of-scope this loop: image/equation body items, `Doc::resolve` (still deferred to when #16 needs it), `insta` snapshots over a real fixture corpus (#3 still blocked).

**Status:** 50/50 lib tests pass via `cargo test --lib`. Cargo found at `/home/ncornec/.cargo/bin/cargo`; whitelist is no longer the blocker prior loops described.

### Loop 2026-05-05

Picked **#9 styles command** — small, self-contained, and feeds `available_styles` into #10. Critical-path unblock.

Approach:
- New `src/styles.rs`: pure XML walker over `word/styles.xml` via `quick-xml` event reader. Emits `w:styleId` of every `<w:style w:type="paragraph">` in document order, dropping `<w:semiHidden/>` ones (PRD #9: "ne garder que ceux utilisables").
- Wired `Command::Styles` to a `run_styles(args, &mut writer)` in `lib.rs` so tests can capture stdout without spawning a process. Output is single-line JSON `{"styles":[...]}` newline-terminated, conforming to PRD §10.2.
- Errors surface as `DocxaiError::Generic` (load failures, malformed styles.xml) → exit code 1. No new error variant needed.

Out-of-scope this loop: snapshot tests via `insta` against the fixture corpus (#3 still blocked); body indexer for #10; markdown rendering #11.

**Status:** code written, tests NOT_RUN — `cargo` still not on `$PATH` in this env. Same blocker as prior two loops; mitigation is install rust toolchain or whitelist `Bash(cargo *)`.

### Loop 2026-05-04

Picked **#8 ref parser** — pure string parsing, self-contained, unblocks `set @pN` / `delete @pN` arg validation and the body indexer in #10.

Approach:
- New `src/refs.rs` with `Ref` enum (`Paragraph | Table | TableCell | Image | Equation`) and `Ref::parse(&str) -> Result<Self, DocxaiError>`.
- Strict grammar: leading `@`, single ASCII sigil (`p`/`t`/`i`/`e`), 1-indexed integer with no leading zero / sign / whitespace, table-cell shape `@tN.rR.cC` only.
- Errors are `DocxaiError::InvalidArgument` so they map to PRD §10.1 exit code 2 (`ref @p99 not found`-style messages will live in the resolver, not the parser).
- Tests: 100-ref valid sweep, 51-ref invalid sweep, plus targeted cases (zero, leading zero, overflow, malformed cell, non-ASCII digits, whitespace, signed). `Display` round-trip covers serialisation.

Out-of-scope this loop: body indexer (`Vec<BodyRef>` order from `document.xml`), `Doc::resolve(&Ref)`. Both need quick-xml walking and slot in naturally with #10.

**Status:** code written, tests NOT_RUN — `cargo` is still not on `$PATH` in this env (not just blocked by `.ralphrc`). Need either rust toolchain installed in the loop image or a local `Bash(/path/to/cargo *)` allowance before any compile/test step can pass.

### Loop 2026-05-03

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
