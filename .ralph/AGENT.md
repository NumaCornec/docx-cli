# Ralph Agent Configuration

## Build Instructions

```bash
cargo build
```

For an optimized binary:

```bash
cargo build --release
```

## Test Instructions

```bash
cargo test
```

Linters (must pass before commit):

```bash
cargo fmt --check
cargo clippy --all-targets -- -D warnings
```

## Run Instructions

```bash
cargo run -- --help
cargo run -- snapshot path/to/file.docx
```

After `cargo build --release`, the binary lives at `target/release/docxai`.

## Notes
- MSRV: Rust 1.85 (edition 2024). The PRD targets 1.95+ / edition 2026; we
  track today's stable until those land.
- Five verbs only: `snapshot`, `add`, `set`, `delete`, `styles`. Adding a
  sixth verb requires explicit approval (PRD §7.1).
- Exit codes follow PRD §10.1: 0 success, 1 generic, 2 invalid argument,
  3 preservation impossible, 4 missing dependency, 64 usage.
