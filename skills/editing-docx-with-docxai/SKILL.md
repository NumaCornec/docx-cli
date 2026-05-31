---
name: editing-docx-with-docxai
description: Create and modify Microsoft Word .docx files from the command line using the docxai CLI. Use whenever a task involves reading, inspecting, generating, or editing a .docx (Word) document — adding or changing paragraphs, headings, tables, images, or equations, applying named styles, or extracting document content as structured JSON. Do not hand-edit the .docx zip/XML directly; drive docxai instead.
---

# Editing .docx files with docxai

`docxai` is a deterministic CLI for editing Word `.docx` files. It edits in place
and preserves everything it does not touch (tracked changes, comments, footnotes,
headers/footers, formatting). Never unzip a `.docx` and patch the XML by hand —
that loses preservation guarantees. Always go through `docxai`.

There are exactly **five verbs**: `snapshot`, `styles`, `add`, `set`, `delete`.
No other verbs exist and the surface is frozen.

## The core loop

Edits target elements by **ref** (`@p3`, `@t1`, …). Refs come from `snapshot`,
so always snapshot before and after an edit:

```
1. snapshot  → learn current refs + available styles
2. add/set/delete  → make ONE change using a ref from step 1
3. snapshot  → confirm; refs may have shifted, so re-read before the next edit
```

Refs are **1-indexed positions**, not stable IDs. Inserting or deleting an
element renumbers everything after it. Never reuse a ref across an edit — always
re-snapshot to get fresh refs.

## Refs

| Ref           | Points to                                  |
|---------------|--------------------------------------------|
| `@p3`         | 3rd paragraph                              |
| `@t1`         | 1st table                                  |
| `@t1.r2.c3`   | row 2, col 3 of table 1 (cells 1-indexed)  |
| `@i1`         | 1st inline image                           |
| `@e1`         | 1st equation                               |

Paragraphs, tables, images, equations each have their **own** counter (`@p`, `@t`,
`@i`, `@e`), numbered by order of appearance within their kind.

## Verbs

```sh
# Inspect — JSON to stdout, document untouched
docxai snapshot report.docx --pretty
docxai snapshot report.docx --table @t1     # drill: all cells of one table
docxai styles report.docx                   # named styles usable with --style

# Add (append by default; --after @ref / --before @ref to position)
docxai add report.docx paragraph --text "Findings" --style Heading1
docxai add report.docx paragraph --text "Body text." --after @p4
docxai add report.docx table --rows 3 --cols 2 --header "Metric,Value"
docxai add report.docx image --path chart.png --width 12cm --caption "Fig 1"
docxai add report.docx equation --latex "x^2 + y^2 = z^2"

# Set (edit existing element by ref)
docxai set report.docx @p3 --text "New text" --style Heading2
docxai set report.docx @t1.r2.c1 --text "42"     # one table cell
docxai set report.docx @i1 --width 8cm --caption "Updated"
docxai set report.docx @e1 --latex "a^2 = b^2 + c^2"

# Delete
docxai delete report.docx @p7
```

`--after` and `--before` are mutually exclusive. Omit both to append at end of body.

## Key rules

- **Styles must exist.** `--style` only accepts a name from `docxai styles`
  (the `available_styles` list). Inventing a style name fails. If no `--style` is
  given, `add paragraph` uses the document's `Body` style when present, else leaves
  the paragraph unstyled to inherit the document default.
- **Text is a markdown subset**, not full markdown: `**bold**`, `*italic*`,
  `***both***`, backslash-escape for literals, inline math `$...$`. Anything else
  (headings via `#`, lists, links, `_underscore_`) is rejected. Use `--style` for
  headings, not `#`.
- **Equations need `pandoc`** on PATH (LaTeX ↔ OOXML math). Only `add equation`,
  `set --latex`, and equation round-tripping in `snapshot` require it; every other
  verb works without pandoc.
- **One change per command.** Compose edits as a sequence of calls, re-snapshotting
  between them rather than batching.

## Exit codes

Check the exit code, not just stdout. `0` success; `2` invalid argument (bad ref,
unknown style, malformed markdown); `3` preservation impossible; `4` missing
dependency (e.g. pandoc absent); `1` other failure. Error text goes to stderr as
`error: ...`.

## Snapshot JSON

`snapshot` prints one JSON object: `file`, `version`, `metadata`,
`available_styles`, `body` (ordered array of items, each with a `kind` and `ref`),
and `preserved_features`. For the full field-by-field schema, the markdown subset
grammar, and the ref grammar, read `references/reference.md`.
