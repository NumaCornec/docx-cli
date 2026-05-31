# docxai reference

Detailed schema and grammar for `docxai`. Read alongside `SKILL.md`.

## Contents

- Snapshot JSON schema
- `--table @tN` drill schema
- Ref grammar
- Markdown subset grammar
- Add/set argument detail

## Snapshot JSON schema

`docxai snapshot <file>` emits one JSON object (compact; add `--pretty` for
indented). Top level:

| Field                | Type            | Notes                                             |
|----------------------|-----------------|---------------------------------------------------|
| `file`               | string          | Path that was snapshotted.                        |
| `version`            | string          | Snapshot schema version.                          |
| `metadata`           | object          | `title`, `author` (each may be absent/null).      |
| `available_styles`   | array<string>   | Named styles usable with `--style`.               |
| `body`               | array<item>     | Ordered body elements (see below).                |
| `preserved_features` | array<string>   | Detected features kept verbatim, e.g. `footnotes`, `comments`, `tracked-changes`, `equations`. |

Each `body` item is tagged by `kind`:

**paragraph**
```json
{ "kind": "paragraph", "ref": "@p3", "style": "Heading1", "text": "Findings" }
```
- `style` omitted when the paragraph is unstyled.
- `text` is rendered in the markdown subset (`**bold**`, `*italic*`, `$math$`).

**table**
```json
{ "kind": "table", "ref": "@t1", "rows": 3, "cols": 2, "header": ["Metric", "Value"] }
```
- `header` omitted when there is no header row. The `body` view does **not**
  include cell contents — use `snapshot --table @t1` to get cells.

**image**
```json
{ "kind": "image", "ref": "@i1", "src": "media/image1.png", "width": "12cm", "caption_ref": "@p9" }
```
- `src`, `width`, `caption_ref` each omitted when unknown. `caption_ref` points
  to the paragraph holding the caption.

**equation**
```json
{ "kind": "equation", "ref": "@e1", "latex": "x^2 + y^2 = z^2", "display": true }
```
- `latex` present only when pandoc round-trip succeeded. `display` true =
  block equation, false = inline.

## `--table @tN` drill schema

`docxai snapshot <file> --table @t1` returns one table with cell text:

```json
{
  "ref": "@t1",
  "rows": 2,
  "cols": 2,
  "cells": [
    [ { "ref": "@t1.r1.c1", "text": "Metric" }, { "ref": "@t1.r1.c2", "text": "Value" } ],
    [ { "ref": "@t1.r2.c1", "text": "Revenue" }, { "ref": "@t1.r2.c2", "text": "42" } ]
  ]
}
```

`cells` is row-major; `cells[r-1][c-1]` is `@t1.r{r}.c{c}`.

## Ref grammar

1-indexed, leading `@`. One counter per kind.

| Form          | Meaning                              |
|---------------|--------------------------------------|
| `@p<n>`       | nth paragraph                        |
| `@t<n>`       | nth table                            |
| `@t<n>.r<r>.c<c>` | cell at row r, col c of table n  |
| `@i<n>`       | nth image                            |
| `@e<n>`       | nth equation                         |

Malformed refs (missing `@`, unknown letter, zero/negative index, partial cell
like `@t1.r2`) fail with exit code 2. Indices are always ≥ 1.

## Markdown subset grammar

Accepted inline syntax for `--text` and snapshot `text` output:

| Syntax        | Result                                  |
|---------------|-----------------------------------------|
| `**bold**`    | bold run                                |
| `*italic*`    | italic run                              |
| `***both***`  | bold + italic run                       |
| `\*`, `\\`    | backslash escape → literal meta-char    |
| `$...$`       | inline math (LaTeX → OOXML, needs pandoc)|

Rendering and parsing are inverse over this subset. **Rejected** (exit 2):
`_underscore_` emphasis (reserved), headings (`#`), lists, links, images,
blockquotes, code spans/blocks, and any other block markdown. For a heading,
apply a heading `--style`, not `#`.

## Add / set argument detail

`add <file> <kind>` subcommands and their flags:

| Kind        | Required           | Optional                                  |
|-------------|--------------------|-------------------------------------------|
| `paragraph` | `--text`           | `--style`, `--after`/`--before`           |
| `table`     | `--rows`, `--cols` | `--header "a,b,c"`, `--after`/`--before`  |
| `image`     | `--path`           | `--width`, `--caption`, `--after`/`--before` |
| `equation`  | `--latex`          | `--after`/`--before`                      |

`--after` / `--before` take a ref and are mutually exclusive; omit both to
append. `--width` accepts a unit suffix: `12cm`, `4.5in`, `300px`.

`set <file> <ref>` applies whichever of these apply to the target kind:
`--text`, `--style` (paragraphs), `--text` (table cells), `--width`,
`--caption` (images), `--latex` (equations). Flags that don't match the
referenced element's kind are rejected.

`--header` cell count should match `--cols`; the header becomes table row 1.
