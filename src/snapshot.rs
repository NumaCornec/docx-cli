//! Snapshot builder (PRD §8 / #10).
//!
//! Walks `word/document.xml` and emits the JSON shape described in §8.1.
//! v0.1 scope per PRD #10:
//!   * paragraphs render to light markdown (bold/italic only for now; full
//!     subset lands with #11)
//!   * tables appear as placeholders with `rows`/`cols`/`header`
//!   * images and equations are deferred to #24 / #30
//!
//! The body indexer is also the basis for `Doc::resolve` (#8 followup).

use std::collections::BTreeMap;

use quick_xml::Reader;
use quick_xml::events::Event;
use serde::Serialize;

use crate::doc::Doc;
use crate::error::DocxaiError;
use crate::refs::Ref;
use crate::styles;

/// Top-level snapshot payload (§8.1).
#[derive(Debug, Serialize)]
pub struct Snapshot {
    pub file: String,
    pub version: &'static str,
    pub metadata: Metadata,
    pub available_styles: Vec<String>,
    pub body: Vec<BodyItem>,
    pub preserved_features: Vec<String>,
}

#[derive(Debug, Default, Serialize)]
pub struct Metadata {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub title: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub author: Option<String>,
}

/// One ordered body element. The variant tag is serialised as `kind`.
#[derive(Debug, Serialize)]
#[serde(tag = "kind", rename_all = "lowercase")]
pub enum BodyItem {
    Paragraph {
        #[serde(rename = "ref")]
        reference: String,
        #[serde(skip_serializing_if = "Option::is_none")]
        style: Option<String>,
        text: String,
    },
    Table {
        #[serde(rename = "ref")]
        reference: String,
        rows: u32,
        cols: u32,
        #[serde(skip_serializing_if = "Vec::is_empty")]
        header: Vec<String>,
    },
}

impl BodyItem {
    /// The ref string (`@p3`, `@t1`) — useful for `--table` lookup.
    pub fn reference(&self) -> &str {
        match self {
            BodyItem::Paragraph { reference, .. } | BodyItem::Table { reference, .. } => reference,
        }
    }
}

/// Drilled view of a single table (§8.3 — `snapshot --table @tN`).
#[derive(Debug, Serialize)]
pub struct TableSnapshot {
    #[serde(rename = "ref")]
    pub reference: String,
    pub rows: u32,
    pub cols: u32,
    pub cells: Vec<Vec<TableCell>>,
}

#[derive(Debug, Serialize)]
pub struct TableCell {
    #[serde(rename = "ref")]
    pub reference: String,
    pub text: String,
}

/// Cached layout from a single body walk; lets `--table @tN` reuse the
/// indexer without re-parsing.
pub struct BodyIndex {
    pub items: Vec<BodyItem>,
    pub tables: Vec<TableLayout>,
}

pub struct TableLayout {
    pub reference: String,
    pub cells: Vec<Vec<TableCell>>,
    pub rows: u32,
    pub cols: u32,
}

pub fn build_snapshot(doc: &Doc) -> Result<Snapshot, DocxaiError> {
    let index = index_body(&doc.parts.document_xml)?;
    let metadata = match doc.parts.core_props.as_deref() {
        Some(xml) => parse_core_metadata(xml)?,
        None => Metadata::default(),
    };
    let file_name = doc
        .path
        .file_name()
        .map(|s| s.to_string_lossy().into_owned())
        .unwrap_or_else(|| doc.path.to_string_lossy().into_owned());

    Ok(Snapshot {
        file: file_name,
        version: "1.0",
        metadata,
        available_styles: styles::list_paragraph_styles(doc)?,
        body: index.items,
        preserved_features: detect_preserved_features(doc),
    })
}

/// Build the drilled view of `--table @tN`. Errors if the ref is malformed,
/// is not a table ref, or names a table that does not exist.
pub fn build_table_snapshot(doc: &Doc, table_ref: &str) -> Result<TableSnapshot, DocxaiError> {
    let parsed = Ref::parse(table_ref)?;
    let n = match parsed {
        Ref::Table(n) => n,
        _ => {
            return Err(DocxaiError::InvalidArgument(format!(
                "--table expects a table ref like @t1, got {table_ref:?}"
            )));
        }
    };
    let index = index_body(&doc.parts.document_xml)?;
    let table = index
        .tables
        .into_iter()
        .nth((n - 1) as usize)
        .ok_or_else(|| {
            DocxaiError::InvalidArgument(format!("table {table_ref} not found in document"))
        })?;
    Ok(TableSnapshot {
        reference: table.reference,
        rows: table.rows,
        cols: table.cols,
        cells: table.cells,
    })
}

/// Walk `document.xml` and assemble ordered body items. We track the
/// nesting depth ourselves so paragraphs *inside* a table cell are not
/// counted as top-level paragraphs.
fn index_body(xml: &[u8]) -> Result<BodyIndex, DocxaiError> {
    let mut reader = Reader::from_reader(xml);
    let mut buf = Vec::new();

    let mut items = Vec::new();
    let mut tables: Vec<TableLayout> = Vec::new();
    let mut p_count: u32 = 0;
    let mut t_count: u32 = 0;

    let mut in_body = false;
    let mut tbl_depth: u32 = 0;

    // Per-table state; valid only while tbl_depth > 0.
    let mut current_table: Option<TableInProgress> = None;

    loop {
        let event = reader.read_event_into(&mut buf).map_err(|e| {
            DocxaiError::Generic(format!(
                "document.xml parse error at {}: {e}",
                reader.buffer_position()
            ))
        })?;

        match event {
            Event::Eof => break,
            Event::Start(ref e) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"body" {
                    in_body = true;
                } else if !in_body {
                    // Skip until <w:body>.
                } else if local == b"tbl" {
                    if tbl_depth == 0 {
                        // New top-level table.
                        t_count += 1;
                        current_table = Some(TableInProgress::new(t_count));
                    }
                    tbl_depth += 1;
                } else if tbl_depth > 0 {
                    // Inside a table — collect rows/cells.
                    if let Some(tbl) = current_table.as_mut() {
                        match local {
                            b"tr" if tbl_depth == 1 => tbl.start_row(),
                            b"tc" if tbl_depth == 1 => tbl.start_cell(),
                            _ => {}
                        }
                    }
                    if local == b"p" && tbl_depth == 1 {
                        // Paragraph inside a table cell — capture its text into the current cell.
                        let para = read_paragraph(&mut reader, &mut buf)?;
                        if let Some(tbl) = current_table.as_mut() {
                            tbl.append_cell_text(&para.text);
                        }
                        continue;
                    }
                } else if local == b"p" {
                    // Top-level paragraph.
                    let para = read_paragraph(&mut reader, &mut buf)?;
                    p_count += 1;
                    items.push(BodyItem::Paragraph {
                        reference: format!("@p{p_count}"),
                        style: para.style,
                        text: para.text,
                    });
                    continue;
                }
            }
            Event::Empty(ref e) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if in_body && tbl_depth == 0 && local == b"p" {
                    // Empty <w:p/> — vanishingly rare but valid.
                    p_count += 1;
                    items.push(BodyItem::Paragraph {
                        reference: format!("@p{p_count}"),
                        style: None,
                        text: String::new(),
                    });
                }
            }
            Event::End(ref e) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                if local == b"body" {
                    in_body = false;
                } else if local == b"tbl" {
                    tbl_depth = tbl_depth.saturating_sub(1);
                    if tbl_depth == 0 {
                        if let Some(tbl) = current_table.take() {
                            let (item, layout) = tbl.finish();
                            items.push(item);
                            tables.push(layout);
                        }
                    }
                } else if tbl_depth > 0 {
                    if let Some(tbl) = current_table.as_mut() {
                        match local {
                            b"tr" if tbl_depth == 1 => tbl.end_row(),
                            _ => {}
                        }
                    }
                }
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(BodyIndex { items, tables })
}

struct ParagraphData {
    style: Option<String>,
    text: String,
}

/// Read a paragraph's interior, consuming events up to (and including)
/// the matching `</w:p>`. Extracts pStyle and concatenates run text with
/// minimal markdown for bold/italic.
fn read_paragraph(
    reader: &mut Reader<&[u8]>,
    buf: &mut Vec<u8>,
) -> Result<ParagraphData, DocxaiError> {
    let mut style: Option<String> = None;
    let mut text = String::new();

    let mut in_run = false;
    let mut bold = false;
    let mut italic = false;
    let mut in_text = false;
    let mut run_text = String::new();

    loop {
        let event = reader.read_event_into(buf).map_err(|e| {
            DocxaiError::Generic(format!(
                "document.xml parse error at {}: {e}",
                reader.buffer_position()
            ))
        })?;
        match event {
            Event::Eof => break,
            Event::Start(ref e) => {
                let local = local_name(e.name().as_ref()).to_owned();
                match local.as_slice() {
                    b"r" => {
                        in_run = true;
                        bold = false;
                        italic = false;
                        run_text.clear();
                    }
                    b"t" if in_run => in_text = true,
                    _ => {}
                }
            }
            Event::Empty(ref e) => {
                let local = local_name(e.name().as_ref()).to_owned();
                match local.as_slice() {
                    b"pStyle" => {
                        for attr in e.attributes().flatten() {
                            if local_name(attr.key.as_ref()) == b"val" {
                                let val = attr
                                    .decode_and_unescape_value(reader.decoder())
                                    .map_err(|err| {
                                        DocxaiError::Generic(format!("pStyle attr decode: {err}"))
                                    })?
                                    .into_owned();
                                style = Some(val);
                            }
                        }
                    }
                    b"b" if in_run => bold = true,
                    b"i" if in_run => italic = true,
                    b"tab" if in_run => run_text.push('\t'),
                    b"br" if in_run => run_text.push(' '),
                    _ => {}
                }
            }
            Event::Text(ref t) if in_text => {
                let s = t
                    .unescape()
                    .map_err(|e| DocxaiError::Generic(format!("text decode: {e}")))?;
                run_text.push_str(&s);
            }
            Event::End(ref e) => {
                let local = local_name(e.name().as_ref()).to_owned();
                match local.as_slice() {
                    b"t" => in_text = false,
                    b"r" => {
                        if !run_text.is_empty() {
                            let escaped = escape_markdown(&run_text);
                            let wrapped = match (bold, italic) {
                                (true, true) => format!("***{escaped}***"),
                                (true, false) => format!("**{escaped}**"),
                                (false, true) => format!("*{escaped}*"),
                                (false, false) => escaped,
                            };
                            text.push_str(&wrapped);
                        }
                        in_run = false;
                    }
                    b"p" => break,
                    _ => {}
                }
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(ParagraphData { style, text })
}

/// Backslash-escape characters that have meaning in our markdown subset.
fn escape_markdown(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for ch in s.chars() {
        match ch {
            '\\' | '*' | '_' | '`' | '$' | '[' | ']' => {
                out.push('\\');
                out.push(ch);
            }
            _ => out.push(ch),
        }
    }
    out
}

struct TableInProgress {
    n: u32,
    rows: Vec<Vec<String>>,
    current_row: Option<Vec<String>>,
}

impl TableInProgress {
    fn new(n: u32) -> Self {
        Self {
            n,
            rows: Vec::new(),
            current_row: None,
        }
    }
    fn start_row(&mut self) {
        self.current_row = Some(Vec::new());
    }
    fn end_row(&mut self) {
        if let Some(row) = self.current_row.take() {
            self.rows.push(row);
        }
    }
    fn start_cell(&mut self) {
        if let Some(row) = self.current_row.as_mut() {
            row.push(String::new());
        }
    }
    fn append_cell_text(&mut self, txt: &str) {
        if let Some(row) = self.current_row.as_mut() {
            if let Some(last) = row.last_mut() {
                if !last.is_empty() && !txt.is_empty() {
                    last.push('\n');
                }
                last.push_str(txt);
            }
        }
    }
    fn finish(self) -> (BodyItem, TableLayout) {
        let rows_n = self.rows.len() as u32;
        let cols_n = self.rows.iter().map(Vec::len).max().unwrap_or(0) as u32;
        let header = self.rows.first().cloned().unwrap_or_default();
        let reference = format!("@t{}", self.n);
        let cells = self
            .rows
            .iter()
            .enumerate()
            .map(|(ri, row)| {
                row.iter()
                    .enumerate()
                    .map(|(ci, text)| TableCell {
                        reference: format!("@t{}.r{}.c{}", self.n, ri + 1, ci + 1),
                        text: text.clone(),
                    })
                    .collect()
            })
            .collect();

        let item = BodyItem::Table {
            reference: reference.clone(),
            rows: rows_n,
            cols: cols_n,
            header,
        };
        let layout = TableLayout {
            reference,
            cells,
            rows: rows_n,
            cols: cols_n,
        };
        (item, layout)
    }
}

/// Best-effort core.xml parse for `dc:title` and `dc:creator`.
fn parse_core_metadata(xml: &[u8]) -> Result<Metadata, DocxaiError> {
    let mut reader = Reader::from_reader(xml);
    let mut buf = Vec::new();
    let mut meta = Metadata::default();
    let mut current: Option<&'static str> = None;

    loop {
        let event = reader.read_event_into(&mut buf).map_err(|e| {
            DocxaiError::Generic(format!(
                "core.xml parse error at {}: {e}",
                reader.buffer_position()
            ))
        })?;
        match event {
            Event::Eof => break,
            Event::Start(ref e) => {
                let name = e.name();
                let local = local_name(name.as_ref());
                current = match local {
                    b"title" => Some("title"),
                    b"creator" => Some("creator"),
                    _ => None,
                };
            }
            Event::Text(ref t) => {
                if let Some(field) = current {
                    let s = t
                        .unescape()
                        .map_err(|e| DocxaiError::Generic(format!("core.xml text: {e}")))?
                        .into_owned();
                    match field {
                        "title" => meta.title = Some(s),
                        "creator" => meta.author = Some(s),
                        _ => {}
                    }
                }
            }
            Event::End(_) => current = None,
            _ => {}
        }
        buf.clear();
    }
    Ok(meta)
}

/// Detect OOXML features we round-trip but cannot edit yet (PRD §8.2).
fn detect_preserved_features(doc: &Doc) -> Vec<String> {
    let mut found: BTreeMap<&'static str, ()> = BTreeMap::new();
    if doc.parts.others.contains_key("word/footnotes.xml") {
        found.insert("footnotes", ());
    }
    if doc.parts.others.contains_key("word/endnotes.xml") {
        found.insert("endnotes", ());
    }
    if doc.parts.others.contains_key("word/comments.xml") {
        found.insert("comments", ());
    }
    // Cheap byte-substring scan; document.xml is the source of truth for
    // tracked changes / OMML even when extra parts are absent.
    let doc_xml = doc.parts.document_xml.as_slice();
    if contains_subseq(doc_xml, b"<w:ins ")
        || contains_subseq(doc_xml, b"<w:del ")
        || contains_subseq(doc_xml, b"<w:moveFrom ")
        || contains_subseq(doc_xml, b"<w:moveTo ")
    {
        found.insert("tracked_changes", ());
    }
    if contains_subseq(doc_xml, b"<m:oMath") {
        found.insert("equations", ());
    }
    found.into_keys().map(String::from).collect()
}

fn contains_subseq(haystack: &[u8], needle: &[u8]) -> bool {
    if needle.is_empty() || haystack.len() < needle.len() {
        return false;
    }
    haystack.windows(needle.len()).any(|w| w == needle)
}

fn local_name(qname: &[u8]) -> &[u8] {
    match qname.iter().position(|b| *b == b':') {
        Some(i) => &qname[i + 1..],
        None => qname,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::doc::Doc;
    use crate::doc::test_fixture::minimal_docx_bytes;
    use std::io::Write as _;
    use tempfile::NamedTempFile;

    fn load_minimal() -> (NamedTempFile, Doc) {
        let mut tmp = NamedTempFile::new().unwrap();
        tmp.write_all(&minimal_docx_bytes()).unwrap();
        tmp.flush().unwrap();
        let doc = Doc::load(tmp.path()).unwrap();
        (tmp, doc)
    }

    #[test]
    fn snapshot_minimal_fixture_has_one_paragraph() {
        let (_tmp, doc) = load_minimal();
        let snap = build_snapshot(&doc).unwrap();
        assert_eq!(snap.version, "1.0");
        assert_eq!(snap.available_styles, vec!["Title", "Body"]);
        assert_eq!(snap.body.len(), 1);
        match &snap.body[0] {
            BodyItem::Paragraph {
                reference,
                style,
                text,
            } => {
                assert_eq!(reference, "@p1");
                assert_eq!(style.as_deref(), Some("Title"));
                assert_eq!(text, "Hello");
            }
            _ => panic!("expected paragraph"),
        }
    }

    #[test]
    fn snapshot_metadata_from_core_xml() {
        let (_tmp, doc) = load_minimal();
        let snap = build_snapshot(&doc).unwrap();
        assert_eq!(snap.metadata.title.as_deref(), Some("Sample"));
        assert_eq!(snap.metadata.author.as_deref(), Some("Tester"));
    }

    #[test]
    fn paragraph_renders_bold_italic_markdown() {
        let xml = br#"<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:pPr><w:pStyle w:val="Body"/></w:pPr>
<w:r><w:t xml:space="preserve">plain </w:t></w:r>
<w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r>
<w:r><w:t xml:space="preserve"> and </w:t></w:r>
<w:r><w:rPr><w:i/></w:rPr><w:t>italic</w:t></w:r>
</w:p>
</w:body></w:document>"#;
        let idx = index_body(xml).unwrap();
        match &idx.items[0] {
            BodyItem::Paragraph { text, .. } => {
                assert_eq!(text, "plain **bold** and *italic*");
            }
            _ => panic!("expected paragraph"),
        }
    }

    #[test]
    fn ref_numbering_is_sequential_and_per_kind() {
        let xml = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t>one</w:t></w:r></w:p>
<w:p><w:r><w:t>two</w:t></w:r></w:p>
<w:tbl>
  <w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>
  <w:tr><w:tc><w:p><w:r><w:t>1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>2</w:t></w:r></w:p></w:tc></w:tr>
</w:tbl>
<w:p><w:r><w:t>three</w:t></w:r></w:p>
</w:body></w:document>"#;
        let idx = index_body(xml).unwrap();
        let refs: Vec<&str> = idx.items.iter().map(BodyItem::reference).collect();
        assert_eq!(refs, vec!["@p1", "@p2", "@t1", "@p3"]);
        // Paragraphs inside table cells must NOT consume @pN slots.
    }

    #[test]
    fn table_records_dimensions_and_header() {
        let xml =
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:tbl>
  <w:tr>
    <w:tc><w:p><w:r><w:t>Metric</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>Q3</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>Q4</w:t></w:r></w:p></w:tc>
  </w:tr>
  <w:tr>
    <w:tc><w:p><w:r><w:t>Revenue</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>$1M</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>$1.2M</w:t></w:r></w:p></w:tc>
  </w:tr>
</w:tbl>
</w:body></w:document>"#;
        let idx = index_body(xml).unwrap();
        match &idx.items[0] {
            BodyItem::Table {
                reference,
                rows,
                cols,
                header,
            } => {
                assert_eq!(reference, "@t1");
                assert_eq!(*rows, 2);
                assert_eq!(*cols, 3);
                assert_eq!(
                    header,
                    &vec!["Metric".to_string(), "Q3".into(), "Q4".into()]
                );
            }
            _ => panic!("expected table"),
        }
    }

    #[test]
    fn table_snapshot_drills_cells_with_refs() {
        let xml =
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:tbl>
  <w:tr>
    <w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
  </w:tr>
  <w:tr>
    <w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc>
    <w:tc><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc>
  </w:tr>
</w:tbl>
</w:body></w:document>"#;
        let mut tmp = NamedTempFile::new().unwrap();
        tmp.write_all(&minimal_docx_bytes()).unwrap();
        tmp.flush().unwrap();
        let mut doc = Doc::load(tmp.path()).unwrap();
        doc.parts.document_xml = xml.to_vec();
        let snap = build_table_snapshot(&doc, "@t1").unwrap();
        assert_eq!(snap.rows, 2);
        assert_eq!(snap.cols, 2);
        assert_eq!(snap.cells[0][0].reference, "@t1.r1.c1");
        assert_eq!(snap.cells[0][0].text, "A");
        assert_eq!(snap.cells[1][1].reference, "@t1.r2.c2");
        assert_eq!(snap.cells[1][1].text, "D");
    }

    #[test]
    fn table_snapshot_rejects_non_table_ref() {
        let (_tmp, doc) = load_minimal();
        let err = build_table_snapshot(&doc, "@p1").unwrap_err();
        assert!(matches!(err, DocxaiError::InvalidArgument(_)));
    }

    #[test]
    fn table_snapshot_rejects_missing_table() {
        let (_tmp, doc) = load_minimal();
        let err = build_table_snapshot(&doc, "@t9").unwrap_err();
        assert!(matches!(err, DocxaiError::InvalidArgument(_)));
    }

    #[test]
    fn escape_markdown_special_chars() {
        assert_eq!(
            escape_markdown(r"a*b_c`d$e[f]g\h"),
            r"a\*b\_c\`d\$e\[f\]g\\h"
        );
    }

    #[test]
    fn preserved_features_detects_footnotes_and_tracked_changes() {
        let (_tmp, doc) = load_minimal();
        let mut doc = doc;
        doc.parts
            .others
            .insert("word/footnotes.xml".into(), b"<w:footnotes/>".to_vec());
        doc.parts
            .others
            .insert("word/comments.xml".into(), b"<w:comments/>".to_vec());
        doc.parts.document_xml = br#"<w:document xmlns:w="x">
<w:body><w:p><w:ins w:id="1"><w:r><w:t>x</w:t></w:r></w:ins></w:p></w:body>
</w:document>"#
            .to_vec();
        let feats = detect_preserved_features(&doc);
        assert!(feats.contains(&"footnotes".to_string()));
        assert!(feats.contains(&"comments".to_string()));
        assert!(feats.contains(&"tracked_changes".to_string()));
    }
}
