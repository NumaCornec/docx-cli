//! Paragraph mutations for M2 (PRD #14–#18).
//!
//! Operations on the `word/document.xml` part of a loaded [`Doc`].
//! All mutations:
//! 1. Index body elements to find byte ranges
//! 2. Modify the XML bytes (splice / replace / remove)
//! 3. Save the document atomically via [`Doc::save`]

use quick_xml::Reader;
use quick_xml::events::Event;

use crate::doc::Doc;
use crate::error::DocxaiError;
use crate::markdown::{self, Run};
use crate::refs::Ref;
use crate::styles;

// ---------------------------------------------------------------------------
// Body indexing
// ---------------------------------------------------------------------------

struct BodySpan {
    kind: char,
    index: u32,
    start: usize,
    end: usize,
}

struct BodyMap {
    spans: Vec<BodySpan>,
    /// Byte offset of the first byte of `</w:body>`.
    body_end: usize,
}

fn index_body_spans(xml: &[u8]) -> Result<BodyMap, DocxaiError> {
    let mut reader = Reader::from_reader(xml);
    let mut buf = Vec::new();
    let mut spans = Vec::new();
    let mut body_end = xml.len();

    let mut in_body = false;
    let mut tbl_depth: u32 = 0;
    let mut p_count: u32 = 0;
    let mut t_count: u32 = 0;
    let mut pending_p_start: Option<usize> = None;
    let mut pending_t_start: Option<usize> = None;

    loop {
        let pos_before = reader.buffer_position() as usize;
        let event = reader
            .read_event_into(&mut buf)
            .map_err(|e| DocxaiError::Generic(format!("document.xml parse error: {e}")))?;
        let pos_after = reader.buffer_position() as usize;

        match event {
            Event::Eof => break,
            Event::Start(ref e) => {
                let local = local_name(e.name().as_ref()).to_owned();
                match local.as_slice() {
                    b"body" => in_body = true,
                    b"p" if in_body && tbl_depth == 0 => {
                        p_count += 1;
                        pending_p_start = Some(pos_before);
                    }
                    b"tbl" if in_body => {
                        if tbl_depth == 0 {
                            t_count += 1;
                            pending_t_start = Some(pos_before);
                        }
                        tbl_depth += 1;
                    }
                    _ => {}
                }
            }
            Event::End(ref e) => {
                let local = local_name(e.name().as_ref()).to_owned();
                match local.as_slice() {
                    b"body" => {
                        body_end = pos_before;
                        in_body = false;
                    }
                    b"p" if in_body && tbl_depth == 0 => {
                        if let Some(start) = pending_p_start.take() {
                            spans.push(BodySpan {
                                kind: 'p',
                                index: p_count,
                                start,
                                end: pos_after,
                            });
                        }
                    }
                    b"tbl" => {
                        tbl_depth = tbl_depth.saturating_sub(1);
                        if tbl_depth == 0 {
                            if let Some(start) = pending_t_start.take() {
                                spans.push(BodySpan {
                                    kind: 't',
                                    index: t_count,
                                    start,
                                    end: pos_after,
                                });
                            }
                        }
                    }
                    _ => {}
                }
            }
            Event::Empty(ref e) => {
                let local = local_name(e.name().as_ref()).to_owned();
                if in_body && tbl_depth == 0 && local.as_slice() == b"p" {
                    p_count += 1;
                    spans.push(BodySpan {
                        kind: 'p',
                        index: p_count,
                        start: pos_before,
                        end: pos_after,
                    });
                }
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(BodyMap { spans, body_end })
}

fn find_span<'a>(spans: &'a [BodySpan], parsed: &Ref) -> Result<&'a BodySpan, DocxaiError> {
    let (kind, n) = match parsed {
        Ref::Paragraph(n) => ('p', *n),
        Ref::Table(n) => ('t', *n),
        _ => {
            return Err(DocxaiError::InvalidArgument(format!(
                "ref {} is not a paragraph or table",
                parsed
            )));
        }
    };
    spans
        .iter()
        .find(|s| s.kind == kind && s.index == n)
        .ok_or_else(|| ref_not_found(parsed, spans))
}

fn ref_not_found(parsed: &Ref, spans: &[BodySpan]) -> DocxaiError {
    match parsed {
        Ref::Paragraph(n) => {
            let max = spans.iter().filter(|s| s.kind == 'p').count();
            DocxaiError::InvalidArgument(format!(
                "ref @p{n} not found (document has {max} paragraphs)"
            ))
        }
        Ref::Table(n) => {
            let max = spans.iter().filter(|s| s.kind == 't').count();
            DocxaiError::InvalidArgument(format!("ref @t{n} not found (document has {max} tables)"))
        }
        _ => DocxaiError::InvalidArgument(format!("ref {parsed} not found")),
    }
}

fn count_paragraphs_before(spans: &[BodySpan], pos: usize) -> u32 {
    spans
        .iter()
        .filter(|s| s.kind == 'p' && s.end <= pos)
        .count() as u32
}

fn determine_insert_pos(
    map: &BodyMap,
    after: Option<&str>,
    before: Option<&str>,
) -> Result<usize, DocxaiError> {
    match (after, before) {
        (Some(ref_str), None) => {
            let parsed = Ref::parse(ref_str)?;
            let span = find_span(&map.spans, &parsed)?;
            Ok(span.end)
        }
        (None, Some(ref_str)) => {
            let parsed = Ref::parse(ref_str)?;
            let span = find_span(&map.spans, &parsed)?;
            Ok(span.start)
        }
        (None, None) => Ok(map.body_end),
        _ => unreachable!("--after and --before are mutually exclusive per clap"),
    }
}

// ---------------------------------------------------------------------------
// XML building
// ---------------------------------------------------------------------------

fn build_paragraph_xml(runs: &[Run], style: Option<&str>) -> String {
    let mut xml = String::from("<w:p>");
    if let Some(s) = style {
        xml.push_str("<w:pPr><w:pStyle w:val=\"");
        xml.push_str(&xml_escape_attr(s));
        xml.push_str("\"/></w:pPr>");
    }
    for run in runs {
        emit_run_xml(&mut xml, run);
    }
    xml.push_str("</w:p>");
    xml
}

fn emit_run_xml(xml: &mut String, run: &Run) {
    let text = &run.text;
    if text.is_empty() {
        return;
    }

    let has_special = text.contains('\n') || text.contains('\t');
    if !has_special {
        emit_text_run(xml, text, run.bold, run.italic);
        return;
    }

    let mut last = 0;
    for (i, b) in text.bytes().enumerate() {
        if b == b'\n' || b == b'\t' {
            if i > last {
                emit_text_run(xml, &text[last..i], run.bold, run.italic);
            }
            if b == b'\n' {
                emit_break_run(xml);
            } else {
                emit_tab_run(xml);
            }
            last = i + 1;
        }
    }
    if last < text.len() {
        emit_text_run(xml, &text[last..], run.bold, run.italic);
    }
}

fn emit_text_run(xml: &mut String, text: &str, bold: bool, italic: bool) {
    xml.push_str("<w:r>");
    if bold || italic {
        xml.push_str("<w:rPr>");
        if bold {
            xml.push_str("<w:b/>");
        }
        if italic {
            xml.push_str("<w:i/>");
        }
        xml.push_str("</w:rPr>");
    }
    let space = if text.starts_with(' ') || text.ends_with(' ') {
        r#" xml:space="preserve""#
    } else {
        ""
    };
    xml.push_str("<w:t");
    xml.push_str(space);
    xml.push('>');
    xml.push_str(&xml_escape_text(text));
    xml.push_str("</w:t></w:r>");
}

fn emit_break_run(xml: &mut String) {
    xml.push_str("<w:r><w:br/></w:r>");
}

fn emit_tab_run(xml: &mut String) {
    xml.push_str("<w:r><w:tab/></w:r>");
}

// ---------------------------------------------------------------------------
// Style manipulation
// ---------------------------------------------------------------------------

fn extract_paragraph_style(para_bytes: &[u8]) -> Option<String> {
    let hay = std::str::from_utf8(para_bytes).ok()?;
    let start = hay.find("<w:pStyle")?;
    let val_start = hay[start..].find("w:val=\"")? + start + 7;
    let val_end = hay[val_start..].find('"')? + val_start;
    Some(hay[val_start..val_end].to_string())
}

fn replace_style_in_bytes(para: &mut Vec<u8>, new_style: &str) -> Result<(), DocxaiError> {
    if let Some(val_range) = find_pstyle_val_range(para) {
        let escaped = xml_escape_attr(new_style);
        para.splice(val_range, escaped.as_bytes().iter().copied());
        return Ok(());
    }
    insert_style_element(para, new_style)
}

fn find_pstyle_val_range(para: &[u8]) -> Option<std::ops::Range<usize>> {
    let pstyle_pos = find_subseq_offset(para, b"<w:pStyle")?;
    let val_prefix = b"w:val=\"";
    let val_pos = find_subseq_offset_from(para, val_prefix, pstyle_pos)?;
    let val_start = val_pos + val_prefix.len();
    let rel_end = para[val_start..].iter().position(|&b| b == b'"')?;
    Some(val_start..val_start + rel_end)
}

fn insert_style_element(para: &mut Vec<u8>, style: &str) -> Result<(), DocxaiError> {
    let escaped = xml_escape_attr(style);
    let pstyle = format!("<w:pStyle w:val=\"{}\"/>", escaped);

    if let Some(pos) = find_subseq_offset(para, b"<w:pPr>") {
        let insert_at = pos + 7;
        para.splice(insert_at..insert_at, pstyle.as_bytes().iter().copied());
        return Ok(());
    }

    if let Some(pos) = find_subseq_offset(para, b"<w:pPr ") {
        if let Some(end) = para[pos..].iter().position(|&b| b == b'>') {
            let insert_at = pos + end + 1;
            para.splice(insert_at..insert_at, pstyle.as_bytes().iter().copied());
            return Ok(());
        }
    }

    let ppr = format!("<w:pPr>{}</w:pPr>", pstyle);
    let open_end = find_element_open_end(para, b"<w:p")
        .ok_or_else(|| DocxaiError::Generic("cannot find <w:p> opening tag".into()))?;
    para.splice(open_end..open_end, ppr.as_bytes().iter().copied());
    Ok(())
}

// ---------------------------------------------------------------------------
// Preservation checks
// ---------------------------------------------------------------------------

fn has_tracked_changes(para_bytes: &[u8]) -> bool {
    contains_any(
        para_bytes,
        &[
            b"<w:ins ".as_slice(),
            b"<w:del ".as_slice(),
            b"<w:moveFrom ".as_slice(),
            b"<w:moveTo ".as_slice(),
        ],
    )
}

fn has_footnote_reference(para_bytes: &[u8]) -> bool {
    contains_any(
        para_bytes,
        &[
            b"<w:footnoteReference".as_slice(),
            b"<w:endnoteReference".as_slice(),
        ],
    )
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

pub fn add_paragraph(
    doc: &mut Doc,
    text: &str,
    style: Option<&str>,
    after: Option<&str>,
    before: Option<&str>,
) -> Result<serde_json::Value, DocxaiError> {
    let runs = markdown::parse_runs(text)?;

    let available = styles::list_paragraph_styles(doc)?;
    let resolved_style = match style {
        Some(s) => {
            if !available.contains(&s.to_string()) {
                return Err(DocxaiError::InvalidArgument(format!(
                    "style {:?} not found in document. Available: {:?}",
                    s, available
                )));
            }
            Some(s.to_string())
        }
        None => available
            .iter()
            .find(|s| s.as_str() == "Body")
            .cloned()
            .or_else(|| available.first().cloned()),
    };

    let xml = build_paragraph_xml(&runs, resolved_style.as_deref());

    let map = index_body_spans(&doc.parts.document_xml)?;
    let insert_pos = determine_insert_pos(&map, after, before)?;
    let new_index = count_paragraphs_before(&map.spans, insert_pos) + 1;

    let mut bytes = std::mem::take(&mut doc.parts.document_xml);
    bytes.splice(insert_pos..insert_pos, xml.as_bytes().iter().copied());
    doc.parts.document_xml = bytes;

    doc.save(&doc.path)?;

    Ok(serde_json::json!({
        "status": "ok",
        "action": "add",
        "ref": format!("@p{}", new_index),
        "kind": "paragraph",
        "style": resolved_style.unwrap_or_default(),
    }))
}

pub fn set_paragraph(
    doc: &mut Doc,
    reference: &str,
    text: Option<&str>,
    style: Option<&str>,
) -> Result<serde_json::Value, DocxaiError> {
    let parsed = Ref::parse(reference)?;
    match &parsed {
        Ref::Paragraph(_) => {}
        _ => {
            return Err(DocxaiError::InvalidArgument(format!(
                "expected paragraph ref (@pN), got {}",
                reference
            )));
        }
    }

    if text.is_none() && style.is_none() {
        return Err(DocxaiError::InvalidArgument(
            "set @pN requires at least one of --text or --style".into(),
        ));
    }

    if let Some(s) = style {
        let available = styles::list_paragraph_styles(doc)?;
        if !available.contains(&s.to_string()) {
            return Err(DocxaiError::InvalidArgument(format!(
                "style {:?} not found in document. Available: {:?}",
                s, available
            )));
        }
    }

    let map = index_body_spans(&doc.parts.document_xml)?;
    let span = find_span(&map.spans, &parsed)?;
    let para_bytes = &doc.parts.document_xml[span.start..span.end];

    let runs = match text {
        Some(t) => {
            if has_tracked_changes(para_bytes) {
                return Err(DocxaiError::PreservationImpossible(format!(
                    "paragraph {} contains tracked changes; cannot modify text",
                    reference
                )));
            }
            Some(markdown::parse_runs(t)?)
        }
        None => None,
    };

    let mut changed = Vec::new();

    match (runs, style) {
        (Some(runs), new_style) => {
            let resolved_style = new_style
                .map(|s| s.to_string())
                .or_else(|| extract_paragraph_style(para_bytes));
            let xml = build_paragraph_xml(&runs, resolved_style.as_deref());

            changed.push("text");
            if style.is_some() {
                changed.push("style");
            }

            let mut bytes = std::mem::take(&mut doc.parts.document_xml);
            bytes.splice(span.start..span.end, xml.as_bytes().iter().copied());
            doc.parts.document_xml = bytes;
        }
        (None, Some(new_style)) => {
            let mut para = doc.parts.document_xml[span.start..span.end].to_vec();
            replace_style_in_bytes(&mut para, new_style)?;

            let mut bytes = std::mem::take(&mut doc.parts.document_xml);
            bytes.splice(span.start..span.end, para);
            doc.parts.document_xml = bytes;

            changed.push("style");
        }
        (None, None) => unreachable!("checked above"),
    }

    doc.save(&doc.path)?;

    Ok(serde_json::json!({
        "status": "ok",
        "action": "set",
        "ref": reference,
        "changed": changed,
    }))
}

pub fn delete_paragraph(doc: &mut Doc, reference: &str) -> Result<serde_json::Value, DocxaiError> {
    let parsed = Ref::parse(reference)?;
    match &parsed {
        Ref::Paragraph(_) => {}
        _ => {
            return Err(DocxaiError::InvalidArgument(format!(
                "expected paragraph ref (@pN), got {}",
                reference
            )));
        }
    }

    let map = index_body_spans(&doc.parts.document_xml)?;
    let span = find_span(&map.spans, &parsed)?;
    let para_bytes = &doc.parts.document_xml[span.start..span.end];

    if has_footnote_reference(para_bytes) {
        return Err(DocxaiError::PreservationImpossible(format!(
            "paragraph {} contains footnote/endnote references; deleting would break notes",
            reference
        )));
    }

    let mut bytes = std::mem::take(&mut doc.parts.document_xml);
    bytes.splice(span.start..span.end, std::iter::empty::<u8>());
    doc.parts.document_xml = bytes;

    doc.save(&doc.path)?;

    Ok(serde_json::json!({
        "status": "ok",
        "action": "delete",
        "ref": reference,
    }))
}

// ---------------------------------------------------------------------------
// Utilities
// ---------------------------------------------------------------------------

fn local_name(qname: &[u8]) -> &[u8] {
    match qname.iter().position(|b| *b == b':') {
        Some(i) => &qname[i + 1..],
        None => qname,
    }
}

fn xml_escape_text(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for ch in s.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            _ => out.push(ch),
        }
    }
    out
}

fn xml_escape_attr(s: &str) -> String {
    let mut out = String::with_capacity(s.len());
    for ch in s.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            _ => out.push(ch),
        }
    }
    out
}

fn contains_any(haystack: &[u8], needles: &[&[u8]]) -> bool {
    needles
        .iter()
        .any(|needle| contains_subseq(haystack, needle))
}

fn contains_subseq(haystack: &[u8], needle: &[u8]) -> bool {
    if needle.is_empty() || haystack.len() < needle.len() {
        return false;
    }
    haystack.windows(needle.len()).any(|w| w == needle)
}

fn find_subseq_offset(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    if needle.is_empty() || haystack.len() < needle.len() {
        return None;
    }
    haystack.windows(needle.len()).position(|w| w == needle)
}

fn find_subseq_offset_from(haystack: &[u8], needle: &[u8], from: usize) -> Option<usize> {
    if from >= haystack.len() {
        return None;
    }
    find_subseq_offset(&haystack[from..], needle).map(|pos| from + pos)
}

fn find_element_open_end(bytes: &[u8], tag_prefix: &[u8]) -> Option<usize> {
    let pos = find_subseq_offset(bytes, tag_prefix)?;
    let after = pos + tag_prefix.len();
    for i in after..bytes.len() {
        if bytes[i] == b'>' {
            if i > 0 && bytes[i - 1] == b'/' {
                return None;
            }
            return Some(i + 1);
        }
    }
    None
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;
    use crate::doc::test_fixture::minimal_docx_bytes;
    use std::io::Write as IoWrite;
    use tempfile::NamedTempFile;

    fn load_doc() -> (NamedTempFile, Doc) {
        let mut tmp = NamedTempFile::new().unwrap();
        tmp.write_all(&minimal_docx_bytes()).unwrap();
        tmp.flush().unwrap();
        let doc = Doc::load(tmp.path()).unwrap();
        (tmp, doc)
    }

    fn load_doc_with_xml(document_xml: &[u8]) -> (NamedTempFile, Doc) {
        let mut tmp = NamedTempFile::new().unwrap();
        tmp.write_all(&minimal_docx_bytes()).unwrap();
        tmp.flush().unwrap();
        let mut doc = Doc::load(tmp.path()).unwrap();
        doc.parts.document_xml = document_xml.to_vec();
        (tmp, doc)
    }

    fn reindex_and_count_p(xml: &[u8]) -> u32 {
        let map = index_body_spans(xml).unwrap();
        map.spans.iter().filter(|s| s.kind == 'p').count() as u32
    }

    // -- #14 add paragraph (append) --

    #[test]
    fn add_paragraph_appends_to_end() {
        let (_tmp, mut doc) = load_doc();
        let result = add_paragraph(&mut doc, "World", None, None, None).unwrap();
        assert_eq!(result["status"], "ok");
        assert_eq!(result["action"], "add");
        assert_eq!(result["ref"], "@p2");
        assert_eq!(result["kind"], "paragraph");
        assert_eq!(result["style"], "Body");

        let reloaded = Doc::load(doc.path).unwrap();
        let map = index_body_spans(&reloaded.parts.document_xml).unwrap();
        let p_spans: Vec<_> = map.spans.iter().filter(|s| s.kind == 'p').collect();
        assert_eq!(p_spans.len(), 2);

        let para_bytes = &reloaded.parts.document_xml[p_spans[1].start..p_spans[1].end];
        let s = std::str::from_utf8(para_bytes).unwrap();
        assert!(s.contains("World"));
        assert!(s.contains(r#"w:val="Body""#));
    }

    #[test]
    fn add_paragraph_with_explicit_style() {
        let (_tmp, mut doc) = load_doc();
        let result = add_paragraph(&mut doc, "Styled", Some("Title"), None, None).unwrap();
        assert_eq!(result["ref"], "@p2");
        assert_eq!(result["style"], "Title");

        let reloaded = Doc::load(doc.path).unwrap();
        let map = index_body_spans(&reloaded.parts.document_xml).unwrap();
        let p2 = map
            .spans
            .iter()
            .find(|s| s.kind == 'p' && s.index == 2)
            .unwrap();
        let para_bytes = &reloaded.parts.document_xml[p2.start..p2.end];
        let s = std::str::from_utf8(para_bytes).unwrap();
        assert!(s.contains(r#"w:val="Title""#));
    }

    #[test]
    fn add_paragraph_rejects_unknown_style() {
        let (_tmp, mut doc) = load_doc();
        let err = add_paragraph(&mut doc, "text", Some("NonExistent"), None, None).unwrap_err();
        assert!(matches!(err, DocxaiError::InvalidArgument(_)));
    }

    #[test]
    fn add_paragraph_with_bold_italic_markdown() {
        let (_tmp, mut doc) = load_doc();
        add_paragraph(&mut doc, "**bold** and *italic*", None, None, None).unwrap();

        let reloaded = Doc::load(doc.path).unwrap();
        let map = index_body_spans(&reloaded.parts.document_xml).unwrap();
        let p2 = map
            .spans
            .iter()
            .find(|s| s.kind == 'p' && s.index == 2)
            .unwrap();
        let para_bytes = &reloaded.parts.document_xml[p2.start..p2.end];
        let s = std::str::from_utf8(para_bytes).unwrap();
        assert!(s.contains("<w:b/>"), "expected bold: {s}");
        assert!(s.contains("<w:i/>"), "expected italic: {s}");
    }

    #[test]
    fn add_paragraph_preserves_other_parts() {
        let (_tmp, mut doc) = load_doc();
        let orig_styles = doc.parts.styles_xml.clone();
        let orig_rels = doc.parts.document_rels.clone();
        let orig_others = doc.parts.others.clone();

        add_paragraph(&mut doc, "extra", None, None, None).unwrap();

        let reloaded = Doc::load(doc.path).unwrap();
        assert_eq!(reloaded.parts.styles_xml, orig_styles);
        assert_eq!(reloaded.parts.document_rels, orig_rels);
        assert_eq!(reloaded.parts.others, orig_others);
    }

    // -- #15 add paragraph --after / --before --

    #[test]
    fn add_paragraph_after_ref() {
        let xml =
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body><w:p><w:r><w:t>one</w:t></w:r></w:p><w:p><w:r><w:t>two</w:t></w:r></w:p></w:body>
</w:document>"#;
        let (_tmp, mut doc) = load_doc_with_xml(xml);
        let result = add_paragraph(&mut doc, "inserted", None, Some("@p1"), None).unwrap();
        assert_eq!(result["ref"], "@p2");

        let reloaded = Doc::load(doc.path).unwrap();
        assert_eq!(reindex_and_count_p(&reloaded.parts.document_xml), 3);

        let map = index_body_spans(&reloaded.parts.document_xml).unwrap();
        let p1 = &map.spans[0];
        let p2 = &map.spans[1];
        let p3 = &map.spans[2];
        let b1 = std::str::from_utf8(&reloaded.parts.document_xml[p1.start..p1.end]).unwrap();
        let b2 = std::str::from_utf8(&reloaded.parts.document_xml[p2.start..p2.end]).unwrap();
        let b3 = std::str::from_utf8(&reloaded.parts.document_xml[p3.start..p3.end]).unwrap();
        assert!(b1.contains("one"));
        assert!(b2.contains("inserted"));
        assert!(b3.contains("two"));
    }

    #[test]
    fn add_paragraph_before_ref() {
        let xml =
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body><w:p><w:r><w:t>one</w:t></w:r></w:p><w:p><w:r><w:t>two</w:t></w:r></w:p></w:body>
</w:document>"#;
        let (_tmp, mut doc) = load_doc_with_xml(xml);
        let result = add_paragraph(&mut doc, "inserted", None, None, Some("@p2")).unwrap();
        assert_eq!(result["ref"], "@p2");

        let reloaded = Doc::load(doc.path).unwrap();
        let map = index_body_spans(&reloaded.parts.document_xml).unwrap();
        let b2 =
            std::str::from_utf8(&reloaded.parts.document_xml[map.spans[1].start..map.spans[1].end])
                .unwrap();
        assert!(b2.contains("inserted"));
    }

    #[test]
    fn add_paragraph_after_invalid_ref_errors() {
        let (_tmp, mut doc) = load_doc();
        let err = add_paragraph(&mut doc, "text", None, Some("@p99"), None).unwrap_err();
        assert!(matches!(err, DocxaiError::InvalidArgument(_)));
    }

    // -- #16 set @pN --text --

    #[test]
    fn set_paragraph_text() {
        let (_tmp, mut doc) = load_doc();
        let result = set_paragraph(&mut doc, "@p1", Some("Replaced"), None).unwrap();
        assert_eq!(result["status"], "ok");
        assert_eq!(result["changed"], serde_json::json!(["text"]));

        let reloaded = Doc::load(doc.path).unwrap();
        let map = index_body_spans(&reloaded.parts.document_xml).unwrap();
        let p1 = &map.spans[0];
        let para_bytes = &reloaded.parts.document_xml[p1.start..p1.end];
        let s = std::str::from_utf8(para_bytes).unwrap();
        assert!(s.contains("Replaced"));
        assert!(s.contains(r#"w:val="Title""#), "style should be preserved");
    }

    #[test]
    fn set_paragraph_text_with_markdown() {
        let (_tmp, mut doc) = load_doc();
        set_paragraph(&mut doc, "@p1", Some("**bold** text"), None).unwrap();

        let reloaded = Doc::load(doc.path).unwrap();
        let map = index_body_spans(&reloaded.parts.document_xml).unwrap();
        let para_bytes = &reloaded.parts.document_xml[map.spans[0].start..map.spans[0].end];
        let s = std::str::from_utf8(para_bytes).unwrap();
        assert!(s.contains("<w:b/>"), "expected bold run: {s}");
        assert!(s.contains("bold"));
        assert!(s.contains("text"));
    }

    #[test]
    fn set_paragraph_text_rejects_tracked_changes() {
        let xml =
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body><w:p><w:ins w:id="1"><w:r><w:t>tracked</w:t></w:r></w:ins></w:p></w:body>
</w:document>"#;
        let (_tmp, mut doc) = load_doc_with_xml(xml);
        let err = set_paragraph(&mut doc, "@p1", Some("new"), None).unwrap_err();
        assert!(matches!(err, DocxaiError::PreservationImpossible(_)));
    }

    #[test]
    fn set_paragraph_text_invalid_ref_errors() {
        let (_tmp, mut doc) = load_doc();
        let err = set_paragraph(&mut doc, "@p99", Some("x"), None).unwrap_err();
        assert!(matches!(err, DocxaiError::InvalidArgument(_)));
    }

    // -- #17 set @pN --style --

    #[test]
    fn set_paragraph_style_surgical() {
        let (_tmp, mut doc) = load_doc();
        let result = set_paragraph(&mut doc, "@p1", None, Some("Body")).unwrap();
        assert_eq!(result["changed"], serde_json::json!(["style"]));

        let reloaded = Doc::load(doc.path).unwrap();
        let map = index_body_spans(&reloaded.parts.document_xml).unwrap();
        let para_bytes = &reloaded.parts.document_xml[map.spans[0].start..map.spans[0].end];
        let s = std::str::from_utf8(para_bytes).unwrap();
        assert!(s.contains(r#"w:val="Body""#), "style should be Body: {s}");
        assert!(s.contains("Hello"), "text should be preserved: {s}");
    }

    #[test]
    fn set_paragraph_style_rejects_unknown() {
        let (_tmp, mut doc) = load_doc();
        let err = set_paragraph(&mut doc, "@p1", None, Some("NonExistent")).unwrap_err();
        assert!(matches!(err, DocxaiError::InvalidArgument(_)));
    }

    #[test]
    fn set_paragraph_text_and_style_together() {
        let (_tmp, mut doc) = load_doc();
        let result = set_paragraph(&mut doc, "@p1", Some("New text"), Some("Body")).unwrap();
        assert_eq!(result["changed"], serde_json::json!(["text", "style"]));

        let reloaded = Doc::load(doc.path).unwrap();
        let map = index_body_spans(&reloaded.parts.document_xml).unwrap();
        let para_bytes = &reloaded.parts.document_xml[map.spans[0].start..map.spans[0].end];
        let s = std::str::from_utf8(para_bytes).unwrap();
        assert!(s.contains("New text"));
        assert!(s.contains(r#"w:val="Body""#));
    }

    #[test]
    fn set_paragraph_nothing_errors() {
        let (_tmp, mut doc) = load_doc();
        let err = set_paragraph(&mut doc, "@p1", None, None).unwrap_err();
        assert!(matches!(err, DocxaiError::InvalidArgument(_)));
    }

    #[test]
    fn set_paragraph_style_on_para_without_ppr() {
        let xml =
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body><w:p><w:r><w:t>no pPr</w:t></w:r></w:p></w:body></w:document>"#;
        let (_tmp, mut doc) = load_doc_with_xml(xml);
        set_paragraph(&mut doc, "@p1", None, Some("Title")).unwrap();

        let reloaded = Doc::load(doc.path).unwrap();
        let map = index_body_spans(&reloaded.parts.document_xml).unwrap();
        let para_bytes = &reloaded.parts.document_xml[map.spans[0].start..map.spans[0].end];
        let s = std::str::from_utf8(para_bytes).unwrap();
        assert!(s.contains(r#"w:val="Title""#), "style inserted: {s}");
        assert!(s.contains("no pPr"), "text preserved: {s}");
    }

    // -- #18 delete @pN --

    #[test]
    fn delete_paragraph_removes_element() {
        let xml =
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body><w:p><w:r><w:t>one</w:t></w:r></w:p><w:p><w:r><w:t>two</w:t></w:r></w:p></w:body>
</w:document>"#;
        let (_tmp, mut doc) = load_doc_with_xml(xml);
        let result = delete_paragraph(&mut doc, "@p1").unwrap();
        assert_eq!(result["status"], "ok");
        assert_eq!(result["action"], "delete");
        assert_eq!(result["ref"], "@p1");

        let reloaded = Doc::load(doc.path).unwrap();
        assert_eq!(reindex_and_count_p(&reloaded.parts.document_xml), 1);

        let map = index_body_spans(&reloaded.parts.document_xml).unwrap();
        let p1 = &map.spans[0];
        let s = std::str::from_utf8(&reloaded.parts.document_xml[p1.start..p1.end]).unwrap();
        assert!(s.contains("two"), "remaining paragraph should be 'two'");
    }

    #[test]
    fn delete_paragraph_with_footnote_errors() {
        let xml =
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body><w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p></w:body></w:document>"#;
        let (_tmp, mut doc) = load_doc_with_xml(xml);
        let err = delete_paragraph(&mut doc, "@p1").unwrap_err();
        assert!(matches!(err, DocxaiError::PreservationImpossible(_)));
    }

    #[test]
    fn delete_paragraph_invalid_ref_errors() {
        let (_tmp, mut doc) = load_doc();
        let err = delete_paragraph(&mut doc, "@p99").unwrap_err();
        assert!(matches!(err, DocxaiError::InvalidArgument(_)));
    }

    // -- XML building --

    #[test]
    fn build_paragraph_xml_with_style_and_runs() {
        let runs = vec![
            Run {
                text: "hello ".into(),
                bold: false,
                italic: false,
            },
            Run {
                text: "world".into(),
                bold: true,
                italic: false,
            },
        ];
        let xml = build_paragraph_xml(&runs, Some("Body"));
        assert!(xml.starts_with("<w:p><w:pPr><w:pStyle w:val=\"Body\"/></w:pPr>"));
        assert!(xml.contains(r#"<w:r><w:rPr><w:b/></w:rPr><w:t>world</w:t></w:r>"#));
        assert!(xml.ends_with("</w:p>"));
    }

    #[test]
    fn build_paragraph_xml_no_style() {
        let runs = vec![Run {
            text: "plain".into(),
            bold: false,
            italic: false,
        }];
        let xml = build_paragraph_xml(&runs, None);
        assert!(!xml.contains("<w:pPr>"));
        assert!(xml.contains("<w:t>plain</w:t>"));
    }

    #[test]
    fn build_paragraph_xml_preserves_space() {
        let runs = vec![Run {
            text: " leading space".into(),
            bold: false,
            italic: false,
        }];
        let xml = build_paragraph_xml(&runs, None);
        assert!(xml.contains(r#"xml:space="preserve""#));
    }

    #[test]
    fn build_paragraph_xml_hard_break() {
        let runs = vec![Run {
            text: "line1\nline2".into(),
            bold: false,
            italic: false,
        }];
        let xml = build_paragraph_xml(&runs, None);
        assert!(xml.contains("<w:br/>"), "expected br element: {xml}");
        assert!(xml.contains("line1"));
        assert!(xml.contains("line2"));
    }

    #[test]
    fn xml_escape_text_escapes_special_chars() {
        assert_eq!(xml_escape_text("a<b>c&d"), "a&lt;b&gt;c&amp;d");
    }

    #[test]
    fn xml_escape_attr_escapes_quotes() {
        assert_eq!(xml_escape_attr(r#"style"name"#), r#"style&quot;name"#);
    }

    #[test]
    fn add_paragraph_to_empty_body() {
        let xml =
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body></w:body></w:document>"#;
        let (_tmp, mut doc) = load_doc_with_xml(xml);
        let result = add_paragraph(&mut doc, "first", None, None, None).unwrap();
        assert_eq!(result["ref"], "@p1");

        let reloaded = Doc::load(doc.path).unwrap();
        assert_eq!(reindex_and_count_p(&reloaded.parts.document_xml), 1);
    }

    #[test]
    fn set_style_inserts_into_ppr_without_pstyle() {
        let xml =
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body><w:p><w:pPr><w:spacing w:after="200"/></w:pPr><w:r><w:t>text</w:t></w:r></w:p></w:body>
</w:document>"#;
        let (_tmp, mut doc) = load_doc_with_xml(xml);
        set_paragraph(&mut doc, "@p1", None, Some("Title")).unwrap();

        let reloaded = Doc::load(doc.path).unwrap();
        let map = index_body_spans(&reloaded.parts.document_xml).unwrap();
        let s =
            std::str::from_utf8(&reloaded.parts.document_xml[map.spans[0].start..map.spans[0].end])
                .unwrap();
        assert!(
            s.contains(r#"w:val="Title""#),
            "should have Title style: {s}"
        );
        assert!(s.contains("spacing"), "spacing should be preserved: {s}");
    }

    #[test]
    fn add_paragraph_roundtrip_snapshot_verifies() {
        let (_tmp, mut doc) = load_doc();
        add_paragraph(&mut doc, "**bold** and plain", Some("Body"), None, None).unwrap();

        let reloaded = Doc::load(doc.path).unwrap();
        let snap = crate::snapshot::build_snapshot(&reloaded).unwrap();

        assert_eq!(snap.body.len(), 2);
        match &snap.body[1] {
            crate::snapshot::BodyItem::Paragraph {
                reference,
                style,
                text,
            } => {
                assert_eq!(reference, "@p2");
                assert_eq!(style.as_deref(), Some("Body"));
                assert_eq!(text, "**bold** and plain");
            }
            _ => panic!("expected paragraph"),
        }
    }
}
