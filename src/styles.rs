//! `word/styles.xml` parser for the `styles` command (PRD §8 / #9).
//!
//! Lists paragraph-style IDs available for use with `--style`.
//! Filters out `semiHidden` styles so the agent only sees what's user-facing.

use quick_xml::Reader;
use quick_xml::events::Event;

use crate::doc::Doc;
use crate::error::DocxaiError;

/// Return paragraph-style IDs in the order they appear in `styles.xml`.
///
/// Empty when the document has no `styles.xml` part. Hidden styles
/// (`<w:semiHidden/>`) are dropped per PRD #9 ("ne garder que ceux utilisables").
pub fn list_paragraph_styles(doc: &Doc) -> Result<Vec<String>, DocxaiError> {
    let Some(styles_xml) = doc.parts.styles_xml.as_deref() else {
        return Ok(Vec::new());
    };
    parse_paragraph_styles(styles_xml)
}

fn parse_paragraph_styles(xml: &[u8]) -> Result<Vec<String>, DocxaiError> {
    let mut reader = Reader::from_reader(xml);
    // No text trimming needed: we only inspect element names + attributes here.

    let mut out = Vec::new();
    let mut buf = Vec::new();

    // State for the currently open <w:style> element.
    let mut in_paragraph_style: Option<String> = None;
    let mut hidden = false;

    loop {
        let event = reader.read_event_into(&mut buf).map_err(|e| {
            DocxaiError::Generic(format!(
                "styles.xml parse error at {}: {e}",
                reader.buffer_position()
            ))
        })?;
        let (e, is_empty) = match event {
            Event::Eof => break,
            Event::Start(ref e) => (e, false),
            Event::Empty(ref e) => (e, true),
            Event::End(ref e) => {
                if local_name(e.name().as_ref()) == b"style" {
                    if let Some(id) = in_paragraph_style.take() {
                        if !hidden {
                            out.push(id);
                        }
                    }
                    hidden = false;
                }
                buf.clear();
                continue;
            }
            _ => {
                buf.clear();
                continue;
            }
        };

        let local = local_name(e.name().as_ref()).to_owned();
        match local.as_slice() {
            b"style" => {
                let mut is_paragraph = false;
                let mut style_id: Option<String> = None;
                for attr in e.attributes().flatten() {
                    let key = local_name(attr.key.as_ref()).to_owned();
                    let val = attr
                        .decode_and_unescape_value(reader.decoder())
                        .map_err(|err| {
                            DocxaiError::Generic(format!("styles.xml attr decode: {err}"))
                        })?
                        .into_owned();
                    match key.as_slice() {
                        b"type" => is_paragraph = val == "paragraph",
                        b"styleId" => style_id = Some(val),
                        _ => {}
                    }
                }
                in_paragraph_style = if is_paragraph { style_id } else { None };
                hidden = false;
                if is_empty {
                    if let Some(id) = in_paragraph_style.take() {
                        out.push(id);
                    }
                }
            }
            b"semiHidden" if in_paragraph_style.is_some() => {
                hidden = true;
            }
            _ => {}
        }
        buf.clear();
    }

    Ok(out)
}

/// Strip an optional `ns:` prefix from a qualified XML name.
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
    use std::io::Write;
    use tempfile::NamedTempFile;

    fn parse(xml: &str) -> Vec<String> {
        parse_paragraph_styles(xml.as_bytes()).expect("parse")
    }

    #[test]
    fn extracts_paragraph_styles_in_document_order() {
        let xml = r#"<?xml version="1.0"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style>
  <w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body"/></w:style>
</w:styles>"#;
        assert_eq!(parse(xml), vec!["Title", "Heading1", "Body"]);
    }

    #[test]
    fn ignores_non_paragraph_styles() {
        let xml = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="character" w:styleId="Code"/>
  <w:style w:type="table" w:styleId="TableGrid"/>
  <w:style w:type="numbering" w:styleId="NoList"/>
  <w:style w:type="paragraph" w:styleId="Body"/>
</w:styles>"#;
        assert_eq!(parse(xml), vec!["Body"]);
    }

    #[test]
    fn drops_semi_hidden_paragraph_styles() {
        let xml = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Hidden"><w:semiHidden/></w:style>
  <w:style w:type="paragraph" w:styleId="Visible"/>
</w:styles>"#;
        assert_eq!(parse(xml), vec!["Visible"]);
    }

    #[test]
    fn semi_hidden_inside_character_style_does_not_leak() {
        // semiHidden in a character style should not affect the next paragraph style.
        let xml = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="character" w:styleId="Hyperlink"><w:semiHidden/></w:style>
  <w:style w:type="paragraph" w:styleId="Body"/>
</w:styles>"#;
        assert_eq!(parse(xml), vec!["Body"]);
    }

    #[test]
    fn handles_missing_styles_xml() {
        let mut tmp = NamedTempFile::new().unwrap();
        tmp.write_all(&minimal_docx_bytes()).unwrap();
        tmp.flush().unwrap();
        let mut doc = Doc::load(tmp.path()).unwrap();
        doc.parts.styles_xml = None;
        assert!(list_paragraph_styles(&doc).unwrap().is_empty());
    }

    #[test]
    fn end_to_end_via_minimal_fixture() {
        let mut tmp = NamedTempFile::new().unwrap();
        tmp.write_all(&minimal_docx_bytes()).unwrap();
        tmp.flush().unwrap();
        let doc = Doc::load(tmp.path()).unwrap();
        let styles = list_paragraph_styles(&doc).unwrap();
        assert_eq!(styles, vec!["Title", "Body"]);
    }

    #[test]
    fn rejects_malformed_xml() {
        let xml = "<w:styles><w:style w:type=\"paragraph\" w:styleId=\"X\"></w:styles>";
        // Missing close — quick-xml in default mode tolerates this; ensure no panic
        // and either Ok with the partial style or a typed error.
        let result = parse_paragraph_styles(xml.as_bytes());
        match result {
            Ok(_) | Err(DocxaiError::Generic(_)) => {}
            Err(other) => panic!("unexpected error kind: {other:?}"),
        }
    }
}
