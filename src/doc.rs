//! In-memory representation of a `.docx` archive.
//!
//! A `.docx` file is a ZIP container of XML parts (see ECMA-376 / OOXML).
//! `Doc` opens that container, splits out the well-known parts the CLI
//! mutates (`word/document.xml`, `word/styles.xml`, the document relations,
//! and core metadata) and keeps every other part as raw bytes so it can be
//! written back byte-for-byte. PRD §6.3 forbids reconstructing untouched
//! XML from scratch — this is the structure that makes that possible.

use std::collections::BTreeMap;
use std::fs::File;
use std::io::Read;
use std::path::{Path, PathBuf};

use crate::error::DocxaiError;

/// Path of `[Content_Types].xml` inside a docx ZIP.
pub const CONTENT_TYPES_PATH: &str = "[Content_Types].xml";
/// Path of the main document part.
pub const DOCUMENT_PATH: &str = "word/document.xml";
/// Path of the style definitions part.
pub const STYLES_PATH: &str = "word/styles.xml";
/// Path of the document-level relationships part.
pub const DOCUMENT_RELS_PATH: &str = "word/_rels/document.xml.rels";
/// Path of the core properties (title, author, …) part.
pub const CORE_PROPS_PATH: &str = "docProps/core.xml";

/// Well-known parts plus a verbatim copy of every other ZIP entry.
///
/// `others` is keyed by archive path so save logic can iterate in a
/// deterministic order and the byte payload survives untouched.
#[derive(Debug, Clone)]
pub struct Parts {
    pub content_types: Vec<u8>,
    pub document_xml: Vec<u8>,
    pub styles_xml: Option<Vec<u8>>,
    pub document_rels: Option<Vec<u8>>,
    pub core_props: Option<Vec<u8>>,
    /// Every other entry in the ZIP, preserved verbatim.
    pub others: BTreeMap<String, Vec<u8>>,
}

/// A loaded `.docx` document.
#[derive(Debug, Clone)]
pub struct Doc {
    /// Path the document was loaded from. `save_in_place()` (later) writes back here.
    pub path: PathBuf,
    pub parts: Parts,
}

impl Doc {
    /// Open a `.docx` from disk, splitting well-known XML parts from the rest.
    ///
    /// Errors map to PRD §10.1 exit codes:
    /// * I/O / not-a-zip / missing required parts → [`DocxaiError::Generic`] (exit 1).
    pub fn load(path: impl AsRef<Path>) -> Result<Self, DocxaiError> {
        let path = path.as_ref();
        let file = File::open(path)
            .map_err(|e| DocxaiError::Generic(format!("cannot open {}: {e}", path.display())))?;
        let mut archive = zip::ZipArchive::new(file)
            .map_err(|e| DocxaiError::Generic(format!("not a valid zip: {e}")))?;

        let mut content_types: Option<Vec<u8>> = None;
        let mut document_xml: Option<Vec<u8>> = None;
        let mut styles_xml: Option<Vec<u8>> = None;
        let mut document_rels: Option<Vec<u8>> = None;
        let mut core_props: Option<Vec<u8>> = None;
        let mut others: BTreeMap<String, Vec<u8>> = BTreeMap::new();

        for i in 0..archive.len() {
            let mut entry = archive
                .by_index(i)
                .map_err(|e| DocxaiError::Generic(format!("zip entry {i}: {e}")))?;
            if !entry.is_file() {
                continue;
            }
            let name = entry.name().to_owned();
            let mut buf = Vec::with_capacity(entry.size() as usize);
            entry
                .read_to_end(&mut buf)
                .map_err(|e| DocxaiError::Generic(format!("read {name}: {e}")))?;

            match name.as_str() {
                CONTENT_TYPES_PATH => content_types = Some(buf),
                DOCUMENT_PATH => document_xml = Some(buf),
                STYLES_PATH => styles_xml = Some(buf),
                DOCUMENT_RELS_PATH => document_rels = Some(buf),
                CORE_PROPS_PATH => core_props = Some(buf),
                _ => {
                    others.insert(name, buf);
                }
            }
        }

        let content_types = content_types.ok_or_else(|| {
            DocxaiError::Generic(format!(
                "{} not a docx: missing [Content_Types].xml",
                path.display()
            ))
        })?;
        let document_xml = document_xml.ok_or_else(|| {
            DocxaiError::Generic(format!(
                "{} not a docx: missing word/document.xml",
                path.display()
            ))
        })?;

        Ok(Self {
            path: path.to_path_buf(),
            parts: Parts {
                content_types,
                document_xml,
                styles_xml,
                document_rels,
                core_props,
                others,
            },
        })
    }
}

#[cfg(test)]
pub(crate) mod test_fixture {
    //! Helpers that synthesise minimal docx archives in memory so tests do
    //! not need committed binary fixtures (PRD #3 corpus is deferred).

    use std::io::{Cursor, Write};

    use zip::write::SimpleFileOptions;

    /// `[Content_Types].xml` declaring the parts a minimal docx needs.
    pub const CONTENT_TYPES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>"#;

    pub const ROOT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
</Relationships>"#;

    pub const DOCUMENT_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>Hello</w:t></w:r></w:p>
  </w:body>
</w:document>"#;

    pub const STYLES_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/></w:style>
  <w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body"/></w:style>
</w:styles>"#;

    pub const DOCUMENT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>"#;

    pub const CORE_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:title>Sample</dc:title>
  <dc:creator>Tester</dc:creator>
</cp:coreProperties>"#;

    /// Build a minimal but ECMA-376-shaped docx in memory.
    pub fn minimal_docx_bytes() -> Vec<u8> {
        let mut buf = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut buf));
            let opts = SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Deflated);
            for (name, body) in [
                ("[Content_Types].xml", CONTENT_TYPES),
                ("_rels/.rels", ROOT_RELS),
                ("word/document.xml", DOCUMENT_XML),
                ("word/styles.xml", STYLES_XML),
                ("word/_rels/document.xml.rels", DOCUMENT_RELS),
                ("docProps/core.xml", CORE_XML),
            ] {
                zip.start_file(name, opts).unwrap();
                zip.write_all(body.as_bytes()).unwrap();
            }
            zip.finish().unwrap();
        }
        buf
    }

    /// Build a docx whose only required-but-missing part is `[Content_Types].xml`.
    pub fn missing_content_types_bytes() -> Vec<u8> {
        let mut buf = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut buf));
            let opts = SimpleFileOptions::default();
            zip.start_file("word/document.xml", opts).unwrap();
            zip.write_all(DOCUMENT_XML.as_bytes()).unwrap();
            zip.finish().unwrap();
        }
        buf
    }

    /// Build a docx with `[Content_Types].xml` but no `word/document.xml`.
    pub fn missing_document_bytes() -> Vec<u8> {
        let mut buf = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut buf));
            let opts = SimpleFileOptions::default();
            zip.start_file("[Content_Types].xml", opts).unwrap();
            zip.write_all(CONTENT_TYPES.as_bytes()).unwrap();
            zip.finish().unwrap();
        }
        buf
    }
}

#[cfg(test)]
mod tests {
    use std::io::Write;

    use tempfile::NamedTempFile;

    use super::test_fixture::{
        minimal_docx_bytes, missing_content_types_bytes, missing_document_bytes, CORE_XML,
        DOCUMENT_RELS, DOCUMENT_XML, STYLES_XML,
    };
    use super::*;

    fn write_tmp(bytes: &[u8]) -> NamedTempFile {
        let mut tmp = NamedTempFile::new().unwrap();
        tmp.write_all(bytes).unwrap();
        tmp.flush().unwrap();
        tmp
    }

    #[test]
    fn load_minimal_docx_extracts_well_known_parts() {
        let tmp = write_tmp(&minimal_docx_bytes());
        let doc = Doc::load(tmp.path()).expect("minimal docx should load");

        assert_eq!(doc.path, tmp.path());
        assert_eq!(doc.parts.document_xml, DOCUMENT_XML.as_bytes());
        assert_eq!(doc.parts.styles_xml.as_deref(), Some(STYLES_XML.as_bytes()));
        assert_eq!(
            doc.parts.document_rels.as_deref(),
            Some(DOCUMENT_RELS.as_bytes())
        );
        assert_eq!(doc.parts.core_props.as_deref(), Some(CORE_XML.as_bytes()));
        // _rels/.rels is not a "well-known mutated part" — it lives in `others`.
        assert!(doc.parts.others.contains_key("_rels/.rels"));
        // No accidental capture of well-known parts in the catch-all.
        for known in [
            CONTENT_TYPES_PATH,
            DOCUMENT_PATH,
            STYLES_PATH,
            DOCUMENT_RELS_PATH,
            CORE_PROPS_PATH,
        ] {
            assert!(
                !doc.parts.others.contains_key(known),
                "{known} should be promoted out of `others`"
            );
        }
    }

    #[test]
    fn load_rejects_missing_content_types() {
        let tmp = write_tmp(&missing_content_types_bytes());
        let err = Doc::load(tmp.path()).expect_err("should fail without content types");
        let msg = err.to_string();
        assert!(msg.contains("[Content_Types].xml"), "msg was: {msg}");
        assert_eq!(err.exit_code(), crate::error::ExitCode::Generic);
    }

    #[test]
    fn load_rejects_missing_document_part() {
        let tmp = write_tmp(&missing_document_bytes());
        let err = Doc::load(tmp.path()).expect_err("should fail without document.xml");
        let msg = err.to_string();
        assert!(msg.contains("word/document.xml"), "msg was: {msg}");
    }

    #[test]
    fn load_rejects_non_zip() {
        let tmp = write_tmp(b"this is plainly not a zip archive");
        let err = Doc::load(tmp.path()).expect_err("non-zip should fail");
        assert!(err.to_string().to_lowercase().contains("zip"));
    }

    #[test]
    fn load_rejects_missing_file() {
        let err = Doc::load("/nonexistent/path/does-not-exist.docx")
            .expect_err("missing file should fail");
        assert!(err.to_string().contains("cannot open"));
    }
}
