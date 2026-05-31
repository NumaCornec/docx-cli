use std::io::Write;
use std::process::Stdio;
use std::sync::OnceLock;

use crate::error::DocxaiError;

static PANDOC_PATH: OnceLock<Result<String, String>> = OnceLock::new();

fn detect_pandoc() -> Result<String, String> {
    let output = std::process::Command::new("pandoc")
        .arg("--version")
        .stdout(Stdio::piped())
        .stderr(Stdio::null())
        .output()
        .map_err(|e| {
            format!(
                "pandoc not found. Install with:\n  macOS: brew install pandoc\n  Linux: sudo apt install pandoc\n  Windows: winget install pandoc\nError: {e}"
            )
        })?;

    if !output.status.success() {
        return Err("pandoc --version failed".into());
    }

    let stdout = String::from_utf8_lossy(&output.stdout);
    let first_line = stdout.lines().next().unwrap_or("pandoc");
    let version = first_line.strip_prefix("pandoc ").unwrap_or("unknown");

    Ok(format!("pandoc at version {version}"))
}

pub fn require_pandoc() -> Result<(), DocxaiError> {
    let result = PANDOC_PATH.get_or_init(detect_pandoc);
    match result {
        Ok(_) => Ok(()),
        Err(msg) => Err(DocxaiError::MissingDependency(msg.clone())),
    }
}

pub fn latex_to_omml(latex: &str) -> Result<String, DocxaiError> {
    let omml = latex_to_omml_raw(latex)?;
    extract_omath_para(&omml)
}

pub fn latex_to_inline_omath(latex: &str) -> Result<String, DocxaiError> {
    let omml = latex_to_omml_raw(latex)?;

    if let Some(start) = find_tag_start(&omml, b"<m:oMath") {
        if let Some(end) = find_closing_tag(&omml, b"</m:oMath>") {
            let omath = &omml[start..end + "</m:oMath>".len()];
            return Ok(omath.to_string());
        }
    }

    Err(DocxaiError::Generic(
        "pandoc output contains no inline math element".into(),
    ))
}

fn latex_to_omml_raw(latex: &str) -> Result<String, DocxaiError> {
    require_pandoc()?;

    let mut child = std::process::Command::new("pandoc")
        .args(["-f", "latex", "-t", "docx"])
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .map_err(|e| DocxaiError::Generic(format!("failed to spawn pandoc: {e}")))?;

    if let Some(stdin) = child.stdin.take() {
        let mut stdin = stdin;
        stdin
            .write_all(latex.as_bytes())
            .map_err(|e| DocxaiError::Generic(format!("failed to write to pandoc: {e}")))?;
    }

    let output = child
        .wait_with_output()
        .map_err(|e| DocxaiError::Generic(format!("pandoc execution failed: {e}")))?;

    if !output.status.success() {
        let stderr = String::from_utf8_lossy(&output.stderr);
        return Err(DocxaiError::InvalidArgument(format!(
            "pandoc could not convert LaTeX: {stderr}"
        )));
    }

    let docx_bytes = output.stdout;
    extract_omath_from_docx(&docx_bytes)
}

fn extract_omath_from_docx(docx_bytes: &[u8]) -> Result<String, DocxaiError> {
    let reader = std::io::Cursor::new(docx_bytes);
    let mut archive = zip::ZipArchive::new(reader)
        .map_err(|e| DocxaiError::Generic(format!("pandoc output is not valid docx: {e}")))?;

    let document_xml = archive
        .by_name("word/document.xml")
        .map_err(|e| DocxaiError::Generic(format!("no word/document.xml in pandoc output: {e}")))?;

    let mut xml_bytes = Vec::new();
    use std::io::Read;
    let mut doc_reader = document_xml;
    doc_reader
        .read_to_end(&mut xml_bytes)
        .map_err(|e| DocxaiError::Generic(format!("read document.xml: {e}")))?;

    let xml = String::from_utf8(xml_bytes)
        .map_err(|e| DocxaiError::Generic(format!("document.xml is not utf8: {e}")))?;

    extract_omath_para(&xml)
}

fn extract_omath_para(xml: &str) -> Result<String, DocxaiError> {
    // Look for <m:oMathPara ...>...</m:oMathPara> (display math)
    if let Some(start) = find_tag_start(xml, b"<m:oMathPara") {
        if let Some(end) = find_closing_tag(xml, b"</m:oMathPara>") {
            return Ok(xml[start..end + "</m:oMathPara>".len()].to_string());
        }
    }

    // Fallback: look for standalone <m:oMath ...>...</m:oMath> (inline math)
    // Search for <m:oMath but NOT <m:oMathPara
    let standalone = find_standalone_omath(xml);
    if let Some((start, end)) = standalone {
        let omath = &xml[start..end];
        return Ok(format!(
            "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMathParaPr><m:jc m:val=\"center\"/></m:oMathParaPr>{omath}</m:oMathPara>"
        ));
    }

    Err(DocxaiError::Generic(
        "pandoc output contains no math element".into(),
    ))
}

fn find_standalone_omath(xml: &str) -> Option<(usize, usize)> {
    let bytes = xml.as_bytes();
    let mut pos = 0;
    while pos < bytes.len() {
        if let Some(offset) = find_tag_start(&xml[pos..], b"<m:oMath") {
            let abs = pos + offset;
            // Check this isn't <m:oMathPara
            let after = &xml[abs + b"<m:oMath".len()..];
            let next_char = after.bytes().next();
            match next_char {
                Some(b'>') | Some(b' ') => {
                    // This is <m:oMath> or <m:oMath ...> — standalone
                    if let Some(end) = find_tag_start(&xml[abs..], b"</m:oMath>") {
                        return Some((abs, abs + end + "</m:oMath>".len()));
                    }
                }
                _ => {
                    // This is <m:oMathPara, skip it
                    pos = abs + 1;
                    continue;
                }
            }
        }
        break;
    }
    None
}

fn find_tag_start(xml: &str, tag: &[u8]) -> Option<usize> {
    let xml_bytes = xml.as_bytes();
    xml_bytes.windows(tag.len()).position(|w| w == tag)
}

fn find_closing_tag(xml: &str, tag: &[u8]) -> Option<usize> {
    find_tag_start(xml, tag)
}

pub fn omml_to_latex(omml: &str) -> Result<String, DocxaiError> {
    require_pandoc()?;

    // Build a minimal docx with the OMML to reverse-convert
    let docx = build_docx_with_omml(omml);

    let mut child = std::process::Command::new("pandoc")
        .args(["-f", "docx", "-t", "latex"])
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .map_err(|e| DocxaiError::Generic(format!("failed to spawn pandoc: {e}")))?;

    if let Some(stdin) = child.stdin.take() {
        let mut stdin = stdin;
        stdin
            .write_all(&docx)
            .map_err(|e| DocxaiError::Generic(format!("failed to write to pandoc: {e}")))?;
    }

    let output = child
        .wait_with_output()
        .map_err(|e| DocxaiError::Generic(format!("pandoc execution failed: {e}")))?;

    if !output.status.success() {
        let stderr = String::from_utf8_lossy(&output.stderr);
        return Err(DocxaiError::Generic(format!(
            "pandoc reverse conversion failed: {stderr}"
        )));
    }

    let latex = String::from_utf8_lossy(&output.stdout);
    let cleaned = latex
        .trim()
        .trim_start_matches('$')
        .trim_end_matches('$')
        .trim()
        .to_string();

    Ok(cleaned)
}

fn build_docx_with_omml(omml: &str) -> Vec<u8> {
    let document_xml = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
         <w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"\
         xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">\
         <w:body>{omml}</w:body></w:document>"
    );

    let content_types = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
        <Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\
        <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\
        <Default Extension=\"xml\" ContentType=\"application/xml\"/>\
        <Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>\
        </Types>";

    let rels = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
        <Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\
        <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>\
        </Relationships>";

    let buf: Vec<u8> = Vec::new();
    let w = std::io::Cursor::new(buf);
    let mut zip = zip::ZipWriter::new(w);

    let options = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(content_types.as_bytes()).unwrap();

    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(rels.as_bytes()).unwrap();

    zip.start_file("word/document.xml", options).unwrap();
    zip.write_all(document_xml.as_bytes()).unwrap();

    zip.finish().unwrap().into_inner()
}

pub fn add_equation(
    doc: &mut crate::doc::Doc,
    latex: &str,
    after: Option<&str>,
    before: Option<&str>,
) -> Result<serde_json::Value, DocxaiError> {
    let omml = latex_to_omml(latex)?;

    let map = crate::mutate::index_body_spans(&doc.parts.document_xml)?;

    let insert_pos = crate::mutate::determine_insert_pos_from(&map, after, before)?;

    let paragraph = format!("<w:p><w:pPr><w:pStyle w:val=\"\"/></w:pPr>{omml}</w:p>");

    let new_bytes = {
        let xml = std::str::from_utf8(&doc.parts.document_xml)
            .map_err(|e| DocxaiError::Generic(format!("utf8: {e}")))?;
        let mut out = String::with_capacity(xml.len() + paragraph.len());
        out.push_str(&xml[..insert_pos]);
        out.push_str(&paragraph);
        out.push_str(&xml[insert_pos..]);
        out.into_bytes()
    };
    doc.parts.document_xml = new_bytes;

    let map2 = crate::mutate::index_body_spans(&doc.parts.document_xml)?;
    let eq_span = map2
        .spans
        .iter()
        .find(|s| s.kind == 'e')
        .ok_or_else(|| DocxaiError::Generic("equation not found after insert".into()))?;

    // Count how many equations are after this one to find its index
    let eq_index = map2
        .spans
        .iter()
        .filter(|s| s.kind == 'e' && s.end <= eq_span.end)
        .count() as u32;

    Ok(serde_json::json!({
        "status": "ok",
        "action": "add",
        "ref": format!("@e{eq_index}"),
        "kind": "equation"
    }))
}

pub fn set_equation(
    doc: &mut crate::doc::Doc,
    reference: &str,
    latex: &str,
) -> Result<serde_json::Value, DocxaiError> {
    let parsed = crate::refs::Ref::parse(reference)?;
    match parsed {
        crate::refs::Ref::Equation(_) => {}
        _ => {
            return Err(DocxaiError::InvalidArgument(format!(
                "{reference} is not an equation ref"
            )));
        }
    };

    let omml = latex_to_omml(latex)?;

    let map = crate::mutate::index_body_spans(&doc.parts.document_xml)?;
    let span = crate::mutate::find_span(&map.spans, &parsed)?;

    let new_bytes = {
        let xml = std::str::from_utf8(&doc.parts.document_xml)
            .map_err(|e| DocxaiError::Generic(format!("utf8: {e}")))?;
        let paragraph = format!("<w:p><w:pPr><w:pStyle w:val=\"\"/></w:pPr>{omml}</w:p>");
        let mut out = String::with_capacity(xml.len() + paragraph.len());
        out.push_str(&xml[..span.start]);
        out.push_str(&paragraph);
        out.push_str(&xml[span.end..]);
        out.into_bytes()
    };
    doc.parts.document_xml = new_bytes;

    Ok(serde_json::json!({
        "status": "ok",
        "action": "set",
        "ref": reference,
        "kind": "equation",
        "changed": ["latex"]
    }))
}

pub fn delete_equation(
    doc: &mut crate::doc::Doc,
    reference: &str,
) -> Result<serde_json::Value, DocxaiError> {
    let parsed = crate::refs::Ref::parse(reference)?;
    match parsed {
        crate::refs::Ref::Equation(_) => {}
        _ => {
            return Err(DocxaiError::InvalidArgument(format!(
                "{reference} is not an equation ref"
            )));
        }
    };

    let map = crate::mutate::index_body_spans(&doc.parts.document_xml)?;
    let span = crate::mutate::find_span(&map.spans, &parsed)?;

    let new_bytes = {
        let xml = std::str::from_utf8(&doc.parts.document_xml)
            .map_err(|e| DocxaiError::Generic(format!("utf8: {e}")))?;
        let mut out = String::with_capacity(xml.len());
        out.push_str(&xml[..span.start]);
        out.push_str(&xml[span.end..]);
        out.into_bytes()
    };
    doc.parts.document_xml = new_bytes;

    Ok(serde_json::json!({
        "status": "ok",
        "action": "delete",
        "ref": reference,
        "kind": "equation"
    }))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn detect_pandoc_caches_result() {
        let r1 = PANDOC_PATH.get_or_init(detect_pandoc);
        let r2 = PANDOC_PATH.get_or_init(detect_pandoc);
        assert!(std::ptr::eq(r1, r2), "should return same reference");
    }

    #[test]
    fn extract_omath_para_finds_display_math() {
        let xml = r#"<?xml version="1.0"?><w:document><w:body><m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></m:oMathPara></w:body></w:document>"#;
        let result = extract_omath_para(xml).unwrap();
        assert!(result.contains("<m:oMathPara"), "should contain oMathPara");
        assert!(
            result.contains("</m:oMathPara>"),
            "should contain closing tag"
        );
    }

    #[test]
    fn extract_omath_para_wraps_inline_math() {
        let xml = r#"<?xml version="1.0"?><w:document><w:body><m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>y</m:t></m:r></m:oMath></w:body></w:document>"#;
        let result = extract_omath_para(xml).unwrap();
        assert!(result.contains("<m:oMathPara"), "inline should be wrapped");
        assert!(result.contains("<m:oMath"), "should contain oMath");
    }

    #[test]
    fn extract_omath_para_errors_on_no_math() {
        let xml =
            r#"<?xml version="1.0"?><w:document><w:body><w:p>hello</w:p></w:body></w:document>"#;
        assert!(extract_omath_para(xml).is_err());
    }

    #[test]
    fn build_docx_with_omml_produces_valid_zip() {
        let omml = "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath><m:r><m:t>E</m:t></m:r></m:oMath></m:oMathPara>";
        let docx = build_docx_with_omml(omml);
        let reader = std::io::Cursor::new(&docx);
        let archive = zip::ZipArchive::new(reader);
        assert!(archive.is_ok(), "should produce valid zip");
        let mut ar = archive.unwrap();
        let mut doc = ar.by_name("word/document.xml").unwrap();
        let mut bytes = Vec::new();
        std::io::Read::read_to_end(&mut doc, &mut bytes).unwrap();
        let xml = String::from_utf8(bytes).unwrap();
        assert!(xml.contains("oMathPara"));
    }

    #[test]
    fn require_pandoc_does_not_panic() {
        let _ = require_pandoc();
    }
}
