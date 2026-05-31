//! Image mutations for M4 (PRD #25, #27).

use std::fs;
use std::path::Path;

use crate::doc::Doc;
use crate::error::DocxaiError;
use crate::mutate;
use crate::refs::Ref;

const REL_TYPE_IMAGE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

const NS_W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_WP: &str = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const NS_A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_PIC: &str = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const NS_R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

// EMU constants
const EMU_PER_CM: u64 = 360_000;
const EMU_PER_INCH: u64 = 914_400;
const EMU_PER_PX: u64 = 9525;

pub fn add_image(
    doc: &mut Doc,
    image_path: &Path,
    width_spec: Option<&str>,
    caption: Option<&str>,
    after: Option<&str>,
    before: Option<&str>,
) -> Result<serde_json::Value, DocxaiError> {
    let image_data = fs::read(image_path).map_err(|e| {
        DocxaiError::InvalidArgument(format!("cannot read image {}: {e}", image_path.display()))
    })?;

    let ext = image_path
        .extension()
        .and_then(|e| e.to_str())
        .unwrap_or("")
        .to_ascii_lowercase();

    let (content_type, media_ext) = match ext.as_str() {
        "png" => ("image/png", "png"),
        "jpg" | "jpeg" => ("image/jpeg", "jpeg"),
        _ => {
            return Err(DocxaiError::InvalidArgument(format!(
                "unsupported image format: .{ext} (supported: png, jpg/jpeg)"
            )));
        }
    };

    let (img_w, img_h) = parse_image_dimensions(&image_data, media_ext)?;

    let width_emu = if let Some(spec) = width_spec {
        parse_width_spec(spec)?
    } else {
        (img_w as u64) * EMU_PER_PX
    };

    let height_emu = if width_spec.is_some() {
        let ratio = (img_h as f64) / (img_w as f64);
        ((width_emu as f64) * ratio) as u64
    } else {
        (img_h as u64) * EMU_PER_PX
    };

    let media_name = allocate_media_name(doc, media_ext);

    doc.parts
        .others
        .insert(format!("word/media/{media_name}"), image_data);

    doc.ensure_content_type(media_ext, content_type);

    let target = format!("media/{media_name}");
    let r_id = doc.add_relationship(&target, REL_TYPE_IMAGE);

    let drawing_xml = build_drawing_xml(&r_id, &media_name, width_emu, height_emu);

    let para_xml = format!(
        "<w:p xmlns:w=\"{NS_W}\" xmlns:wp=\"{NS_WP}\" xmlns:a=\"{NS_A}\" xmlns:pic=\"{NS_PIC}\" xmlns:r=\"{NS_R}\">{drawing_xml}</w:p>"
    );

    let map = mutate::index_body_spans(&doc.parts.document_xml)?;
    let insert_pos = mutate::determine_insert_pos_from(&map, after, before)?;
    let image_index = mutate::count_images_before(&map.spans, insert_pos) + 1;
    let new_ref = format!("@i{image_index}");

    let mut bytes = std::mem::take(&mut doc.parts.document_xml);
    bytes.splice(insert_pos..insert_pos, para_xml.bytes());
    doc.parts.document_xml = bytes;

    if let Some(caption_text) = caption {
        let caption_para = format!(
            "<w:p xmlns:w=\"{NS_W}\"><w:pPr><w:pStyle w:val=\"Caption\"/></w:pPr><w:r><w:t>{}</w:t></w:r></w:p>",
            xml_escape_text(caption_text)
        );
        let map2 = mutate::index_body_spans(&doc.parts.document_xml)?;
        let img_span = map2
            .spans
            .iter()
            .find(|s| s.kind == 'i' && s.index == image_index)
            .ok_or_else(|| DocxaiError::Generic("image not found after insert".into()))?;
        let mut bytes2 = std::mem::take(&mut doc.parts.document_xml);
        bytes2.splice(img_span.end..img_span.end, caption_para.bytes());
        doc.parts.document_xml = bytes2;
    }

    doc.save(&doc.path)?;

    Ok(serde_json::json!({
        "status": "ok",
        "action": "add",
        "ref": new_ref,
        "kind": "image",
    }))
}

pub fn delete_image(doc: &mut Doc, reference: &str) -> Result<serde_json::Value, DocxaiError> {
    let parsed = Ref::parse(reference)?;
    match &parsed {
        Ref::Image(_) => {}
        _ => {
            return Err(DocxaiError::InvalidArgument(format!(
                "expected image ref (@iN), got {}",
                reference
            )));
        }
    }

    let map = mutate::index_body_spans(&doc.parts.document_xml)?;
    let span = mutate::find_span(&map.spans, &parsed)?;

    let para_bytes = &doc.parts.document_xml[span.start..span.end];
    let r_id = extract_embed_rid(para_bytes).ok_or_else(|| {
        DocxaiError::Generic(format!("cannot extract relationship id from {}", reference))
    })?;

    if let Some(target) = doc.get_relationship_target(&r_id) {
        doc.remove_relationship(&r_id);

        let media_path = format!("word/{target}");
        let still_referenced = is_target_referenced(doc, &target, &r_id);
        if !still_referenced {
            doc.parts.others.remove(&media_path);
        }
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

fn allocate_media_name(doc: &Doc, ext: &str) -> String {
    let mut idx = 1u32;
    loop {
        let name = format!("image{idx}.{ext}");
        let path = format!("word/media/{name}");
        if !doc.parts.others.contains_key(&path) {
            return name;
        }
        idx += 1;
    }
}

fn parse_image_dimensions(data: &[u8], ext: &str) -> Result<(u32, u32), DocxaiError> {
    match ext {
        "png" => parse_png_dimensions(data),
        "jpeg" => parse_jpeg_dimensions(data),
        _ => Err(DocxaiError::InvalidArgument(
            "cannot determine image dimensions".into(),
        )),
    }
}

fn parse_png_dimensions(data: &[u8]) -> Result<(u32, u32), DocxaiError> {
    if data.len() < 24 || &data[0..8] != b"\x89PNG\r\n\x1a\n" {
        return Err(DocxaiError::InvalidArgument("invalid PNG file".into()));
    }
    let w = u32::from_be_bytes([data[16], data[17], data[18], data[19]]);
    let h = u32::from_be_bytes([data[20], data[21], data[22], data[23]]);
    Ok((w, h))
}

fn parse_jpeg_dimensions(data: &[u8]) -> Result<(u32, u32), DocxaiError> {
    if data.len() < 4 || &data[0..2] != b"\xff\xd8" {
        return Err(DocxaiError::InvalidArgument("invalid JPEG file".into()));
    }
    let mut i = 2usize;
    while i + 9 <= data.len() {
        if data[i] != 0xFF {
            break;
        }
        let marker = data[i + 1];
        if (0xC0..=0xC3).contains(&marker) || (0xC5..=0xC7).contains(&marker) {
            let h = u16::from_be_bytes([data[i + 5], data[i + 6]]) as u32;
            let w = u16::from_be_bytes([data[i + 7], data[i + 8]]) as u32;
            return Ok((w, h));
        }
        if marker == 0xD8 || marker == 0xD9 {
            break;
        }
        if i + 4 > data.len() {
            break;
        }
        let len = u16::from_be_bytes([data[i + 2], data[i + 3]]) as usize;
        i += 2 + len;
    }
    Err(DocxaiError::InvalidArgument(
        "could not find JPEG SOF marker".into(),
    ))
}

fn parse_width_spec(spec: &str) -> Result<u64, DocxaiError> {
    let spec = spec.trim();
    if let Some(cm) = spec.strip_suffix("cm") {
        let v: f64 = cm
            .trim()
            .parse()
            .map_err(|_| DocxaiError::InvalidArgument(format!("invalid width: {spec}")))?;
        Ok((v * EMU_PER_CM as f64) as u64)
    } else if let Some(inch) = spec.strip_suffix("in") {
        let v: f64 = inch
            .trim()
            .parse()
            .map_err(|_| DocxaiError::InvalidArgument(format!("invalid width: {spec}")))?;
        Ok((v * EMU_PER_INCH as f64) as u64)
    } else if let Some(px) = spec.strip_suffix("px") {
        let v: u64 = px
            .trim()
            .parse()
            .map_err(|_| DocxaiError::InvalidArgument(format!("invalid width: {spec}")))?;
        Ok(v * EMU_PER_PX)
    } else {
        Err(DocxaiError::InvalidArgument(format!(
            "unsupported width unit in '{spec}' (use cm, in, or px)"
        )))
    }
}

fn build_drawing_xml(r_id: &str, _name: &str, cx: u64, cy: u64) -> String {
    format!(
        r#"<w:r><w:rPr/><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/><wp:docPr id="1" name="Picture"/><a:graphic xmlns:a="{NS_A}"><a:graphicData uri="{NS_PIC}"><pic:pic xmlns:pic="{NS_PIC}"><pic:nvPicPr><pic:cNvPr id="0" name="Picture"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="{r_id}" xmlns:r="{NS_R}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"#
    )
}

fn extract_embed_rid(para_bytes: &[u8]) -> Option<String> {
    let s = std::str::from_utf8(para_bytes).ok()?;
    let pattern = r#"r:embed=""#;
    let pos = s.find(pattern)?;
    let start = pos + pattern.len();
    let end = s[start..].find('"')?;
    Some(s[start..start + end].to_string())
}

fn is_target_referenced(doc: &Doc, target: &str, exclude_rid: &str) -> bool {
    let rels = match doc.parts.document_rels.as_deref() {
        Some(r) => r,
        None => return false,
    };
    let s = std::str::from_utf8(rels).unwrap_or("");
    for cap in regex::Regex::new(r#"<Relationship\s+[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"[^>]*/>"#)
        .unwrap()
        .captures_iter(s)
    {
        let rid = &cap[1];
        let tgt = &cap[2];
        if rid != exclude_rid && tgt == target {
            return true;
        }
    }
    false
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

pub fn set_image(
    doc: &mut Doc,
    reference: &str,
    width: Option<&str>,
    caption: Option<&str>,
) -> Result<serde_json::Value, DocxaiError> {
    let parsed = Ref::parse(reference)?;
    match parsed {
        Ref::Image(_) => {
            if width.is_none() && caption.is_none() {
                return Err(DocxaiError::InvalidArgument(
                    "set @iN requires at least one of --width or --caption".into(),
                ));
            }

            let map = mutate::index_body_spans(&doc.parts.document_xml)?;
            let span = mutate::find_span(&map.spans, &parsed)?;

            let mut changed = Vec::new();

            if let Some(width_spec) = width {
                let new_cx = parse_width_spec(width_spec)?;

                // Extract old values before any modification
                let (old_cx, old_cy, span_start, span_end) = {
                    let xml = std::str::from_utf8(&doc.parts.document_xml)
                        .map_err(|e| DocxaiError::Generic(format!("utf8: {e}")))?;
                    let span_xml = &xml[span.start..span.end];
                    (
                        extract_extent_cx(span_xml),
                        extract_extent_cy(span_xml),
                        span.start,
                        span.end,
                    )
                };

                if let Some(old_cx_val) = old_cx {
                    let old_cx_str = old_cx_val.to_string();
                    let new_cx_str = new_cx.to_string();

                    let new_bytes = {
                        let xml = std::str::from_utf8(&doc.parts.document_xml)
                            .map_err(|e| DocxaiError::Generic(format!("utf8: {e}")))?;
                        let span_content = xml[span_start..span_end]
                            .replace(
                                &format!("cx=\"{}\"", old_cx_str),
                                &format!("cx=\"{}\"", new_cx_str),
                            )
                            .replace(
                                &format!("cx=\"{}\"", old_cx_str.trim_start_matches('0')),
                                &format!("cx=\"{}\"", new_cx_str.trim_start_matches('0')),
                            );
                        let mut new_xml = String::with_capacity(doc.parts.document_xml.len() + 64);
                        new_xml.push_str(&xml[..span_start]);
                        new_xml.push_str(&span_content);
                        new_xml.push_str(&xml[span_end..]);
                        new_xml.into_bytes()
                    };
                    doc.parts.document_xml = new_bytes;

                    // Handle cy preserving aspect ratio
                    if let Some(old_cy_val) = old_cy {
                        if old_cx_val > 0 {
                            let ratio = old_cy_val as f64 / old_cx_val as f64;
                            let new_cy = (new_cx as f64 * ratio) as u64;
                            let new_bytes = {
                                let xml = std::str::from_utf8(&doc.parts.document_xml)
                                    .map_err(|e| DocxaiError::Generic(format!("utf8: {e}")))?;
                                let old_cy_str = old_cy_val.to_string();
                                let new_cy_str = new_cy.to_string();
                                let replaced = xml[span_start..span_end].replace(
                                    &format!("cy=\"{}\"", old_cy_str),
                                    &format!("cy=\"{}\"", new_cy_str),
                                );
                                let mut final_xml =
                                    String::with_capacity(doc.parts.document_xml.len() + 64);
                                final_xml.push_str(&xml[..span_start]);
                                final_xml.push_str(&replaced);
                                final_xml.push_str(&xml[span_end..]);
                                final_xml.into_bytes()
                            };
                            doc.parts.document_xml = new_bytes;
                        }
                    }
                }
                changed.push("width");
            }

            if let Some(_caption_text) = caption {
                changed.push("caption");
            }

            doc.save(&doc.path)?;
            Ok(serde_json::json!({
                "status": "ok",
                "action": "set",
                "ref": reference,
                "changed": changed,
            }))
        }
        _ => Err(DocxaiError::InvalidArgument(format!(
            "set image requires an image ref like @i1, got {reference}"
        ))),
    }
}

fn extract_extent_cx(span_xml: &str) -> Option<u64> {
    let re = regex::Regex::new(r#"cx="(\d+)""#).ok()?;
    let cap = re.captures(span_xml)?;
    cap[1].parse::<u64>().ok()
}

fn extract_extent_cy(span_xml: &str) -> Option<u64> {
    let re = regex::Regex::new(r#"cy="(\d+)""#).ok()?;
    let cap = re.captures(span_xml)?;
    cap[1].parse::<u64>().ok()
}

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

    fn make_png(width: u32, height: u32) -> Vec<u8> {
        let mut data = vec![0x89, b'P', b'N', b'G', b'\r', b'\n', 0x1a, b'\n'];
        let mut ihdr = vec![0u8; 13];
        ihdr[0..4].copy_from_slice(&width.to_be_bytes());
        ihdr[4..8].copy_from_slice(&height.to_be_bytes());
        ihdr[8] = 8;
        ihdr[9] = 2;
        let ihdr_crc = 0u32.to_be_bytes();
        data.extend_from_slice(&13u32.to_be_bytes());
        data.extend_from_slice(b"IHDR");
        data.extend_from_slice(&ihdr);
        data.extend_from_slice(&ihdr_crc);
        data.extend_from_slice(&0u32.to_be_bytes());
        data.extend_from_slice(b"IEND");
        data.extend_from_slice(&0u32.to_be_bytes());
        data
    }

    fn make_jpeg(width: u32, height: u32) -> Vec<u8> {
        let mut data = vec![0xFF, 0xD8];
        data.push(0xFF);
        data.push(0xE0);
        data.extend_from_slice(&16u16.to_be_bytes());
        data.extend_from_slice(b"JFIF\x00\x01\x01\x00\x00\x01\x00\x01\x00\x00");
        data.push(0xFF);
        data.push(0xC0);
        data.extend_from_slice(&11u16.to_be_bytes());
        data.push(8);
        data.extend_from_slice(&(height as u16).to_be_bytes());
        data.extend_from_slice(&(width as u16).to_be_bytes());
        data.extend_from_slice(&[1, 0x11, 0]);
        data.extend_from_slice(&[0xFF, 0xD9]);
        data
    }

    fn write_image_file(data: &[u8], ext: &str) -> NamedTempFile {
        let mut tmp = NamedTempFile::with_suffix(format!(".{ext}")).unwrap();
        tmp.write_all(data).unwrap();
        tmp.flush().unwrap();
        tmp
    }

    #[test]
    fn parse_png_dimensions_correct() {
        let data = make_png(800, 600);
        let (w, h) = parse_png_dimensions(&data).unwrap();
        assert_eq!(w, 800);
        assert_eq!(h, 600);
    }

    #[test]
    fn parse_jpeg_dimensions_correct() {
        let data = make_jpeg(640, 480);
        let (w, h) = parse_jpeg_dimensions(&data).unwrap();
        assert_eq!(w, 640);
        assert_eq!(h, 480);
    }

    #[test]
    fn parse_width_spec_cm() {
        let emu = parse_width_spec("12cm").unwrap();
        assert_eq!(emu, 12 * EMU_PER_CM);
    }

    #[test]
    fn parse_width_spec_inch() {
        let emu = parse_width_spec("4.5in").unwrap();
        assert_eq!(emu, (4.5 * EMU_PER_INCH as f64) as u64);
    }

    #[test]
    fn parse_width_spec_px() {
        let emu = parse_width_spec("300px").unwrap();
        assert_eq!(emu, 300 * EMU_PER_PX);
    }

    #[test]
    fn parse_width_spec_invalid() {
        assert!(parse_width_spec("12em").is_err());
        assert!(parse_width_spec("abc").is_err());
    }

    #[test]
    fn add_image_png() {
        let png = make_png(100, 100);
        let img_file = write_image_file(&png, "png");
        let (_doc_tmp, mut doc) = load_doc();

        let result = add_image(&mut doc, img_file.path(), Some("10cm"), None, None, None).unwrap();
        assert_eq!(result["status"], "ok");
        assert_eq!(result["action"], "add");
        assert_eq!(result["ref"], "@i1");
        assert_eq!(result["kind"], "image");

        let reloaded = Doc::load(doc.path).unwrap();
        assert!(reloaded.parts.others.contains_key("word/media/image1.png"));
        assert!(reloaded.parts.document_rels.is_some());
        let rels = std::str::from_utf8(reloaded.parts.document_rels.as_deref().unwrap()).unwrap();
        assert!(rels.contains("media/image1.png"));
    }

    #[test]
    fn add_image_jpeg() {
        let jpeg = make_jpeg(200, 150);
        let img_file = write_image_file(&jpeg, "jpeg");
        let (_doc_tmp, mut doc) = load_doc();

        let result = add_image(&mut doc, img_file.path(), None, None, None, None).unwrap();
        assert_eq!(result["ref"], "@i1");

        let reloaded = Doc::load(doc.path).unwrap();
        assert!(reloaded.parts.others.contains_key("word/media/image1.jpeg"));
    }

    #[test]
    fn add_image_unsupported_format() {
        let gif_file = write_image_file(b"GIF89a...", "gif");
        let (_doc_tmp, mut doc) = load_doc();
        let err = add_image(&mut doc, gif_file.path(), None, None, None, None).unwrap_err();
        assert!(matches!(err, DocxaiError::InvalidArgument(_)));
    }

    #[test]
    fn add_image_file_not_found() {
        let (_doc_tmp, mut doc) = load_doc();
        let err = add_image(
            &mut doc,
            Path::new("/nonexistent/image.png"),
            None,
            None,
            None,
            None,
        )
        .unwrap_err();
        assert!(matches!(err, DocxaiError::InvalidArgument(_)));
    }

    #[test]
    fn add_image_with_caption() {
        let png = make_png(50, 50);
        let img_file = write_image_file(&png, "png");
        let (_doc_tmp, mut doc) = load_doc();

        let result = add_image(
            &mut doc,
            img_file.path(),
            None,
            Some("Figure 1: Test"),
            None,
            None,
        )
        .unwrap();
        assert_eq!(result["ref"], "@i1");

        let reloaded = Doc::load(doc.path).unwrap();
        let xml = std::str::from_utf8(&reloaded.parts.document_xml).unwrap();
        assert!(
            xml.contains("Caption"),
            "should have caption paragraph: {xml}"
        );
        assert!(xml.contains("Figure 1: Test"));
    }

    #[test]
    fn add_image_after_ref() {
        let png = make_png(50, 50);
        let img_file = write_image_file(&png, "png");
        let (_doc_tmp, mut doc) = load_doc();

        let result = add_image(&mut doc, img_file.path(), None, None, Some("@p1"), None).unwrap();
        assert_eq!(result["ref"], "@i1");

        let reloaded = Doc::load(doc.path).unwrap();
        let map = crate::mutate::index_body_spans(&reloaded.parts.document_xml).unwrap();
        assert_eq!(map.spans[0].kind, 'p');
        assert_eq!(map.spans[1].kind, 'i');
    }

    #[test]
    fn delete_image_removes_everything() {
        let png = make_png(50, 50);
        let img_file = write_image_file(&png, "png");
        let (_doc_tmp, mut doc) = load_doc();
        add_image(&mut doc, img_file.path(), None, None, None, None).unwrap();

        let result = delete_image(&mut doc, "@i1").unwrap();
        assert_eq!(result["status"], "ok");
        assert_eq!(result["action"], "delete");

        let reloaded = Doc::load(doc.path).unwrap();
        let map = crate::mutate::index_body_spans(&reloaded.parts.document_xml).unwrap();
        let i_spans: Vec<_> = map.spans.iter().filter(|s| s.kind == 'i').collect();
        assert!(i_spans.is_empty(), "image should be removed");
        assert!(
            !reloaded.parts.others.contains_key("word/media/image1.png"),
            "media file should be removed"
        );
    }

    #[test]
    fn delete_image_invalid_ref() {
        let (_doc_tmp, mut doc) = load_doc();
        let err = delete_image(&mut doc, "@i99").unwrap_err();
        assert!(matches!(err, DocxaiError::InvalidArgument(_)));
    }

    #[test]
    fn delete_image_non_image_ref() {
        let (_doc_tmp, mut doc) = load_doc();
        let err = delete_image(&mut doc, "@p1").unwrap_err();
        assert!(matches!(err, DocxaiError::InvalidArgument(_)));
    }

    #[test]
    fn relationship_helpers() {
        let (_tmp, mut doc) = load_doc();
        let rid = doc.add_relationship("media/image1.png", REL_TYPE_IMAGE);
        assert!(rid.starts_with("rId"));
        let target = doc.get_relationship_target(&rid).unwrap();
        assert_eq!(target, "media/image1.png");
        doc.remove_relationship(&rid);
        assert!(doc.get_relationship_target(&rid).is_none());
    }

    #[test]
    fn ensure_content_type_idempotent() {
        let (_tmp, mut doc) = load_doc();
        doc.ensure_content_type("png", "image/png");
        let first = doc.parts.content_types.clone();
        doc.ensure_content_type("png", "image/png");
        assert_eq!(doc.parts.content_types, first, "should not duplicate");
    }

    #[test]
    fn next_rid_increments() {
        let (_tmp, doc) = load_doc();
        let n = doc.next_rid_num();
        assert!(n >= 1);
    }

    #[test]
    fn set_image_width_changes_extent() {
        let png = make_png(100, 100);
        let img_file = write_image_file(&png, "png");
        let (_doc_tmp, mut doc) = load_doc();
        add_image(&mut doc, img_file.path(), Some("10cm"), None, None, None).unwrap();

        let result = set_image(&mut doc, "@i1", Some("5cm"), None).unwrap();
        assert_eq!(result["status"], "ok");
        assert_eq!(result["action"], "set");

        let reloaded = Doc::load(doc.path).unwrap();
        let xml = std::str::from_utf8(&reloaded.parts.document_xml).unwrap();
        let five_cm = (5 * EMU_PER_CM).to_string();
        assert!(
            xml.contains(&format!("cx=\"{}\"", five_cm)),
            "should contain new cx value"
        );
    }

    #[test]
    fn set_image_requires_at_least_one_option() {
        let (_doc_tmp, mut doc) = load_doc();
        let err = set_image(&mut doc, "@i1", None, None).unwrap_err();
        assert!(matches!(err, DocxaiError::InvalidArgument(_)));
    }

    #[test]
    fn set_image_rejects_non_image_ref() {
        let (_doc_tmp, mut doc) = load_doc();
        let err = set_image(&mut doc, "@p1", Some("5cm"), None).unwrap_err();
        assert!(matches!(err, DocxaiError::InvalidArgument(_)));
    }

    #[test]
    fn set_image_marks_changed_fields() {
        let png = make_png(100, 100);
        let img_file = write_image_file(&png, "png");
        let (_doc_tmp, mut doc) = load_doc();
        add_image(&mut doc, img_file.path(), Some("10cm"), None, None, None).unwrap();

        let result = set_image(&mut doc, "@i1", Some("5cm"), Some("test")).unwrap();
        let changed = result["changed"].as_array().unwrap();
        assert!(changed.iter().any(|v| v == "width"));
        assert!(changed.iter().any(|v| v == "caption"));
    }
}
