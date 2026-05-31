//! Edge case tests for corrupt, partial, and empty .docx files (PRD #35).
//!
//! All tests verify that bad input produces clear errors without panicking.

use assert_cmd::Command;
use predicates::prelude::*;
use std::io::Write;
use tempfile::NamedTempFile;

fn docxai() -> Command {
    Command::cargo_bin("docxai").expect("binary `docxai` should be built")
}

fn write_temp(bytes: &[u8]) -> NamedTempFile {
    let mut tmp = NamedTempFile::new().unwrap();
    tmp.write_all(bytes).unwrap();
    tmp.flush().unwrap();
    tmp
}

// Returns a TempPath (the OS file handle is closed) so the spawned `docxai`
// binary can atomically replace the file on Windows, where renaming over a
// still-open file is denied. The file is deleted when the TempPath drops.
fn make_minimal_docx() -> tempfile::TempPath {
    let mut tmp = NamedTempFile::new().unwrap();
    let buf: Vec<u8> = Vec::new();
    let w = std::io::Cursor::new(buf);
    let mut zip = zip::ZipWriter::new(w);
    let options = zip::write::SimpleFileOptions::default();

    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
        <Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\
        <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\
        <Default Extension=\"xml\" ContentType=\"application/xml\"/>\
        <Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>\
        <Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/>\
        </Types>").unwrap();

    zip.start_file("_rels/.rels", options).unwrap();
    zip.write_all(b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
        <Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\
        <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>\
        <Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"word/styles.xml\"/>\
        </Relationships>").unwrap();

    zip.start_file("word/_rels/document.xml.rels", options)
        .unwrap();
    zip.write_all(
        b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
        <Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\
        </Relationships>",
    )
    .unwrap();

    zip.start_file("word/document.xml", options).unwrap();
    zip.write_all(
        b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
        <w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"\
        xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\
        <w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body></w:document>",
    )
    .unwrap();

    zip.start_file("word/styles.xml", options).unwrap();
    zip.write_all(
        b"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\
        <w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">\
        <w:style w:type=\"paragraph\" w:styleId=\"Title\"><w:name w:val=\"Title\"/></w:style>\
        <w:style w:type=\"paragraph\" w:styleId=\"Body\"><w:name w:val=\"Body\"/></w:style>\
        </w:styles>",
    )
    .unwrap();

    let buf = zip.finish().unwrap().into_inner();
    tmp.write_all(&buf).unwrap();
    tmp.flush().unwrap();
    tmp.into_temp_path()
}

#[test]
fn not_a_zip_produces_error_not_panic() {
    let tmp = write_temp(b"this is not a zip file at all");
    docxai()
        .args(["snapshot", tmp.path().to_str().unwrap()])
        .assert()
        .failure()
        .stderr(predicate::str::contains("not a valid zip"));
}

#[test]
fn empty_file_produces_error() {
    let tmp = write_temp(b"");
    docxai()
        .args(["snapshot", tmp.path().to_str().unwrap()])
        .assert()
        .failure();
}

#[test]
fn zip_without_content_types_produces_error() {
    let buf: Vec<u8> = Vec::new();
    let w = std::io::Cursor::new(buf);
    let mut zip = zip::ZipWriter::new(w);
    let options = zip::write::SimpleFileOptions::default();
    zip.start_file("word/document.xml", options).unwrap();
    zip.write_all(b"<w:document/>").unwrap();
    let buf = zip.finish().unwrap().into_inner();

    let tmp = write_temp(&buf);
    docxai()
        .args(["snapshot", tmp.path().to_str().unwrap()])
        .assert()
        .failure()
        .stderr(predicate::str::contains("Content_Types"));
}

#[test]
fn zip_without_document_xml_produces_error() {
    let buf: Vec<u8> = Vec::new();
    let w = std::io::Cursor::new(buf);
    let mut zip = zip::ZipWriter::new(w);
    let options = zip::write::SimpleFileOptions::default();
    zip.start_file("[Content_Types].xml", options).unwrap();
    zip.write_all(b"<Types/>").unwrap();
    let buf = zip.finish().unwrap().into_inner();

    let tmp = write_temp(&buf);
    docxai()
        .args(["snapshot", tmp.path().to_str().unwrap()])
        .assert()
        .failure()
        .stderr(predicate::str::contains("document.xml"));
}

#[test]
fn missing_file_produces_clear_error() {
    docxai()
        .args(["snapshot", "/nonexistent/file.docx"])
        .assert()
        .failure()
        .stderr(predicate::str::contains("cannot open"));
}

#[test]
fn invalid_ref_on_set_produces_error() {
    let tmp = make_minimal_docx();
    docxai()
        .args(["set", tmp.to_str().unwrap(), "@p999", "--text", "x"])
        .assert()
        .failure()
        .stderr(predicate::str::contains("not found").or(predicate::str::contains("error")));
}

#[test]
fn invalid_ref_on_delete_produces_error() {
    let tmp = make_minimal_docx();
    docxai()
        .args(["delete", tmp.to_str().unwrap(), "@t999"])
        .assert()
        .failure();
}

#[test]
fn malformed_ref_produces_error() {
    let tmp = make_minimal_docx();
    docxai()
        .args(["set", tmp.to_str().unwrap(), "not_a_ref", "--text", "x"])
        .assert()
        .failure()
        .stderr(predicate::str::contains("invalid ref"));
}

#[test]
fn garbage_binary_does_not_panic() {
    let mut tmp = NamedTempFile::new().unwrap();
    let garbage: Vec<u8> = (0..1000).map(|i| (i as u8).wrapping_mul(37)).collect();
    tmp.write_all(&garbage).unwrap();
    tmp.flush().unwrap();

    let result = docxai()
        .args(["snapshot", tmp.path().to_str().unwrap()])
        .assert();
    let _ = result;
}

#[test]
fn snapshot_on_minimal_docx_succeeds() {
    let tmp = make_minimal_docx();
    docxai()
        .args(["snapshot", tmp.to_str().unwrap()])
        .assert()
        .success()
        .stdout(predicate::str::contains("\"ref\":\"@p1\""));
}

#[test]
fn styles_on_minimal_docx_succeeds() {
    let tmp = make_minimal_docx();
    docxai()
        .args(["styles", tmp.to_str().unwrap()])
        .assert()
        .success()
        .stdout(predicate::str::contains("Title"));
}

#[test]
fn add_paragraph_on_minimal_docx_succeeds() {
    let tmp = make_minimal_docx();
    docxai()
        .args([
            "add",
            tmp.to_str().unwrap(),
            "paragraph",
            "--text",
            "New paragraph",
        ])
        .assert()
        .success()
        .stdout(predicate::str::contains("\"status\":\"ok\""));
}

#[test]
fn set_paragraph_on_minimal_docx_succeeds() {
    let tmp = make_minimal_docx();
    docxai()
        .args(["set", tmp.to_str().unwrap(), "@p1", "--text", "Updated"])
        .assert()
        .success()
        .stdout(predicate::str::contains("\"status\":\"ok\""));
}

#[test]
fn delete_paragraph_on_minimal_docx_succeeds() {
    let tmp = make_minimal_docx();
    docxai()
        .args(["delete", tmp.to_str().unwrap(), "@p1"])
        .assert()
        .success();
}
