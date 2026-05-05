//! Parser for `docxai`'s ref grammar (PRD §7.4).
//!
//! Refs are 1-indexed pointers into the document body:
//! `@p3` (3rd paragraph), `@t1` (1st table), `@t1.r2.c3` (cell),
//! `@i1` (1st image), `@e1` (1st equation).
//!
//! This module is strictly the *parser*. Resolving a ref to its XML
//! position in a loaded [`crate::doc::Doc`] requires walking the body
//! tree and is the next step (PRD #8 acceptance row 3).

use std::fmt;
use std::str::FromStr;

use crate::error::DocxaiError;

/// One parsed ref. Indices are always >= 1, matching PRD §7.4.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Ref {
    Paragraph(u32),
    Table(u32),
    TableCell { table: u32, row: u32, col: u32 },
    Image(u32),
    Equation(u32),
}

impl Ref {
    /// Parse a ref string. Errors are [`DocxaiError::InvalidArgument`]
    /// (PRD §10.1 → exit code 2).
    pub fn parse(input: &str) -> Result<Self, DocxaiError> {
        let body = input
            .strip_prefix('@')
            .ok_or_else(|| invalid(input, "missing leading '@'"))?;

        // Discriminate on the first byte. Body must be non-empty after '@'.
        let first = body
            .as_bytes()
            .first()
            .copied()
            .ok_or_else(|| invalid(input, "empty ref"))?;

        match first {
            b'p' => parse_simple(input, body, 'p').map(Ref::Paragraph),
            b'i' => parse_simple(input, body, 'i').map(Ref::Image),
            b'e' => parse_simple(input, body, 'e').map(Ref::Equation),
            b't' => parse_table_or_cell(input, body),
            _ => Err(invalid(input, "expected one of @p / @t / @i / @e")),
        }
    }
}

impl FromStr for Ref {
    type Err = DocxaiError;
    fn from_str(s: &str) -> Result<Self, Self::Err> {
        Self::parse(s)
    }
}

impl fmt::Display for Ref {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Ref::Paragraph(n) => write!(f, "@p{n}"),
            Ref::Table(n) => write!(f, "@t{n}"),
            Ref::TableCell { table, row, col } => write!(f, "@t{table}.r{row}.c{col}"),
            Ref::Image(n) => write!(f, "@i{n}"),
            Ref::Equation(n) => write!(f, "@e{n}"),
        }
    }
}

fn invalid(input: &str, why: &str) -> DocxaiError {
    DocxaiError::InvalidArgument(format!("invalid ref {input:?}: {why}"))
}

/// Parse `@<sigil><N>` where `<sigil>` is the single ASCII letter we already
/// matched and `<N>` is a 1-indexed integer with no leading zero.
fn parse_simple(input: &str, body: &str, sigil: char) -> Result<u32, DocxaiError> {
    let digits = &body[sigil.len_utf8()..];
    parse_index(input, digits, sigil)
}

fn parse_table_or_cell(input: &str, body: &str) -> Result<Ref, DocxaiError> {
    // Strip leading 't'.
    let rest = &body[1..];

    // `@tN` vs `@tN.rR.cC`.
    if let Some(dot) = rest.find('.') {
        let table = parse_index(input, &rest[..dot], 't')?;
        let after = &rest[dot + 1..];

        let r_dot = after
            .find('.')
            .ok_or_else(|| invalid(input, "table cell ref needs '.rR.cC'"))?;
        let r_part = &after[..r_dot];
        let c_part = &after[r_dot + 1..];

        let row_digits = r_part
            .strip_prefix('r')
            .ok_or_else(|| invalid(input, "expected '.rR' after table index"))?;
        let row = parse_index(input, row_digits, 'r')?;

        let col_digits = c_part
            .strip_prefix('c')
            .ok_or_else(|| invalid(input, "expected '.cC' after row index"))?;
        let col = parse_index(input, col_digits, 'c')?;

        Ok(Ref::TableCell { table, row, col })
    } else {
        Ok(Ref::Table(parse_index(input, rest, 't')?))
    }
}

/// Parse a 1-indexed positive integer. Reject empty, leading zero, sign,
/// non-digit. Returns u32 — overflow is also an error.
fn parse_index(input: &str, digits: &str, sigil: char) -> Result<u32, DocxaiError> {
    if digits.is_empty() {
        return Err(invalid(input, &format!("missing index after '{sigil}'")));
    }
    if !digits.bytes().all(|b| b.is_ascii_digit()) {
        return Err(invalid(input, &format!("non-digit after '{sigil}'")));
    }
    if digits.len() > 1 && digits.starts_with('0') {
        return Err(invalid(input, &format!("leading zero after '{sigil}'")));
    }
    let n: u32 = digits
        .parse()
        .map_err(|_| invalid(input, &format!("index after '{sigil}' overflows u32")))?;
    if n == 0 {
        return Err(invalid(
            input,
            &format!("index after '{sigil}' must be >= 1"),
        ));
    }
    Ok(n)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn ok(s: &str) -> Ref {
        Ref::parse(s).unwrap_or_else(|e| panic!("expected {s:?} to parse, got {e}"))
    }

    fn err(s: &str) -> DocxaiError {
        Ref::parse(s).expect_err(&format!("expected {s:?} to fail"))
    }

    #[test]
    fn parses_paragraph() {
        assert_eq!(ok("@p1"), Ref::Paragraph(1));
        assert_eq!(ok("@p42"), Ref::Paragraph(42));
        assert_eq!(ok("@p99999"), Ref::Paragraph(99_999));
    }

    #[test]
    fn parses_table() {
        assert_eq!(ok("@t1"), Ref::Table(1));
        assert_eq!(ok("@t10"), Ref::Table(10));
    }

    #[test]
    fn parses_table_cell() {
        assert_eq!(
            ok("@t1.r2.c3"),
            Ref::TableCell {
                table: 1,
                row: 2,
                col: 3
            }
        );
        assert_eq!(
            ok("@t12.r34.c56"),
            Ref::TableCell {
                table: 12,
                row: 34,
                col: 56
            }
        );
    }

    #[test]
    fn parses_image_and_equation() {
        assert_eq!(ok("@i1"), Ref::Image(1));
        assert_eq!(ok("@e7"), Ref::Equation(7));
    }

    #[test]
    fn round_trips_via_display() {
        for input in ["@p1", "@t3", "@t1.r2.c3", "@i9", "@e2"] {
            assert_eq!(ok(input).to_string(), input);
        }
    }

    #[test]
    fn rejects_missing_at_sign() {
        assert!(matches!(err("p1"), DocxaiError::InvalidArgument(_)));
        assert!(matches!(err(""), DocxaiError::InvalidArgument(_)));
    }

    #[test]
    fn rejects_unknown_sigil() {
        assert!(matches!(err("@x1"), DocxaiError::InvalidArgument(_)));
        assert!(matches!(err("@P1"), DocxaiError::InvalidArgument(_)));
    }

    #[test]
    fn rejects_zero_index() {
        for s in ["@p0", "@t0", "@i0", "@e0", "@t1.r0.c1", "@t1.r1.c0"] {
            assert!(
                matches!(err(s), DocxaiError::InvalidArgument(_)),
                "{s} should reject zero"
            );
        }
    }

    #[test]
    fn rejects_leading_zero() {
        for s in ["@p01", "@t02", "@t1.r02.c1"] {
            assert!(
                matches!(err(s), DocxaiError::InvalidArgument(_)),
                "{s} should reject leading zero"
            );
        }
    }

    #[test]
    fn rejects_negative_or_signed() {
        assert!(matches!(err("@p-1"), DocxaiError::InvalidArgument(_)));
        assert!(matches!(err("@p+1"), DocxaiError::InvalidArgument(_)));
    }

    #[test]
    fn rejects_missing_index() {
        for s in ["@p", "@t", "@i", "@e", "@t1.r.c1", "@t1.r1.c"] {
            assert!(
                matches!(err(s), DocxaiError::InvalidArgument(_)),
                "{s} should reject empty index"
            );
        }
    }

    #[test]
    fn rejects_malformed_cell() {
        for s in [
            "@t1.r2",       // missing .cC
            "@t1.c2.r3",    // wrong order
            "@t1.r2.c3.x4", // trailing junk
            "@t1..r2.c3",   // double dot
            "@t1.r2.c3 ",   // trailing space
            "@t1r2c3",      // missing dots
        ] {
            assert!(
                matches!(err(s), DocxaiError::InvalidArgument(_)),
                "{s} should reject malformed cell"
            );
        }
    }

    #[test]
    fn rejects_overflow() {
        let huge = format!("@p{}", u64::MAX);
        assert!(matches!(err(&huge), DocxaiError::InvalidArgument(_)));
    }

    #[test]
    fn invalid_argument_maps_to_exit_code_2() {
        let e = err("@x1");
        assert_eq!(e.exit_code(), crate::error::ExitCode::InvalidArgument);
    }

    #[test]
    fn from_str_works() {
        let r: Ref = "@p7".parse().unwrap();
        assert_eq!(r, Ref::Paragraph(7));
    }

    /// PRD #8: "100 refs valides parsées correctement". Sweep across all 5
    /// kinds and a wide range of indices.
    #[test]
    fn sweep_100_valid_refs() {
        let mut count = 0;
        for n in 1..=20u32 {
            assert_eq!(ok(&format!("@p{n}")), Ref::Paragraph(n));
            assert_eq!(ok(&format!("@t{n}")), Ref::Table(n));
            assert_eq!(ok(&format!("@i{n}")), Ref::Image(n));
            assert_eq!(ok(&format!("@e{n}")), Ref::Equation(n));
            assert_eq!(
                ok(&format!("@t{n}.r{n}.c{n}")),
                Ref::TableCell {
                    table: n,
                    row: n,
                    col: n
                }
            );
            count += 5;
        }
        assert_eq!(count, 100);
    }

    /// PRD #8: "50 refs invalides retournent erreur typée".
    #[test]
    fn sweep_50_invalid_refs() {
        let cases = [
            // missing prefix
            "p1",
            "t1",
            "",
            " @p1",
            "@",
            // wrong sigil
            "@x1",
            "@P1",
            "@T1",
            "@I1",
            "@E1",
            "@1",
            // missing index
            "@p",
            "@t",
            "@i",
            "@e",
            // zero / negative / signed
            "@p0",
            "@t0",
            "@i0",
            "@e0",
            "@p-1",
            "@p+1",
            // leading zero
            "@p01",
            "@t02",
            "@i03",
            // float / hex / non-ascii digits
            "@p1.5",
            "@p0x1",
            "@p一",
            // table cell malformations
            "@t1.r0.c1",
            "@t1.r1.c0",
            "@t0.r1.c1",
            "@t1.r2",
            "@t1.c2.r3",
            "@t1..r2.c3",
            "@t1.r.c1",
            "@t1.r1.c",
            "@t1r2c3",
            "@t1.x2.c3",
            "@t1.r2.x3",
            "@t1.r2.c3.x4",
            // whitespace
            "@p 1",
            "@p1 ",
            " @p1",
            "@ p1",
            // overflow
            "@p4294967296",
            // junk after a valid prefix
            "@p1abc",
            "@t1abc",
            "@i1abc",
            "@e1abc",
            "@t1.r2.c3abc",
            // upper-case row/col
            "@t1.R2.c3",
            "@t1.r2.C3",
        ];
        assert!(cases.len() >= 50, "need >=50 cases, got {}", cases.len());
        for s in cases {
            assert!(
                matches!(Ref::parse(s), Err(DocxaiError::InvalidArgument(_))),
                "{s:?} should be InvalidArgument"
            );
        }
    }
}
