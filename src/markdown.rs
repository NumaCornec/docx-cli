//! Markdown subset rendering and parsing (PRD §9, issues #11/#12).
//!
//! v0.1 supports the inline subset: `**bold**`, `*italic*`, `***both***`,
//! backslash escapes for the meta-characters, and a hard line break (`  \n`)
//! that becomes a literal `\n` inside a run. Anything outside that subset is
//! rejected with [`DocxaiError::InvalidArgument`] (PRD §9.2).
//!
//! The shape exposed is symmetric: [`render_runs`] is the inverse of
//! [`parse_runs`] over the supported subset. Snapshot (#10) consumes
//! `render_runs`; paragraph mutations (#14/#16) will consume `parse_runs`.
//!
//! Note: PRD §9.1 lists `_italic_` as accepted, but the v0.1 contract for
//! this issue uses `*` as the sole emphasis marker (`_` is reserved for a
//! later issue and is rejected with a clear message for now).

use crate::error::DocxaiError;

/// One inline run extracted from / destined for `<w:r>`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Run {
    pub text: String,
    pub bold: bool,
    pub italic: bool,
}

/// Render a sequence of runs to the markdown subset (PRD §9.1, used by snapshot).
/// Inverse of [`parse_runs`] over the supported subset.
pub fn render_runs(runs: &[Run]) -> String {
    let mut out = String::new();
    for run in runs {
        if run.text.is_empty() {
            continue;
        }
        let escaped = escape_markdown(&run.text);
        match (run.bold, run.italic) {
            (true, true) => {
                out.push_str("***");
                out.push_str(&escaped);
                out.push_str("***");
            }
            (true, false) => {
                out.push_str("**");
                out.push_str(&escaped);
                out.push_str("**");
            }
            (false, true) => {
                out.push('*');
                out.push_str(&escaped);
                out.push('*');
            }
            (false, false) => out.push_str(&escaped),
        }
    }
    out
}

/// Parse the markdown subset into runs (PRD §9.1). Errors with
/// [`DocxaiError::InvalidArgument`] if input contains unsupported syntax (§9.2).
pub fn parse_runs(input: &str) -> Result<Vec<Run>, DocxaiError> {
    // Reject block-level constructs up front. The v0.1 subset is single-paragraph
    // inline-only, so pure newlines (i.e. not part of a `  \n` hard break) are
    // unsupported per §9.2 ("\n: paragraphe break? Non, erreur.").
    reject_unsupported_block_constructs(input)?;

    let mut runs: Vec<Run> = Vec::new();
    let mut buf = String::new();
    let mut bold = false;
    let mut italic = false;

    let bytes = input.as_bytes();
    let mut i = 0;
    while i < bytes.len() {
        let b = bytes[i];

        // Backslash escapes: §9.3 + the explicit task contract.
        if b == b'\\' {
            if i + 1 >= bytes.len() {
                return Err(DocxaiError::InvalidArgument(
                    "trailing backslash without escaped character".into(),
                ));
            }
            let next = bytes[i + 1];
            match next {
                b'\\' | b'*' | b'_' | b'`' | b'$' | b'[' | b']' => {
                    buf.push(next as char);
                    i += 2;
                    continue;
                }
                b'\n' => {
                    // `\<newline>` is not part of our subset; the hard break is two
                    // spaces plus newline, handled below.
                    return Err(DocxaiError::InvalidArgument(
                        "backslash before newline is not supported; use two trailing spaces for a hard line break".into(),
                    ));
                }
                _ => {
                    return Err(DocxaiError::InvalidArgument(format!(
                        "unsupported escape sequence: \\{}",
                        next as char
                    )));
                }
            }
        }

        // Hard line break: `  \n` → literal `\n` in the current run's text.
        if b == b' ' && i + 2 < bytes.len() && bytes[i + 1] == b' ' && bytes[i + 2] == b'\n' {
            // Skip leading whitespace+newline runs from the input. Consume
            // the two spaces and the newline; emit a literal `\n` to keep
            // the snapshot's `<w:br/>` contract.
            buf.push('\n');
            i += 3;
            continue;
        }

        // Reject characters that can only mean unsupported syntax (already
        // covered by the block-construct check, but catch inline cases too).
        match b {
            b'`' => {
                return Err(DocxaiError::InvalidArgument(
                    "code spans (`) are not supported in the v0.1 markdown subset".into(),
                ));
            }
            b'$' => {
                return Err(DocxaiError::InvalidArgument(
                    "inline math ($...$) is not supported in the v0.1 markdown subset".into(),
                ));
            }
            b'[' | b']' | b'(' | b')' => {
                // Parentheses on their own are fine prose, but in conjunction
                // with `[..]` they form a link. Detect the link shape.
                if b == b'[' && contains_link_shape(&bytes[i..]) {
                    return Err(DocxaiError::InvalidArgument(
                        "links ([text](url)) are not supported in the v0.1 markdown subset".into(),
                    ));
                }
                if b == b'[' || b == b']' {
                    return Err(DocxaiError::InvalidArgument(
                        "brackets must be escaped (\\[, \\]) in the v0.1 markdown subset".into(),
                    ));
                }
                // Bare `(` / `)` are passthrough literals.
                buf.push(b as char);
                i += 1;
                continue;
            }
            b'_' => {
                return Err(DocxaiError::InvalidArgument(
                    "underscore emphasis is not supported in v0.1; use * for emphasis or escape with \\_".into(),
                ));
            }
            b'\n' => {
                // A bare newline (not preceded by two spaces) means multi-paragraph
                // input, which is rejected per PRD §9.1: one paragraph per command.
                return Err(DocxaiError::InvalidArgument(
                    "newline in --text is not supported; use a separate command per paragraph"
                        .into(),
                ));
            }
            _ => {}
        }

        // Emphasis markers.
        if b == b'*' {
            // Count consecutive asterisks (1, 2, or 3).
            let mut n = 1usize;
            while i + n < bytes.len() && bytes[i + n] == b'*' {
                n += 1;
            }
            if n > 3 {
                return Err(DocxaiError::InvalidArgument(
                    "more than three consecutive '*' is not supported".into(),
                ));
            }

            // Decide whether this is an opener or closer.
            let want_bold = n >= 2;
            let want_italic = n == 1 || n == 3;

            // Mismatch detection: a closer must match the active state.
            let is_closer = match n {
                1 => italic,
                2 => bold,
                3 => bold && italic,
                _ => unreachable!(),
            };
            let is_opener = match n {
                1 => !italic,
                2 => !bold,
                3 => !bold && !italic,
                _ => unreachable!(),
            };

            if !is_opener && !is_closer {
                return Err(DocxaiError::InvalidArgument(format!(
                    "unmatched emphasis run of {n} '*'"
                )));
            }

            // Flush current buffered text into a run with the *current* formatting.
            flush_run(&mut runs, &mut buf, bold, italic);

            if is_opener {
                if want_bold {
                    bold = true;
                }
                if want_italic {
                    italic = true;
                }
            } else {
                // Closer.
                if want_bold {
                    bold = false;
                }
                if want_italic {
                    italic = false;
                }
            }

            i += n;
            continue;
        }

        // Plain character — push to buffer. (Use the original UTF-8 character
        // boundary, not the byte, to preserve multi-byte chars.)
        let ch_len = utf8_char_len(b);
        // Safety: `input` is a valid &str, so a complete codepoint is present.
        let ch_end = i + ch_len;
        let ch = &input[i..ch_end];
        buf.push_str(ch);
        i = ch_end;
    }

    if bold || italic {
        return Err(DocxaiError::InvalidArgument(
            "unmatched emphasis: input ends with an open '*' run".into(),
        ));
    }

    flush_run(&mut runs, &mut buf, bold, italic);
    Ok(runs)
}

/// Backslash-escape characters that have meaning in our markdown subset.
/// Preserves `\n` and `\t` literally — those represent `<w:br/>` and `<w:tab/>`
/// inside a run (PRD #10's snapshot contract).
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

fn flush_run(runs: &mut Vec<Run>, buf: &mut String, bold: bool, italic: bool) {
    if buf.is_empty() {
        return;
    }
    let text = std::mem::take(buf);
    runs.push(Run { text, bold, italic });
}

fn utf8_char_len(first_byte: u8) -> usize {
    match first_byte {
        0x00..=0x7F => 1,
        0xC0..=0xDF => 2,
        0xE0..=0xEF => 3,
        0xF0..=0xF7 => 4,
        // Continuation byte at the start would mean an invalid &str — but &str
        // is guaranteed valid, so this branch is unreachable in practice.
        _ => 1,
    }
}

/// Cheap structural check: does this slice (starting at `[`) look like
/// `[...](...)`? Used only to produce a precise error message.
fn contains_link_shape(rest: &[u8]) -> bool {
    // Find a `]` then `(` then `)` in order, skipping escaped `]`.
    let mut i = 1; // we know rest[0] == b'['
    let mut found_close_bracket = None;
    while i < rest.len() {
        let b = rest[i];
        if b == b'\\' && i + 1 < rest.len() {
            i += 2;
            continue;
        }
        if b == b']' {
            found_close_bracket = Some(i);
            break;
        }
        i += 1;
    }
    let Some(close) = found_close_bracket else {
        return false;
    };
    // Next non-trivial byte should be `(`.
    if close + 1 >= rest.len() || rest[close + 1] != b'(' {
        return false;
    }
    rest[close + 2..].contains(&b')')
}

/// Reject constructs that the v0.1 subset cannot represent (§9.2). This is the
/// first line of defence; some are also caught inside the inline loop, but
/// detecting them on the raw input gives a more precise error.
fn reject_unsupported_block_constructs(input: &str) -> Result<(), DocxaiError> {
    // Triple-backtick code fence anywhere is unsupported.
    if input.contains("```") {
        return Err(DocxaiError::InvalidArgument(
            "fenced code blocks (```) are not supported in the v0.1 markdown subset".into(),
        ));
    }
    // The whole input is treated as one paragraph; line-prefix checks operate
    // on its leading whitespace-trimmed start. Multi-line input is itself
    // rejected later, but headings/lists at position 0 deserve a tailored msg.
    let trimmed = input.trim_start_matches([' ', '\t']);
    if let Some(rest) = trimmed.strip_prefix('#') {
        // Treat `#`, `##`, ... followed by space as ATX heading.
        let after_hashes = rest.trim_start_matches('#');
        if after_hashes.starts_with(' ') {
            return Err(DocxaiError::InvalidArgument(
                "ATX headings (# ) are not supported; use --style Heading1 etc.".into(),
            ));
        }
    }
    if trimmed.starts_with("- ") || trimmed.starts_with("* ") {
        // `* ` at the start would also be parsed as italic-open; the list
        // form takes priority because it's a more useful error.
        // Disambiguate: `* foo` looks like a list; `*foo*` is italic.
        // We've already trimmed leading whitespace, and the second char is a
        // space, so this is unambiguously a list bullet.
        return Err(DocxaiError::InvalidArgument(
            "unordered lists (- , * ) are not supported; use --style ListBullet".into(),
        ));
    }
    if let Some(after) = strip_ordered_list_prefix(trimmed) {
        let _ = after; // we don't need the remainder for the error
        return Err(DocxaiError::InvalidArgument(
            "ordered lists (1. ) are not supported; use --style ListNumber".into(),
        ));
    }
    Ok(())
}

/// If `s` starts with `<digits>. ` (one or more digits, then a dot, then a
/// space), return the remainder. Otherwise None.
fn strip_ordered_list_prefix(s: &str) -> Option<&str> {
    let bytes = s.as_bytes();
    let mut i = 0;
    while i < bytes.len() && bytes[i].is_ascii_digit() {
        i += 1;
    }
    if i == 0 {
        return None;
    }
    if bytes.get(i).copied() == Some(b'.') && bytes.get(i + 1).copied() == Some(b' ') {
        Some(&s[i + 2..])
    } else {
        None
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn run(text: &str, bold: bool, italic: bool) -> Run {
        Run {
            text: text.to_string(),
            bold,
            italic,
        }
    }

    #[test]
    fn render_runs_emits_bold_italic_both() {
        let runs = vec![
            run("plain", false, false),
            run("b", true, false),
            run("i", false, true),
            run("both", true, true),
        ];
        assert_eq!(render_runs(&runs), "plain**b***i****both***");
    }

    #[test]
    fn render_runs_escapes_special_chars() {
        let runs = vec![run(r"a*b_c`d$e[f]g\h", false, false)];
        assert_eq!(render_runs(&runs), r"a\*b\_c\`d\$e\[f\]g\\h");
    }

    #[test]
    fn render_runs_preserves_tab_and_newline() {
        // <w:br/> and <w:tab/> are inserted as literal '\n' / '\t' in run text;
        // render must not escape them.
        let runs = vec![run("a\tb\nc", false, false)];
        assert_eq!(render_runs(&runs), "a\tb\nc");
    }

    #[test]
    fn render_runs_skips_empty_runs() {
        let runs = vec![run("", true, true), run("x", false, false)];
        assert_eq!(render_runs(&runs), "x");
    }

    #[test]
    fn parse_runs_round_trips_render_for_simple_cases() {
        for s in ["plain", "**bold**", "*italic*", "***both***", "a**b**c"] {
            let runs = parse_runs(s).unwrap_or_else(|e| panic!("parse {s:?}: {e}"));
            assert_eq!(render_runs(&runs), s, "round-trip {s:?}");
        }
    }

    #[test]
    fn parse_runs_handles_escapes() {
        let runs = parse_runs(r"a\*b").unwrap();
        assert_eq!(runs, vec![run("a*b", false, false)]);
    }

    #[test]
    fn parse_runs_handles_all_documented_escapes() {
        let runs = parse_runs(r"\\ \* \_ \` \$ \[ \]").unwrap();
        assert_eq!(runs, vec![run(r"\ * _ ` $ [ ]", false, false)]);
    }

    #[test]
    fn parse_runs_handles_adjacent_emphasis() {
        let runs = parse_runs("**a**b*c*").unwrap();
        assert_eq!(
            runs,
            vec![
                run("a", true, false),
                run("b", false, false),
                run("c", false, true),
            ]
        );
    }

    #[test]
    fn parse_runs_handles_bold_italic_combo() {
        let runs = parse_runs("***both***").unwrap();
        assert_eq!(runs, vec![run("both", true, true)]);
    }

    #[test]
    fn parse_runs_rejects_lists() {
        for input in ["- item", "* item", "1. item"] {
            let err = parse_runs(input).unwrap_err();
            assert!(matches!(err, DocxaiError::InvalidArgument(_)), "{input:?}");
        }
    }

    #[test]
    fn parse_runs_rejects_headings() {
        let err = parse_runs("# title").unwrap_err();
        assert!(matches!(err, DocxaiError::InvalidArgument(_)));
    }

    #[test]
    fn parse_runs_rejects_code_spans() {
        for input in ["`code`", "```\nblock\n```"] {
            let err = parse_runs(input).unwrap_err();
            assert!(matches!(err, DocxaiError::InvalidArgument(_)), "{input:?}");
        }
    }

    #[test]
    fn parse_runs_rejects_links() {
        let err = parse_runs("[t](u)").unwrap_err();
        assert!(matches!(err, DocxaiError::InvalidArgument(_)));
    }

    #[test]
    fn parse_runs_rejects_inline_math() {
        let err = parse_runs("$x$").unwrap_err();
        assert!(matches!(err, DocxaiError::InvalidArgument(_)));
    }

    #[test]
    fn parse_runs_rejects_underscore_emphasis() {
        let err = parse_runs("_x_").unwrap_err();
        assert!(matches!(err, DocxaiError::InvalidArgument(_)));
    }

    #[test]
    fn parse_runs_rejects_multiline() {
        let err = parse_runs("a\nb").unwrap_err();
        assert!(matches!(err, DocxaiError::InvalidArgument(_)));
    }

    #[test]
    fn parse_runs_accepts_hard_break() {
        let runs = parse_runs("a  \nb").unwrap();
        let combined: String = runs.iter().map(|r| r.text.as_str()).collect();
        assert!(combined.contains('\n'), "combined text was {combined:?}");
        // Specifically: the `\n` separates `a` and `b`.
        assert!(combined.contains("a\nb") || combined.contains("a\n"));
    }

    #[test]
    fn parse_runs_unmatched_emphasis_errors() {
        let err = parse_runs("**unclosed").unwrap_err();
        assert!(matches!(err, DocxaiError::InvalidArgument(_)));
    }

    #[test]
    fn parse_runs_trailing_backslash_errors() {
        let err = parse_runs(r"foo\").unwrap_err();
        assert!(matches!(err, DocxaiError::InvalidArgument(_)));
    }

    #[test]
    fn parse_runs_empty_input_yields_no_runs() {
        let runs = parse_runs("").unwrap();
        assert!(runs.is_empty());
    }

    #[test]
    fn render_then_parse_round_trip_with_escapes() {
        let original = vec![run("a*b[c]d", false, false), run("bold", true, false)];
        let rendered = render_runs(&original);
        let parsed = parse_runs(&rendered).unwrap();
        assert_eq!(parsed, original);
    }
}
