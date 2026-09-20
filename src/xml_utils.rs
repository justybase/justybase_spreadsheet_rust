//! Shared low-level XML helpers (port of `xmlUtils.ts`).

use quick_xml::events::Event;
use quick_xml::Reader as XmlReader;

fn has_invalid_control_chars(s: &str) -> bool {
    for c in s.chars() {
        let v = c as u32;
        if v < 0x20 && v != 0x09 && v != 0x0a && v != 0x0d {
            return true;
        }
    }
    false
}

fn strip_invalid_control_chars(s: &str) -> String {
    s.chars()
        .filter(|&c| {
            let v = c as u32;
            v == 0x09
                || v == 0x0a
                || v == 0x0d
                || (0x20..=0xd7ff).contains(&v)
                || (0xe000..=0xfffd).contains(&v)
        })
        .collect()
}

/// Escape text content for use inside an XML element body.
pub fn escape_xml_text(s: &str) -> String {
    if !s.contains(['&', '<', '>', '"', '\'']) && !has_invalid_control_chars(s) {
        return s.to_owned();
    }
    let mut out = s
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;");
    if has_invalid_control_chars(&out) {
        out = strip_invalid_control_chars(&out);
    }
    out
}

/// Unescape the five predefined XML entities plus numeric character references.
pub fn unescape_xml(s: &str) -> String {
    if s.is_empty() || !s.contains('&') {
        return s.to_owned();
    }
    // &#xHEX; references
    let mut out = String::with_capacity(s.len());
    let bytes = s.as_bytes();
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] == b'&' && i + 1 < bytes.len() && bytes[i + 1] == b'#' {
            if i + 2 < bytes.len() && (bytes[i + 2] == b'x' || bytes[i + 2] == b'X') {
                if let Some(semi) = s[i..].find(';') {
                    let hex = &s[i + 3..i + semi];
                    if let Ok(cp) = u32::from_str_radix(hex, 16) {
                        if let Some(ch) = char::from_u32(cp) {
                            out.push(ch);
                            i += semi + 1;
                            continue;
                        }
                    }
                }
            } else if let Some(semi) = s[i..].find(';') {
                let dec = &s[i + 2..i + semi];
                if !dec.is_empty() && dec.bytes().all(|b| b.is_ascii_digit()) {
                    if let Ok(cp) = dec.parse::<u32>() {
                        if let Some(ch) = char::from_u32(cp) {
                            out.push(ch);
                            i += semi + 1;
                            continue;
                        }
                    }
                }
            }
        }
        out.push(s[i..].chars().next().unwrap());
        i += s[i..].chars().next().unwrap().len_utf8();
    }
    out.replace("&amp;", "&")
        .replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", "\"")
        .replace("&apos;", "'")
}

/// Return whether an XLSX workbook declares the Macintosh 1904 date system.
pub fn uses_1904_date_system(xml: &str) -> bool {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(false);
    loop {
        match reader.read_event() {
            Ok(Event::Start(element)) | Ok(Event::Empty(element))
                if element.local_name().as_ref() == b"workbookPr" =>
            {
                return element
                    .attributes()
                    .flatten()
                    .find_map(|attr| {
                        (attr.key.local_name().as_ref() == b"date1904")
                            .then(|| String::from_utf8_lossy(attr.value.as_ref()).into_owned())
                    })
                    .is_some_and(|value| value == "1" || value.eq_ignore_ascii_case("true"));
            }
            Ok(Event::Eof) | Err(_) => return false,
            _ => {}
        }
    }
}

/// 0-based column index -> "A", "B", ..., "AA", ... (bijective base-26).
pub fn column_index_to_letter(col_index: usize) -> String {
    let mut n = col_index + 1;
    let mut result = String::new();
    while n > 0 {
        let rem = (n - 1) % 26;
        result.insert(0, (b'A' + rem as u8) as char);
        n = (n - 1) / 26;
    }
    result
}

/// "A", "AB", ... -> 0-based column index.
pub fn column_letter_to_index(letter: &str) -> usize {
    try_column_letter_to_index(letter).expect("column letter must contain A-Z only")
}

/// Checked variant of [`column_letter_to_index`].
pub fn try_column_letter_to_index(letter: &str) -> Option<usize> {
    if letter.is_empty() || !letter.bytes().all(|c| c.is_ascii_uppercase()) {
        return None;
    }
    let mut column: usize = 0;
    for c in letter.bytes() {
        column = column
            .checked_mul(26)?
            .checked_add((c - b'A' + 1) as usize)?;
    }
    column.checked_sub(1)
}

/// Remove trailing rows that contain no non-empty cells.
///
/// Functions take a predicate so both readers and updaters can reuse the
/// trimming rule (TS only trims `null`/`undefined`).
pub fn trim_trailing_empty_rows<T: Clone>(
    rows: Vec<Vec<T>>,
    is_empty: impl Fn(&T) -> bool,
) -> Vec<Vec<T>> {
    let mut end = rows.len();
    while end > 0 && rows[end - 1].iter().all(&is_empty) {
        end -= 1;
    }
    if end == rows.len() {
        rows
    } else {
        rows[..end].to_vec()
    }
}

/// Parse `xl/sharedStrings.xml` into string values.
///
/// Handles plain `<si><t>…</t></si>` and rich-text `<si><r><t>…</t></r></si>`
/// runs, with optional namespace prefixes.
pub fn parse_shared_strings_xml(xml: &str) -> Vec<String> {
    let mut reader = XmlReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut result = Vec::new();
    let mut in_item = false;
    let mut in_text = false;
    let mut value = String::new();
    loop {
        match reader.read_event() {
            Ok(Event::Start(element)) if element.local_name().as_ref() == b"si" => {
                in_item = true;
                in_text = false;
                value.clear();
            }
            Ok(Event::Start(element)) if in_item && element.local_name().as_ref() == b"t" => {
                in_text = true;
            }
            Ok(Event::End(element)) if element.local_name().as_ref() == b"t" => {
                in_text = false;
            }
            Ok(Event::End(element)) if element.local_name().as_ref() == b"si" => {
                if in_item {
                    result.push(std::mem::take(&mut value));
                    in_item = false;
                    in_text = false;
                }
            }
            Ok(Event::Text(text)) if in_item && in_text => {
                value.push_str(&String::from_utf8_lossy(text.as_ref()));
            }
            Ok(Event::GeneralRef(reference)) if in_item && in_text => {
                value.push_str(&decode_xml_reference(reference.as_ref()));
            }
            Ok(Event::CData(text)) if in_item && in_text => {
                value.push_str(&String::from_utf8_lossy(text.as_ref()));
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    result
}

fn decode_xml_reference(reference: &[u8]) -> String {
    let raw = format!("&{};", String::from_utf8_lossy(reference));
    quick_xml::escape::unescape(&raw)
        .map(|value| value.into_owned())
        .unwrap_or(raw)
}

fn tag_name_matches(xml: &str, pos: usize, names: &[&str]) -> Option<usize> {
    // pos points at '<'; returns offset just past '>' of the opening tag
    // when the tag name (modulo namespace prefix) is in `names`.
    if xml.as_bytes().get(pos) != Some(&b'<') {
        return None;
    }
    if xml[pos + 1..].starts_with('/') {
        return None;
    }
    let end = xml[pos..].find('>')? + pos;
    let inner = xml[pos + 1..end].trim();
    let raw_name = inner.split_whitespace().next().unwrap_or("");
    let local = raw_name.rsplit(':').next().unwrap_or(raw_name);
    // strip trailing '/' for self-closing tags
    let local = local.trim_end_matches('/');
    if names.contains(&local) {
        Some(end + 1)
    } else {
        None
    }
}

/// Find an opening XML tag by local name, ignoring namespace prefixes.
///
/// The spreadsheet readers intentionally use a small forward scanner instead
/// of building a DOM.  Keeping the tag scanner here makes all readers handle
/// both `sheetData` and `x:sheetData` documents consistently.
pub(crate) fn find_open_tag(xml: &str, from: usize, local: &str) -> Option<(usize, usize)> {
    let mut i = from.min(xml.len());
    while let Some(rel) = xml[i..].find('<') {
        let pos = i + rel;
        if let Some(after) = tag_name_matches(xml, pos, &[local]) {
            return Some((pos, after));
        }
        i = pos + 1;
        if i >= xml.len() {
            break;
        }
    }
    None
}

/// Find the byte offset of a closing XML tag by local name, ignoring prefixes.
pub(crate) fn find_close_tag(xml: &str, from: usize, local: &str) -> Option<usize> {
    let mut i = from.min(xml.len());
    while let Some(rel) = xml[i..].find("</") {
        let pos = i + rel;
        let end = xml[pos..].find('>')? + pos;
        let inner = xml[pos + 2..end].trim();
        let name = inner.split_whitespace().next().unwrap_or("");
        let name = name.rsplit(':').next().unwrap_or(name);
        if name.eq_ignore_ascii_case(local) {
            return Some(pos);
        }
        i = pos + 2;
        if i >= xml.len() {
            break;
        }
    }
    None
}

/// Read an XML attribute from an opening tag.
///
/// Both quote styles and arbitrary whitespace are accepted.  For prefixed
/// attributes such as `r:id`, the local part is also accepted so documents
/// using a different relationship prefix remain readable.
pub(crate) fn tag_attribute(tag: &str, wanted: &str) -> Option<String> {
    let wanted_local = wanted.rsplit(':').next().unwrap_or(wanted);
    let bytes = tag.as_bytes();
    let mut i = 0;
    // Skip the opening `<tag` name before scanning attributes.
    while i < bytes.len() && bytes[i] != b'<' {
        i += 1;
    }
    if i < bytes.len() {
        i += 1;
    }
    while i < bytes.len() && !bytes[i].is_ascii_whitespace() && bytes[i] != b'>' && bytes[i] != b'/'
    {
        i += 1;
    }
    while i < bytes.len() {
        while i < bytes.len() && (bytes[i].is_ascii_whitespace() || bytes[i] == b'<') {
            i += 1;
        }
        if i >= bytes.len() || bytes[i] == b'/' || bytes[i] == b'>' {
            break;
        }
        let key_start = i;
        while i < bytes.len()
            && !bytes[i].is_ascii_whitespace()
            && bytes[i] != b'='
            && bytes[i] != b'>'
            && bytes[i] != b'/'
        {
            i += 1;
        }
        let key = &tag[key_start..i];
        while i < bytes.len() && bytes[i].is_ascii_whitespace() {
            i += 1;
        }
        if i >= bytes.len() || bytes[i] != b'=' {
            while i < bytes.len() && !bytes[i].is_ascii_whitespace() {
                i += 1;
            }
            continue;
        }
        i += 1;
        while i < bytes.len() && bytes[i].is_ascii_whitespace() {
            i += 1;
        }
        let quote = *bytes.get(i)?;
        if quote != b'"' && quote != b'\'' {
            continue;
        }
        i += 1;
        let value_start = i;
        while i < bytes.len() && bytes[i] != quote {
            i += 1;
        }
        let value = &tag[value_start..i.min(bytes.len())];
        let key_local = key.rsplit(':').next().unwrap_or(key);
        if key.eq_ignore_ascii_case(wanted)
            || (wanted.contains(':') && key_local.eq_ignore_ascii_case(wanted_local))
        {
            return Some(unescape_xml(value));
        }
        if i < bytes.len() {
            i += 1;
        }
    }
    None
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn columns_roundtrip() {
        assert_eq!(column_index_to_letter(0), "A");
        assert_eq!(column_index_to_letter(25), "Z");
        assert_eq!(column_index_to_letter(26), "AA");
        assert_eq!(column_index_to_letter(27), "AB");
        assert_eq!(column_letter_to_index("A"), 0);
        assert_eq!(column_letter_to_index("AA"), 26);
        assert_eq!(column_letter_to_index("AB"), 27);
        assert_eq!(try_column_letter_to_index(""), None);
        assert_eq!(try_column_letter_to_index("a"), None);
        assert_eq!(try_column_letter_to_index(&"A".repeat(128)), None);
    }

    #[test]
    fn escape_unescape() {
        assert_eq!(escape_xml_text("a&b<c"), "a&amp;b&lt;c");
        assert_eq!(unescape_xml("a&amp;b&lt;c&#x41;&#66;"), "a&b<cAB");
        assert_eq!(escape_xml_text("plain"), "plain");
    }

    #[test]
    fn parses_1904_workbook_flag() {
        assert!(uses_1904_date_system(
            r#"<workbook><workbookPr date1904="1"/></workbook>"#
        ));
        assert!(uses_1904_date_system(
            r#"<workbook><workbookPr date1904="true"/></workbook>"#
        ));
        assert!(uses_1904_date_system(
            "<workbook><workbookPr date1904 = 'true'/></workbook>"
        ));
        assert!(!uses_1904_date_system(
            r#"<workbook><workbookPr date1904="0"/></workbook>"#
        ));
        assert!(!uses_1904_date_system(r#"<workbook/>"#));
    }

    #[test]
    fn shared_strings_parsing() {
        let xml = r#"<sst><!-- <si><t>ignored</t></si> --><si><t>Hello</t></si><si><r><t>Wo&amp;rld</t></r><r><t><![CDATA[!]]></t></r></si></sst>"#;
        assert_eq!(parse_shared_strings_xml(xml), vec!["Hello", "Wo&rld!"]);
    }

    #[test]
    fn shared_strings_ignore_formatting_between_text_runs() {
        let xml = "<sst>\n  <si>\n    <r><rPr/><t>first</t></r>\n    <r><t>second</t></r>\n  </si>\n</sst>";
        assert_eq!(parse_shared_strings_xml(xml), vec!["firstsecond"]);
    }
}
