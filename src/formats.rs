//! Cell model and built-in number formats (port of `Formats.ts`).

use chrono::NaiveDateTime;

/// Primitive (unformatted) cell value.
#[derive(Debug, Clone, PartialEq)]
pub enum PrimitiveCellValue {
    Empty,
    Text(String),
    Number(f64),
    Integer(i64),
    Boolean(bool),
    DateTime(NaiveDateTime),
}

/// A cell with an explicit Excel number-format string.
#[derive(Debug, Clone, PartialEq)]
pub struct FormattedCell {
    pub value: PrimitiveCellValue,
    pub format: String,
}

/// Any cell value: primitive or formatted.
#[derive(Debug, Clone, PartialEq)]
pub enum CellValue {
    Empty,
    Text(String),
    Number(f64),
    Integer(i64),
    Boolean(bool),
    DateTime(NaiveDateTime),
    Formatted(Box<FormattedCell>),
}

/// A borrowed cell value returned by the high-performance streaming readers.
///
/// Text values borrow either the shared-string table or the reader's scratch
/// buffer and are valid until the next call that advances the reader.
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum CellValueRef<'a> {
    Empty,
    Text(&'a str),
    Number(f64),
    Integer(i64),
    Boolean(bool),
    DateTime(NaiveDateTime),
}

/// A cell position and borrowed value produced by a streaming reader.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct CellRef<'a> {
    pub row: u32,
    pub column: u32,
    pub value: CellValueRef<'a>,
}

impl CellValueRef<'_> {
    /// Convert the borrowed value into the owned value used by the legacy API.
    pub fn to_owned(self) -> CellValue {
        match self {
            Self::Empty => CellValue::Empty,
            Self::Text(value) => CellValue::Text(value.to_owned()),
            Self::Number(value) => CellValue::Number(value),
            Self::Integer(value) => CellValue::Integer(value),
            Self::Boolean(value) => CellValue::Boolean(value),
            Self::DateTime(value) => CellValue::DateTime(value),
        }
    }
}

impl CellValue {
    pub fn empty() -> Self {
        CellValue::Empty
    }

    pub fn formatted(value: CellValue, format: impl Into<String>) -> Self {
        let primitive = match value {
            CellValue::Formatted(f) => f.value,
            CellValue::Empty => PrimitiveCellValue::Empty,
            CellValue::Text(s) => PrimitiveCellValue::Text(s),
            CellValue::Number(n) => PrimitiveCellValue::Number(n),
            CellValue::Integer(i) => PrimitiveCellValue::Integer(i),
            CellValue::Boolean(b) => PrimitiveCellValue::Boolean(b),
            CellValue::DateTime(d) => PrimitiveCellValue::DateTime(d),
        };
        CellValue::Formatted(Box::new(FormattedCell {
            value: primitive,
            format: format.into(),
        }))
    }
}

impl From<&str> for CellValue {
    fn from(s: &str) -> Self {
        CellValue::Text(s.to_owned())
    }
}

impl From<String> for CellValue {
    fn from(s: String) -> Self {
        CellValue::Text(s)
    }
}

impl From<f64> for CellValue {
    fn from(n: f64) -> Self {
        CellValue::Number(n)
    }
}

impl From<i64> for CellValue {
    fn from(n: i64) -> Self {
        CellValue::Integer(n)
    }
}

impl From<i32> for CellValue {
    fn from(n: i32) -> Self {
        CellValue::Integer(n as i64)
    }
}

impl From<bool> for CellValue {
    fn from(b: bool) -> Self {
        CellValue::Boolean(b)
    }
}

impl From<NaiveDateTime> for CellValue {
    fn from(d: NaiveDateTime) -> Self {
        CellValue::DateTime(d)
    }
}

/// Built-in Excel number-format strings for [`CellValue::Formatted`].
pub struct F;

impl F {
    pub const THOUSANDS_SEP: &'static str = "#,##0";
    pub const CURRENCY_PLN: &'static str = "#,##0.00 \"z\u{142}\"";
    pub const CURRENCY_EUR: &'static str = "#,##0.00 \u{20ac}";
    pub const PERCENTAGE: &'static str = "0%";
    pub const SCIENTIFIC: &'static str = "0.00E+00";
    pub const TWO_DECIMALS: &'static str = "#,##0.00";
    pub const TEXT: &'static str = "@";
    pub const LEADING_ZEROS: &'static str = "000000000";

    pub const DATE_SHORT: &'static str = "dd.mm.yyyy";
    pub const DATE_LONG: &'static str = "d mmmm yyyy";
    pub const DATE_DAY_MONTH_YEAR: &'static str = "dd-mm-yyyy";
    pub const DATE_ISO: &'static str = "yyyy-mm-dd";
    pub const DATE_MONTH_YEAR: &'static str = "mmmm yyyy";
    pub const DATE_WEEKDAY: &'static str = "dddd, d mmmm yyyy";
    pub const DATE_DAY_MONTH: &'static str = "d mmmm";
    pub const DATE_YEAR_ONLY: &'static str = "yyyy";

    pub const DATETIME_SHORT: &'static str = "dd.mm.yyyy hh:mm";
    pub const DATETIME_LONG: &'static str = "d mmmm yyyy hh:mm:ss";
    pub const TIME_HH_MM: &'static str = "hh:mm";
    pub const TIME_HH_MM_SS: &'static str = "hh:mm:ss";
    pub const TIME_12H: &'static str = "h:mm AM/PM";
    pub const DATETIME_24H: &'static str = "dd.mm.yyyy hh:mm:ss";
    pub const DATETIME_ISO: &'static str = "yyyy-mm-dd\"T\"hh:mm:ss";
    pub const TIME_MS: &'static str = "hh:mm:ss.000";
}

/// True when the value carries an explicit format.
pub fn is_formatted_cell(val: &CellValue) -> bool {
    matches!(val, CellValue::Formatted(_))
}

/// Strip the format wrapper, returning the primitive payload.
pub fn unwrap_cell(val: &CellValue) -> CellValue {
    match val {
        CellValue::Formatted(f) => match &f.value {
            PrimitiveCellValue::Empty => CellValue::Empty,
            PrimitiveCellValue::Text(s) => CellValue::Text(s.clone()),
            PrimitiveCellValue::Number(n) => CellValue::Number(*n),
            PrimitiveCellValue::Integer(i) => CellValue::Integer(*i),
            PrimitiveCellValue::Boolean(b) => CellValue::Boolean(*b),
            PrimitiveCellValue::DateTime(d) => CellValue::DateTime(*d),
        },
        other => other.clone(),
    }
}

/// Borrow the primitive value without cloning text payloads.
pub(crate) fn borrow_cell(val: &CellValue) -> CellValueRef<'_> {
    let primitive = match val {
        CellValue::Formatted(formatted) => &formatted.value,
        CellValue::Empty => return CellValueRef::Empty,
        CellValue::Text(value) => return CellValueRef::Text(value),
        CellValue::Number(value) => return CellValueRef::Number(*value),
        CellValue::Integer(value) => return CellValueRef::Integer(*value),
        CellValue::Boolean(value) => return CellValueRef::Boolean(*value),
        CellValue::DateTime(value) => return CellValueRef::DateTime(*value),
    };
    match primitive {
        PrimitiveCellValue::Empty => CellValueRef::Empty,
        PrimitiveCellValue::Text(value) => CellValueRef::Text(value),
        PrimitiveCellValue::Number(value) => CellValueRef::Number(*value),
        PrimitiveCellValue::Integer(value) => CellValueRef::Integer(*value),
        PrimitiveCellValue::Boolean(value) => CellValueRef::Boolean(*value),
        PrimitiveCellValue::DateTime(value) => CellValueRef::DateTime(*value),
    }
}

/// Detect date/time tokens in an Excel number format while ignoring quoted
/// literals and escaped characters such as `"mm"` and `\m`.
pub(crate) fn is_date_format_code(code: &str) -> bool {
    let mut quoted = false;
    let mut escaped = false;
    let mut bracket = String::new();
    let mut in_bracket = false;
    for ch in code.chars().flat_map(char::to_lowercase) {
        if escaped {
            escaped = false;
            continue;
        }
        if ch == '\\' && !quoted {
            escaped = true;
            continue;
        }
        if ch == '"' {
            quoted = !quoted;
            continue;
        }
        if quoted {
            continue;
        }
        if ch == '[' {
            in_bracket = true;
            bracket.clear();
            continue;
        }
        if ch == ']' && in_bracket {
            in_bracket = false;
            if matches!(bracket.as_str(), "h" | "m" | "s" | "d" | "y") {
                return true;
            }
            continue;
        }
        if in_bracket {
            bracket.push(ch);
        } else if matches!(ch, 'y' | 'm' | 'd' | 'h' | 's') {
            return true;
        }
    }
    false
}

/// The explicit format string, if any.
pub fn get_format(val: &CellValue) -> Option<&str> {
    match val {
        CellValue::Formatted(f) => Some(f.format.as_str()),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn formatted_helpers_roundtrip() {
        let v = CellValue::formatted(CellValue::Number(1.5), F::TWO_DECIMALS);
        assert!(is_formatted_cell(&v));
        assert_eq!(get_format(&v), Some("#,##0.00"));
        assert_eq!(unwrap_cell(&v), CellValue::Number(1.5));
        assert!(!is_formatted_cell(&CellValue::Text("x".into())));
        assert_eq!(get_format(&CellValue::Integer(3)), None);
    }

    #[test]
    fn date_format_detection_ignores_literals() {
        assert!(is_date_format_code("yyyy-mm-dd hh:mm"));
        assert!(is_date_format_code("[h]:mm:ss"));
        assert!(!is_date_format_code("0.00 \"mm\""));
        assert!(!is_date_format_code("0\\m"));
    }
}
