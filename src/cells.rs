//! Turning raw Excel output into compact, unambiguous tool results.

use serde::Serialize;
use serde_json::{Map, Value};

use crate::bridge::RawRange;

/// Excel's error values. `value()` reports these as `""`; `stringValue()` has the code.
const ERROR_VALUES: &[&str] = &[
    "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA", "#SPILL!", "#CALC!",
    "#FIELD!", "#BLOCKED!", "#CONNECT!", "#BUSY!", "#UNKNOWN!",
];

#[derive(Debug, Serialize)]
pub struct RangeOutput {
    pub workbook: String,
    pub sheet: String,
    /// Address of the cells in `values`, e.g. `A1:C3`.
    pub address: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub truncated: Option<Truncation>,
    /// Row-major cell values. Empty cells are `null`; dates are `YYYY-MM-DD`; errors are
    /// their code, e.g. `#DIV/0!`.
    pub values: Vec<Vec<Value>>,
    /// Only cells containing formulas, keyed by address.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub formulas: Option<Map<String, Value>>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub areas: Option<Vec<String>>,
}

#[derive(Debug, Serialize, PartialEq)]
pub struct Truncation {
    pub total_rows: usize,
    pub total_cols: usize,
    /// Where to continue reading, when rows were cut off.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub next_range: Option<String>,
}

pub fn normalize(raw: RawRange) -> RangeOutput {
    let address = strip_dollars(&raw.address);
    let (top, left) = top_left(&address).unwrap_or((1, 1));
    let rows = raw.values.len();
    let cols = raw.values.first().map_or(0, Vec::len);

    let values = raw
        .values
        .into_iter()
        .zip(raw.text)
        .map(|(row, text_row)| {
            row.into_iter()
                .enumerate()
                .map(|(c, v)| normalize_cell(v, text_row.get(c)))
                .collect()
        })
        .collect();

    let formulas = raw.formulas.map(|grid| {
        let mut map = Map::new();
        for (r, row) in grid.into_iter().enumerate() {
            for (c, f) in row.into_iter().enumerate() {
                if matches!(&f, Value::String(s) if s.starts_with('=')) {
                    map.insert(cell_address(top + r as u32, left + c as u32), f);
                }
            }
        }
        map
    });

    let truncated = (rows < raw.total_rows || cols < raw.total_cols).then(|| Truncation {
        total_rows: raw.total_rows,
        total_cols: raw.total_cols,
        next_range: (rows < raw.total_rows && cols == raw.total_cols).then(|| {
            let right = left + cols as u32 - 1;
            let bottom = top + raw.total_rows as u32 - 1;
            format!("{}:{}", cell_address(top + rows as u32, left), cell_address(bottom, right))
        }),
    });

    RangeOutput {
        workbook: raw.workbook,
        sheet: raw.sheet,
        address,
        truncated,
        values,
        formulas,
        areas: raw.areas.map(|a| a.iter().map(|s| strip_dollars(s)).collect()),
    }
}

fn normalize_cell(value: Value, text: Option<&Value>) -> Value {
    match value {
        Value::String(s) if s.is_empty() => match text {
            Some(Value::String(t)) if ERROR_VALUES.contains(&t.as_str()) => Value::String(t.clone()),
            _ => Value::Null,
        },
        v => v,
    }
}

pub(crate) fn strip_dollars(address: &str) -> String {
    address.replace('$', "")
}

/// Row and column (1-based) of the first cell in an A1 address like `B2:D9`.
pub(crate) fn top_left(address: &str) -> Option<(u32, u32)> {
    let first = address.split([':', ',']).next()?;
    let split = first.find(|c: char| c.is_ascii_digit())?;
    let (letters, digits) = first.split_at(split);
    let col = letters
        .chars()
        .try_fold(0u32, |acc, ch| ch.is_ascii_alphabetic().then(|| acc * 26 + (ch.to_ascii_uppercase() as u32 - 'A' as u32 + 1)))?;
    Some((digits.parse().ok()?, col))
}

/// (top, left, bottom, right) of an A1 range such as `$B$2:$D$9` or `C5`.
pub(crate) fn bounds(address: &str) -> Option<(u32, u32, u32, u32)> {
    let address = strip_dollars(address);
    let address = address.rsplit('!').next()?;
    let (top, left) = top_left(address)?;
    let (bottom, right) = match address.split_once(':') {
        Some((_, end)) => top_left(end)?,
        None => (top, left),
    };
    Some((top, left, bottom, right))
}

pub(crate) fn column_letters(mut col: u32) -> String {
    let mut out = Vec::new();
    while col > 0 {
        let rem = (col - 1) % 26;
        out.push(b'A' + rem as u8);
        col = (col - 1) / 26;
    }
    out.reverse();
    String::from_utf8(out).expect("ascii")
}

pub(crate) fn cell_address(row: u32, col: u32) -> String {
    format!("{}{row}", column_letters(col))
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    fn raw(address: &str, total: (usize, usize), values: Value, text: Value, formulas: Option<Value>) -> RawRange {
        RawRange {
            workbook: "Book1".into(),
            sheet: "Sheet1".into(),
            address: address.into(),
            total_rows: total.0,
            total_cols: total.1,
            values: serde_json::from_value(values).unwrap(),
            text: serde_json::from_value(text).unwrap(),
            formulas: formulas.map(|f| serde_json::from_value(f).unwrap()),
            areas: None,
        }
    }

    #[test]
    fn column_letter_round_trip() {
        for (n, s) in [(1, "A"), (26, "Z"), (27, "AA"), (52, "AZ"), (703, "AAA"), (16384, "XFD")] {
            assert_eq!(column_letters(n), s);
            assert_eq!(top_left(&format!("{s}7")), Some((7, n)));
        }
        assert_eq!(top_left("B2:D9"), Some((2, 2)));
    }

    #[test]
    fn errors_and_blanks_are_distinguished() {
        let out = normalize(raw(
            "$A$1:$C$1",
            (1, 3),
            json!([["", "", 3]]),
            json!([["", "#DIV/0!", "3"]]),
            None,
        ));
        assert_eq!(out.values, vec![vec![Value::Null, json!("#DIV/0!"), json!(3)]]);
        assert_eq!(out.address, "A1:C1");
        assert!(out.truncated.is_none());
    }

    #[test]
    fn formulas_are_sparse_and_addressed() {
        let out = normalize(raw(
            "$B$2:$C$3",
            (2, 2),
            json!([[1, 2], [3, 5]]),
            json!([["1", "2"], ["3", "5"]]),
            Some(json!([["1", "2"], ["3", "=B3+2"]])),
        ));
        assert_eq!(Value::Object(out.formulas.unwrap()), json!({"C3": "=B3+2"}));
    }

    #[test]
    fn truncation_points_at_remaining_rows() {
        let out = normalize(raw("$A$1:$B$2", (10, 2), json!([[1, 2], [3, 4]]), json!([["1", "2"], ["3", "4"]]), None));
        assert_eq!(
            out.truncated,
            Some(Truncation { total_rows: 10, total_cols: 2, next_range: Some("A3:B10".into()) })
        );
    }
}
