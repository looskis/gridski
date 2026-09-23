//! `build_rows`: turn row specs for a period sheet into one write plan.
//!
//! A time-series model is rows × periods with one formula per row. The caller gives each
//! row's label, units, total and first-period formula; we resolve `{key}` references
//! between rows, lay out the block, and describe the formats, so the bridge can write,
//! fill right and format everything in one Excel round trip.

use std::collections::HashMap;

use serde::{Deserialize, Serialize};
use serde_json::Value;

use crate::cells::{cell_address, column_letters, top_left};

pub const MAX_ROWS: usize = 500;
pub const MAX_CELLS: usize = 60_000;

#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize, schemars::JsonSchema)]
#[serde(rename_all = "lowercase")]
pub enum RowStyle {
    /// Section heading: bold label.
    Header,
    /// Subtotal: bold, thin top border across the row.
    Total,
    /// Link to another sheet: green period cells.
    Link,
    /// Hardcoded inputs: blue period and total cells.
    Input,
}

#[derive(Debug, Clone, Deserialize, schemars::JsonSchema)]
pub struct RowSpec {
    /// Name other rows use to refer to this one: {key}, {prev:key}, {row:key}, {total:key}.
    pub key: Option<String>,
    pub label: Option<String>,
    pub units: Option<String>,
    /// Total/constant column: a formula (may use {row:key}, {total:key}) or a value.
    pub total: Option<Value>,
    /// First-period cell, copied across every period: a formula or a value. Omit for rows
    /// with no period values (headings, blanks, constants).
    pub first: Option<Value>,
    /// Number format for the total and period cells, e.g. "#,##0;(#,##0);-".
    pub number_format: Option<String>,
    pub style: Option<RowStyle>,
}

#[derive(Debug, Clone)]
pub struct Layout {
    pub start_row: u32,
    pub label_col: u32,
    pub units_col: u32,
    pub total_col: u32,
    pub first_col: u32,
    pub last_col: u32,
}

#[derive(Debug, Serialize, PartialEq)]
pub struct FormatOp {
    pub range: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub number_format: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_color: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub top_border: Option<bool>,
}

#[derive(Debug, Serialize)]
pub struct Plan {
    /// Whole block the call replaces, label column to last period.
    pub target: String,
    /// Label column through the first period, written as a grid.
    pub head_start: String,
    pub head: Vec<Vec<Value>>,
    /// Period columns after the first, filled from each row's first cell (R1C1).
    pub fill: Option<String>,
    pub first_col: u32,
    pub formats: Vec<FormatOp>,
    pub keys: HashMap<String, u32>,
}

const GREEN: &str = "#008000";
const BLUE: &str = "#0000FF";

pub fn column(letters: &str) -> Result<u32, String> {
    top_left(&format!("{}1", letters.trim().to_ascii_uppercase()))
        .map(|(_, c)| c)
        .ok_or_else(|| format!("\"{letters}\" is not a column letter."))
}

pub fn plan(layout: &Layout, rows: &[RowSpec]) -> Result<Plan, String> {
    let l = layout;
    if rows.is_empty() || rows.len() > MAX_ROWS {
        return Err(format!("rows must hold 1 to {MAX_ROWS} rows."));
    }
    if !(l.label_col < l.first_col && l.units_col < l.first_col && l.total_col < l.first_col && l.first_col <= l.last_col) {
        return Err("Columns must run label/units/total before first_period, and first_period ≤ last_period.".into());
    }
    let width = (l.last_col - l.label_col + 1) as usize;
    if rows.len() * width > MAX_CELLS {
        return Err(format!("This block has {} cells; split it into calls of at most {MAX_CELLS}.", rows.len() * width));
    }

    let mut keys = HashMap::new();
    for (i, r) in rows.iter().enumerate() {
        if let Some(k) = &r.key {
            if !is_key(k) {
                return Err(format!("Key \"{k}\" must be letters, digits, _ or . and start with a letter."));
            }
            if keys.insert(k.clone(), l.start_row + i as u32).is_some() {
                return Err(format!("Key \"{k}\" is used twice."));
            }
        }
    }

    let head_width = (l.first_col - l.label_col + 1) as usize;
    let mut head = Vec::with_capacity(rows.len());
    for (i, r) in rows.iter().enumerate() {
        let mut line = vec![Value::Null; head_width];
        let at = |col: u32| (col - l.label_col) as usize;
        if let Some(v) = &r.label {
            line[at(l.label_col)] = text(v);
        }
        if let Some(v) = &r.units {
            line[at(l.units_col)] = text(v);
        }
        let row_no = l.start_row + i as u32;
        for (value, col) in [(&r.total, l.total_col), (&r.first, l.first_col)] {
            if let Some(v) = value {
                line[at(col)] = resolve_value(v, &keys, l).map_err(|e| format!("Row {row_no}: {e}"))?;
            }
        }
        head.push(line);
    }

    let bottom = l.start_row + rows.len() as u32 - 1;
    Ok(Plan {
        target: range(l.start_row, l.label_col, bottom, l.last_col),
        head_start: cell_address(l.start_row, l.label_col),
        head,
        fill: (l.last_col > l.first_col).then(|| range(l.start_row, l.first_col + 1, bottom, l.last_col)),
        first_col: l.first_col,
        formats: formats(l, rows),
        keys,
    })
}

/// Labels and units are text: keep Excel from reading "1/0" as a date or "12" as a number.
fn text(s: &str) -> Value {
    let parsable = s.starts_with(|c: char| c.is_ascii_digit() || matches!(c, '-' | '+' | '.' | '='));
    Value::String(if parsable { format!("'{s}") } else { s.to_string() })
}

fn is_key(s: &str) -> bool {
    let mut chars = s.chars();
    chars.next().is_some_and(|c| c.is_ascii_alphabetic() || c == '_')
        && chars.all(|c| c.is_ascii_alphanumeric() || c == '_' || c == '.')
}

fn range(t: u32, l: u32, b: u32, r: u32) -> String {
    if t == b && l == r {
        cell_address(t, l)
    } else {
        format!("{}:{}", cell_address(t, l), cell_address(b, r))
    }
}

fn resolve_value(v: &Value, keys: &HashMap<String, u32>, l: &Layout) -> Result<Value, String> {
    match v {
        Value::String(s) if s.starts_with('=') => resolve(s, keys, l).map(Value::String),
        Value::String(_) | Value::Number(_) | Value::Bool(_) | Value::Null => Ok(v.clone()),
        _ => Err("cells must be a formula, number, string, or boolean.".into()),
    }
}

/// Replace {key}, {prev:key}, {row:key}, {total:key}. Braces that don't hold a key-like
/// token (array constants such as {1,2}) are left alone; an unknown key is an error.
pub fn resolve(formula: &str, keys: &HashMap<String, u32>, l: &Layout) -> Result<String, String> {
    let mut out = String::with_capacity(formula.len());
    let mut rest = formula;
    while let Some(open) = rest.find('{') {
        out.push_str(&rest[..open]);
        let Some(len) = rest[open..].find('}') else {
            out.push_str(&rest[open..]);
            return Ok(out);
        };
        let token = &rest[open + 1..open + len];
        let (kind, key) = match token.split_once(':') {
            Some((k @ ("prev" | "row" | "total"), key)) => (k, key),
            _ => ("", token),
        };
        if !is_key(key) {
            out.push_str(&rest[open..=open + len]);
        } else {
            let r = *keys.get(key).ok_or_else(|| format!("unknown key {{{token}}}."))?;
            out.push_str(&match kind {
                "prev" => cell_address(r, l.first_col - 1),
                "row" => format!("${}${r}:${}${r}", column_letters(l.first_col), column_letters(l.last_col)),
                "total" => format!("${}${r}", column_letters(l.total_col)),
                _ => cell_address(r, l.first_col),
            });
        }
        rest = &rest[open + len + 1..];
    }
    out.push_str(rest);
    Ok(out)
}

/// Group consecutive rows that share a format into one range, since each format call
/// over Apple Events costs about the same for one row as for fifty.
fn formats(l: &Layout, rows: &[RowSpec]) -> Vec<FormatOp> {
    let mut ops = Vec::new();
    let mut run = |pick: &dyn Fn(&RowSpec) -> Option<FormatOp>, cols: (u32, u32)| {
        let mut i = 0;
        while i < rows.len() {
            let Some(op) = pick(&rows[i]) else {
                i += 1;
                continue;
            };
            let mut j = i;
            while j + 1 < rows.len() && pick(&rows[j + 1]).as_ref() == Some(&op) {
                j += 1;
            }
            let (t, b) = (l.start_row + i as u32, l.start_row + j as u32);
            ops.push(FormatOp { range: range(t, cols.0, b, cols.1), ..op });
            i = j + 1;
        }
    };
    let blank = || FormatOp { range: String::new(), number_format: None, bold: None, font_color: None, top_border: None };
    run(&|r| r.number_format.clone().map(|f| FormatOp { number_format: Some(f), ..blank() }), (l.total_col, l.last_col));
    run(&|r| (r.style == Some(RowStyle::Header)).then(|| FormatOp { bold: Some(true), ..blank() }), (l.label_col, l.label_col));
    run(
        &|r| (r.style == Some(RowStyle::Total)).then(|| FormatOp { bold: Some(true), top_border: Some(true), ..blank() }),
        (l.label_col, l.last_col),
    );
    run(&|r| (r.style == Some(RowStyle::Link)).then(|| FormatOp { font_color: Some(GREEN.into()), ..blank() }), (l.first_col, l.last_col));
    run(&|r| (r.style == Some(RowStyle::Input)).then(|| FormatOp { font_color: Some(BLUE.into()), ..blank() }), (l.total_col, l.last_col));
    ops
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    fn layout() -> Layout {
        Layout { start_row: 5, label_col: 1, units_col: 2, total_col: 3, first_col: 5, last_col: 16 }
    }

    fn spec(v: Value) -> Vec<RowSpec> {
        serde_json::from_value(v).unwrap()
    }

    #[test]
    fn resolves_keys_and_lays_out() {
        let rows = spec(json!([
            {"label": "Loan", "style": "header"},
            {"key": "open", "label": "Opening", "first": "={prev:close}"},
            {"key": "draw", "label": "Draw", "units": "$", "first": "=Amount*Timing!E$9", "total": "=SUM({row:draw})", "number_format": "#,##0"},
            {"key": "close", "label": "Closing", "first": "={open}+{draw}", "style": "total", "number_format": "#,##0"},
            {"label": "Peak", "total": "=MAX({row:close})/{total:draw}"}
        ]));
        let p = plan(&layout(), &rows).unwrap();
        assert_eq!(p.target, "A5:P9");
        assert_eq!(p.fill.as_deref(), Some("F5:P9"));
        assert_eq!(p.head[1][4], json!("=D8"));
        assert_eq!(p.head[2][2], json!("=SUM($E$7:$P$7)"));
        assert_eq!(p.head[3][4], json!("=E6+E7"));
        assert_eq!(p.head[4][2], json!("=MAX($E$8:$P$8)/$C$7"));
        assert_eq!(p.keys["close"], 8);
        assert_eq!(text("1/0"), json!("'1/0"));
        assert_eq!(text("$"), json!("$"));
        let nf: Vec<_> = p.formats.iter().filter(|f| f.number_format.is_some()).map(|f| f.range.as_str()).collect();
        assert_eq!(nf, ["C7:P8"]);
        assert!(p.formats.iter().any(|f| f.range == "A8:P8" && f.top_border == Some(true)));
    }

    #[test]
    fn rejects_unknown_keys_but_keeps_array_constants() {
        let l = layout();
        let keys = HashMap::from([("a".to_string(), 5)]);
        assert_eq!(resolve("=SUM({1,2})+{a}", &keys, &l).unwrap(), "=SUM({1,2})+E5");
        assert!(resolve("={b}", &keys, &l).unwrap_err().contains("unknown key"));
        assert!(plan(&l, &spec(json!([{"key": "x"}, {"key": "x"}]))).is_err());
    }
}
