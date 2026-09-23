//! Formula consistency checks for time-series models.
//!
//! A well-built row uses one formula across every period. Comparing R1C1 text (where a
//! copied formula reads identically in every column) finds the cells that break the
//! pattern, and scanning A1 text finds numbers typed into formulas.

use serde::Serialize;
use serde_json::Value;

use crate::cells::{cell_address, top_left};

/// Numbers a formula may contain without counting as a hardcode.
const STRUCTURAL: &[f64] = &[0.0, 1.0, 12.0];

/// Rows need this many formula cells before their consistency is judged; below it
/// they are summary blocks, not time series.
const MIN_SERIES: usize = 3;

#[derive(Debug, Serialize)]
pub struct AuditOutput {
    pub workbook: String,
    pub sheet: String,
    pub address: String,
    pub formula_cells: usize,
    /// Set when cells went unexamined or findings were left out; `note` says which.
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub truncated: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub note: Option<String>,
    /// Flagged row blocks, counting every row folded into `same_rows`.
    pub inconsistent_rows_total: usize,
    /// Rows where formulas change pattern, or constants sit between formulas.
    pub inconsistent_rows: Vec<RowFinding>,
    pub embedded_numbers_total: usize,
    /// Formulas with numbers typed in, grouped by identical R1C1 pattern.
    pub embedded_numbers: Vec<EmbeddedNumbers>,
}

#[derive(Debug, Serialize, PartialEq)]
pub struct RowFinding {
    pub row: u32,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub label: Option<String>,
    /// Other rows whose runs have exactly this shape (same columns, same R1C1 patterns),
    /// e.g. "12-20, 31"; `runs` shows this row's formulas only.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub same_rows: Option<String>,
    /// Consecutive cells sharing one formula pattern, in column order.
    pub runs: Vec<Run>,
    /// Cells holding typed values between the row's first and last formula.
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub constants: Vec<String>,
}

#[derive(Debug, Serialize, PartialEq)]
pub struct Run {
    pub range: String,
    /// The first cell's formula, in A1 notation.
    pub formula: String,
}

#[derive(Debug, Serialize, PartialEq)]
pub struct EmbeddedNumbers {
    /// First cell with this pattern, and how many cells share it.
    pub cell: String,
    pub count: usize,
    pub formula: String,
    pub numbers: Vec<String>,
}

/// Audit a block of formulas (A1 and R1C1 text, row by row), listing at most
/// `max_findings` inconsistent rows and `max_findings` embedded-number groups.
pub fn audit(
    workbook: String,
    sheet: String,
    address: &str,
    a1: &[Vec<Value>],
    r1c1: &[Vec<Value>],
    max_findings: usize,
) -> AuditOutput {
    let address = address.replace('$', "");
    let (top, left) = top_left(&address).unwrap_or((1, 1));
    let text = |v: &Value| v.as_str().unwrap_or_default().to_string();
    let is_formula = |s: &str| s.starts_with('=');

    let mut formula_cells = 0;
    // Findings keyed by shape, so rows copied down from one another collapse into one.
    let mut inconsistent: Vec<(String, RowFinding, Vec<u32>)> = Vec::new();
    let mut inconsistent_total = 0;
    let mut embedded: Vec<(String, EmbeddedNumbers)> = Vec::new();

    for (ri, (row_a1, row_r1c1)) in a1.iter().zip(r1c1).enumerate() {
        let row = top + ri as u32;
        let cells: Vec<(u32, String, String)> = row_a1
            .iter()
            .zip(row_r1c1)
            .enumerate()
            .map(|(ci, (f, r))| (left + ci as u32, text(f), text(r)))
            .collect();
        let formulas: Vec<&(u32, String, String)> = cells.iter().filter(|(_, f, _)| is_formula(f)).collect();
        formula_cells += formulas.len();

        for (col, f, r) in &formulas {
            let numbers = typed_numbers(f);
            if numbers.is_empty() {
                continue;
            }
            match embedded.iter_mut().find(|(pattern, _)| pattern == r) {
                Some((_, e)) => e.count += 1,
                None => embedded.push((
                    r.clone(),
                    EmbeddedNumbers { cell: cell_address(row, *col), count: 1, formula: f.clone(), numbers },
                )),
            }
        }

        // A blank cell separates blocks (e.g. a totals column, a spacer, then the periods),
        // so each block is judged on its own.
        let label = cells
            .iter()
            .find(|(_, f, _)| !f.is_empty() && !is_formula(f) && f.parse::<f64>().is_err())
            .map(|(_, f, _)| f.clone());
        for block in cells.split(|(_, f, _)| f.is_empty()) {
            if let Some((shape, finding)) = judge_block(row, block, label.clone()) {
                inconsistent_total += 1;
                match inconsistent.iter_mut().find(|(s, _, _)| *s == shape) {
                    Some((_, _, rows)) => rows.push(row),
                    None => inconsistent.push((shape, finding, Vec::new())),
                }
            }
        }
    }

    let (inconsistent_groups, embedded_total) = (inconsistent.len(), embedded.len());
    let inconsistent_rows: Vec<RowFinding> = inconsistent
        .into_iter()
        .take(max_findings)
        .map(|(_, mut finding, rows)| {
            finding.same_rows = (!rows.is_empty()).then(|| row_list(&rows));
            finding
        })
        .collect();
    embedded.truncate(max_findings);
    let truncated = inconsistent_groups > max_findings || embedded_total > max_findings;
    AuditOutput {
        workbook,
        sheet,
        address,
        formula_cells,
        truncated,
        note: truncated.then(|| {
            format!(
                "Listed {} of {inconsistent_groups} inconsistent-row groups and {} of {embedded_total} embedded-number \
                 groups; narrow range or raise max_findings to see the rest.",
                inconsistent_rows.len(),
                embedded.len()
            )
        }),
        inconsistent_rows_total: inconsistent_total,
        inconsistent_rows,
        embedded_numbers_total: embedded_total,
        embedded_numbers: embedded.into_iter().map(|(_, e)| e).collect(),
    }
}

/// Compact row numbers: [3, 4, 5, 9] → "3-5, 9".
fn row_list(rows: &[u32]) -> String {
    let mut parts: Vec<String> = Vec::new();
    let mut i = 0;
    while i < rows.len() {
        let mut j = i;
        while j + 1 < rows.len() && rows[j + 1] == rows[j] + 1 {
            j += 1;
        }
        parts.push(if i == j { rows[i].to_string() } else { format!("{}-{}", rows[i], rows[j]) });
        i = j + 1;
    }
    parts.join(", ")
}

/// Judge one contiguous block of non-blank cells in a row: flag it when its formulas
/// change pattern or typed values sit between them. Also returns the finding's shape
/// (run columns and R1C1 patterns, constant columns), which rows copied down share.
fn judge_block(row: u32, block: &[(u32, String, String)], label: Option<String>) -> Option<(String, RowFinding)> {
    let is_formula = |s: &str| s.starts_with('=');
    let formulas: Vec<&(u32, String, String)> = block.iter().filter(|(_, f, _)| is_formula(f)).collect();
    if formulas.len() < MIN_SERIES {
        return None;
    }
    let (first, last) = (formulas[0].0, formulas[formulas.len() - 1].0);
    let mut runs: Vec<(u32, u32, &str, &str)> = Vec::new();
    for (col, f, r) in &formulas {
        match runs.last_mut() {
            Some((_, end, pattern, _)) if *pattern == r.as_str() && *end + 1 == *col => *end = *col,
            _ => runs.push((*col, *col, r, f)),
        }
    }
    let constant_cols: Vec<u32> = block
        .iter()
        .filter(|(col, f, _)| *col > first && *col < last && !is_formula(f))
        .map(|(col, _, _)| *col)
        .collect();
    let mut patterns: Vec<&str> = runs.iter().map(|r| r.2).collect();
    patterns.sort_unstable();
    patterns.dedup();
    if patterns.len() == 1 && constant_cols.is_empty() {
        return None;
    }
    let mut shape: String =
        runs.iter().map(|(start, end, pattern, _)| format!("{start}-{end}{pattern}\u{1f}")).collect();
    shape.extend(constant_cols.iter().map(|col| format!("c{col}\u{1f}")));
    let constants = constant_cols.iter().map(|col| cell_address(row, *col)).collect();
    let runs = runs
        .into_iter()
        .map(|(start, end, _, f)| Run {
            range: if start == end {
                cell_address(row, start)
            } else {
                format!("{}:{}", cell_address(row, start), cell_address(row, end))
            },
            formula: f.to_string(),
        })
        .collect();
    Some((shape, RowFinding { row, label, same_rows: None, runs, constants }))
}


/// Numeric literals in an A1 formula other than structural 0, 1, and 12. Skips string
/// literals, quoted sheet names, and digits that belong to references or names (`B12`,
/// `$F$34`, `LOG10`, `Sheet2`).
fn typed_numbers(formula: &str) -> Vec<String> {
    let chars: Vec<char> = formula.chars().collect();
    let mut found = Vec::new();
    let mut i = 0;
    while i < chars.len() {
        let c = chars[i];
        if c == '"' || c == '\'' {
            // Skip to the closing quote; doubled quotes are escapes and just re-enter.
            i += 1;
            while i < chars.len() && chars[i] != c {
                i += 1;
            }
            i += 1;
            continue;
        }
        let starts_number = c.is_ascii_digit() || (c == '.' && chars.get(i + 1).is_some_and(char::is_ascii_digit));
        if !starts_number {
            i += 1;
            continue;
        }
        let part_of_word = i > 0 && {
            let p = chars[i - 1];
            p.is_alphanumeric() || matches!(p, '$' | '_' | '.' | ':' | '!')
        };
        let start = i;
        while i < chars.len() && (chars[i].is_ascii_digit() || chars[i] == '.') {
            i += 1;
        }
        if i < chars.len() && matches!(chars[i], 'e' | 'E') && chars.get(i + 1).is_some_and(|n| n.is_ascii_digit() || *n == '-' || *n == '+') {
            i += 2;
            while i < chars.len() && chars[i].is_ascii_digit() {
                i += 1;
            }
        }
        // Row-range references like `5:5` and names continuing after the digits.
        let continues_word = chars.get(i).is_some_and(|n| n.is_alphabetic() || matches!(n, '_' | ':' | '!'));
        if part_of_word || continues_word {
            continue;
        }
        let literal: String = chars[start..i].iter().collect();
        let percent = chars.get(i) == Some(&'%');
        let value = literal.parse::<f64>().unwrap_or(f64::NAN) / if percent { 100.0 } else { 1.0 };
        if !STRUCTURAL.contains(&value) {
            found.push(if percent { format!("{literal}%") } else { literal });
        }
    }
    found
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    #[test]
    fn finds_typed_numbers() {
        assert_eq!(typed_numbers("=+Loan_Amount*0.02"), ["0.02"]);
        assert_eq!(typed_numbers("=IRR(D59:EF59,0%)*12"), Vec::<String>::new());
        assert_eq!(typed_numbers("='Sheet 2'!$F$34*B12+LOG10(C3)"), Vec::<String>::new());
        assert_eq!(typed_numbers("=IF(A1=\"5\",3,1)"), ["3"]);
        assert_eq!(typed_numbers("=E5+1.5e3-5%"), ["1.5e3", "5%"]);
        assert_eq!(typed_numbers("=SUM(5:5)"), Vec::<String>::new());
    }

    #[test]
    fn flags_broken_series() {
        let a1 = vec![vec![json!("Rent"), json!("=100"), json!("=C1"), json!("=D1*1.02"), json!("=E1"), json!("7"), json!("=G1")]];
        let r1c1 = vec![vec![json!("Rent"), json!("=100"), json!("=RC[-1]"), json!("=RC[-1]*1.02"), json!("=RC[-1]"), json!("7"), json!("=RC[-1]")]];
        let out = audit("wb".into(), "s".into(), "$A$1:$G$1", &a1, &r1c1, 50);
        assert_eq!(out.formula_cells, 5);
        let row = &out.inconsistent_rows[0];
        assert_eq!(row.label.as_deref(), Some("Rent"));
        let ranges: Vec<_> = row.runs.iter().map(|r| r.range.as_str()).collect();
        assert_eq!(ranges, ["B1", "C1", "D1", "E1", "G1"]);
        assert_eq!(row.constants, ["F1"]);
        let cells: Vec<_> = out.embedded_numbers.iter().map(|e| (e.cell.as_str(), e.numbers.clone())).collect();
        assert_eq!(cells, [("B1", vec!["100".to_string()]), ("D1", vec!["1.02".to_string()])]);
    }

    #[test]
    fn consistent_series_passes() {
        let a1 = vec![vec![json!("=A1+1"), json!("=B1+1"), json!("=C1+1")]];
        let r1c1 = vec![vec![json!("=RC[-1]+1"), json!("=RC[-1]+1"), json!("=RC[-1]+1")]];
        let out = audit("wb".into(), "s".into(), "B1:D1", &a1, &r1c1, 50);
        assert!(out.inconsistent_rows.is_empty() && out.embedded_numbers.is_empty());
    }

    #[test]
    fn blank_column_separates_blocks() {
        // label, total, spacer, then periods: the total is not part of the series.
        let a1 = vec![
            vec![json!("NOI"), json!("=SUM(E1:G1)"), json!(""), json!("=A2"), json!("=B2"), json!("=C2")],
            vec![json!("Tax"), json!("=SUM(E2:G2)"), json!(""), json!("=A3"), json!("=B3*2"), json!("=C3")],
        ];
        let r1c1 = vec![
            vec![json!("NOI"), json!("=SUM(RC[2]:RC[4])"), json!(""), json!("=R[1]C[-3]"), json!("=R[1]C[-3]"), json!("=R[1]C[-3]")],
            vec![json!("Tax"), json!("=SUM(RC[2]:RC[4])"), json!(""), json!("=R[1]C[-3]"), json!("=R[1]C[-3]*2"), json!("=R[1]C[-3]")],
        ];
        let out = audit("wb".into(), "s".into(), "B1:G2", &a1, &r1c1, 50);
        assert_eq!(out.inconsistent_rows.len(), 1);
        let row = &out.inconsistent_rows[0];
        assert_eq!((row.row, row.label.as_deref()), (2, Some("Tax")));
        let ranges: Vec<_> = row.runs.iter().map(|r| r.range.as_str()).collect();
        assert_eq!(ranges, ["E2", "F2", "G2"]);
    }

    #[test]
    fn collapses_copied_rows_and_caps_findings() {
        // Rows 1-3 break the same way (copied down); row 4 breaks differently.
        let a1: Vec<Vec<Value>> = (1..=4)
            .map(|r| vec![json!(format!("=A{r}")), json!(format!("=B{r}*2")), json!(format!("=C{r}"))])
            .collect();
        let mut r1c1: Vec<Vec<Value>> =
            (1..=3).map(|_| vec![json!("=RC[-1]"), json!("=RC[-1]*2"), json!("=RC[-1]")]).collect();
        r1c1.push(vec![json!("=RC[-1]*2"), json!("=RC[-1]"), json!("=RC[-1]")]);

        let out = audit("wb".into(), "s".into(), "B1:D4", &a1, &r1c1, 50);
        assert_eq!(out.inconsistent_rows_total, 4);
        let rows: Vec<_> = out.inconsistent_rows.iter().map(|f| (f.row, f.same_rows.as_deref())).collect();
        assert_eq!(rows, [(1, Some("2-3")), (4, None)]);
        assert_eq!(out.embedded_numbers_total, 2);
        assert!(!out.truncated);

        let out = audit("wb".into(), "s".into(), "B1:D4", &a1, &r1c1, 1);
        assert_eq!((out.inconsistent_rows.len(), out.embedded_numbers.len()), (1, 1));
        assert_eq!((out.inconsistent_rows_total, out.embedded_numbers_total), (4, 2));
        assert!(out.truncated && out.note.is_some());
    }

    #[test]
    fn row_lists_compact() {
        assert_eq!(row_list(&[3, 4, 5, 9, 11, 12]), "3-5, 9, 11-12");
    }
}
