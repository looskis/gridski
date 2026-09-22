//! Cell formatting as tool output: rectangles of identically formatted cells, listing
//! only what differs from Excel's defaults.

use std::collections::HashMap;

use serde::{Deserialize, Serialize};

use crate::cells::cell_address;

#[derive(Debug, Clone, Default, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub struct Borders {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub top: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub bottom: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub left: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub right: Option<String>,
}

impl Borders {
    fn is_empty(&self) -> bool {
        self == &Borders::default()
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Hash, Deserialize)]
pub struct Format {
    pub font_color: String,
    pub bold: bool,
    pub italic: bool,
    pub fill: Option<String>,
    pub number_format: String,
    #[serde(default)]
    pub borders: Borders,
}

impl Format {
    fn is_default(&self) -> bool {
        self.font_color == "#000000"
            && !self.bold
            && !self.italic
            && self.fill.is_none()
            && self.number_format == "General"
            && self.borders.is_empty()
    }
}

/// One rectangle of cells sharing a format. Fields at their default are omitted.
#[derive(Debug, Serialize, PartialEq)]
pub struct Block {
    pub range: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub font_color: Option<String>,
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub bold: bool,
    #[serde(skip_serializing_if = "std::ops::Not::not")]
    pub italic: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub fill: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub number_format: Option<String>,
    #[serde(skip_serializing_if = "Borders::is_empty")]
    pub borders: Borders,
}

impl Block {
    fn new(top: u32, left: u32, bottom: u32, right: u32, f: Format) -> Self {
        let tl = cell_address(top, left);
        let range = if (top, left) == (bottom, right) { tl } else { format!("{tl}:{}", cell_address(bottom, right)) };
        Block {
            range,
            font_color: (f.font_color != "#000000").then_some(f.font_color),
            bold: f.bold,
            italic: f.italic,
            fill: f.fill,
            number_format: (f.number_format != "General").then_some(f.number_format),
            borders: f.borders,
        }
    }
}

#[derive(Debug, Serialize)]
pub struct FormatsOutput {
    pub workbook: String,
    pub sheet: String,
    pub address: String,
    /// `file` when read from the saved workbook, `live` when queried from Excel.
    pub source: &'static str,
    pub blocks: Vec<Block>,
    /// Ranges the live reader ran out of time for.
    #[serde(skip_serializing_if = "Vec::is_empty")]
    pub unread: Vec<String>,
}

/// Merge per-cell formats into rectangles: runs along each row, then identical runs
/// stacked in consecutive rows. Default-formatted cells are dropped.
pub fn blocks_from_cells(mut cells: Vec<(u32, u32, Format)>) -> Vec<Block> {
    cells.retain(|(_, _, f)| !f.is_default());
    cells.sort_by_key(|&(r, c, _)| (r, c));

    // Row runs: (row, left, right, format).
    let mut runs: Vec<(u32, u32, u32, Format)> = Vec::new();
    for (row, col, f) in cells {
        match runs.last_mut() {
            Some((r, _, right, last)) if *r == row && *right + 1 == col && *last == f => *right = col,
            _ => runs.push((row, col, col, f)),
        }
    }

    // Stack runs: key (left, right, format) → index of the open rectangle.
    let mut rects: Vec<(u32, u32, u32, u32, Format)> = Vec::new();
    let mut open: HashMap<(u32, u32, Format), usize> = HashMap::new();
    for (row, left, right, f) in runs {
        let key = (left, right, f);
        match open.get(&key) {
            Some(&i) if rects[i].2 + 1 == row => rects[i].2 = row,
            _ => {
                open.insert(key.clone(), rects.len());
                rects.push((row, left, row, right, key.2));
            }
        }
    }
    rects.into_iter().map(|(t, l, b, r, f)| Block::new(t, l, b, r, f)).collect()
}

/// A rectangle the live reader found uniform (see `read_formats.js`).
#[derive(Debug, Clone, Deserialize)]
pub struct LiveBlock {
    pub top: u32,
    pub left: u32,
    pub bottom: u32,
    pub right: u32,
    pub font_color: Option<String>,
    pub bold: bool,
    pub italic: bool,
    pub fill: Option<String>,
    pub number_format: Option<String>,
}

pub fn blocks_from_live(live: Vec<LiveBlock>) -> Vec<Block> {
    let mut cells = Vec::new();
    for b in live {
        let f = Format {
            font_color: b.font_color.unwrap_or_else(|| "#000000".into()),
            bold: b.bold,
            italic: b.italic,
            fill: b.fill.filter(|c| c != "#FFFFFF"),
            number_format: b.number_format.unwrap_or_else(|| "General".into()),
            borders: Borders::default(),
        };
        for r in b.top..=b.bottom {
            for c in b.left..=b.right {
                cells.push((r, c, f.clone()));
            }
        }
    }
    blocks_from_cells(cells)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn fmt(color: &str, bold: bool) -> Format {
        Format {
            font_color: color.into(),
            bold,
            italic: false,
            fill: None,
            number_format: "General".into(),
            borders: Borders::default(),
        }
    }

    #[test]
    fn merges_rows_then_stacks() {
        let blue = fmt("#0000FF", false);
        let cells = vec![
            (2, 3, blue.clone()),
            (2, 4, blue.clone()),
            (3, 3, blue.clone()),
            (3, 4, blue.clone()),
            (4, 2, fmt("#000000", true)),
            (5, 2, fmt("#000000", false)), // default: dropped
        ];
        let blocks = blocks_from_cells(cells);
        assert_eq!(blocks.len(), 2);
        assert_eq!(blocks[0].range, "C2:D3");
        assert_eq!(blocks[0].font_color.as_deref(), Some("#0000FF"));
        assert_eq!(blocks[1].range, "B4");
        assert!(blocks[1].bold && blocks[1].font_color.is_none());
    }

    #[test]
    fn gaps_split_blocks() {
        let blue = fmt("#0000FF", false);
        let blocks = blocks_from_cells(vec![(1, 1, blue.clone()), (1, 3, blue.clone()), (3, 1, blue)]);
        let ranges: Vec<_> = blocks.iter().map(|b| b.range.as_str()).collect();
        assert_eq!(ranges, ["A1", "C1", "A3"]);
    }
}
