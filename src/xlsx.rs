//! Reading cell formatting straight from a saved `.xlsx`/`.xlsm` file.
//!
//! Asking Excel for formats costs one Apple Event per property per range, which is far
//! too slow for a whole model. The file on disk has every cell's style in a few XML
//! parts, so when the workbook has no unsaved changes we read those instead.

use std::collections::HashMap;
use std::fs::File;
use std::io::Read;

use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};

use crate::cells::top_left;
use crate::formats::{Borders, Format};

#[derive(Debug, thiserror::Error)]
pub enum XlsxError {
    #[error("could not read {0}: {1}")]
    Io(String, String),
    #[error("{0} is not a readable workbook file: {1}")]
    Format(String, String),
    #[error("no sheet named \"{0}\" in the saved file")]
    NoSheet(String),
}

/// Every styled or non-empty cell in `sheet` inside `bounds` (top, left, bottom, right),
/// as (row, column, format).
pub fn read_cell_formats(
    path: &str,
    sheet: &str,
    bounds: (u32, u32, u32, u32),
) -> Result<Vec<(u32, u32, Format)>, XlsxError> {
    let file = File::open(path).map_err(|e| XlsxError::Io(path.into(), e.to_string()))?;
    let mut zip = zip::ZipArchive::new(file).map_err(|e| XlsxError::Format(path.into(), e.to_string()))?;
    let mut part = |name: &str| -> Result<Option<String>, XlsxError> {
        let mut entry = match zip.by_name(name) {
            Ok(entry) => entry,
            Err(zip::result::ZipError::FileNotFound) => return Ok(None),
            Err(e) => return Err(XlsxError::Format(path.into(), e.to_string())),
        };
        let mut text = String::new();
        entry.read_to_string(&mut text).map_err(|e| XlsxError::Format(path.into(), e.to_string()))?;
        Ok(Some(text))
    };
    let bad = |what: &str| XlsxError::Format(path.into(), format!("missing {what}"));

    let workbook = parse(&part("xl/workbook.xml")?.ok_or_else(|| bad("xl/workbook.xml"))?);
    let rels = parse(&part("xl/_rels/workbook.xml.rels")?.ok_or_else(|| bad("workbook relationships"))?);
    let rel_id = workbook
        .find_all("sheet")
        .find(|s| s.attr("name") == Some(sheet))
        .and_then(|s| s.attr("id"))
        .ok_or_else(|| XlsxError::NoSheet(sheet.into()))?;
    let target = rels
        .find_all("Relationship")
        .find(|r| r.attr("Id") == Some(rel_id))
        .and_then(|r| r.attr("Target"))
        .ok_or_else(|| bad("sheet relationship"))?;
    let sheet_part = match target.strip_prefix('/') {
        Some(absolute) => absolute.to_string(),
        None => format!("xl/{target}"),
    };

    let theme = part("xl/theme/theme1.xml")?.map(|t| theme_colors(&parse(&t))).unwrap_or_default();
    let styles = Styles::parse(&parse(&part("xl/styles.xml")?.ok_or_else(|| bad("xl/styles.xml"))?), &theme);
    let sheet_xml = parse(&part(&sheet_part)?.ok_or_else(|| bad(&sheet_part))?);

    let (top, left, bottom, right) = bounds;
    let mut cells = Vec::new();
    for c in sheet_xml.find_all("c") {
        let Some((row, col)) = c.attr("r").and_then(top_left) else { continue };
        if row < top || row > bottom || col < left || col > right {
            continue;
        }
        let xf = c.attr("s").and_then(|s| s.parse().ok()).unwrap_or(0);
        cells.push((row, col, styles.format(xf)));
    }
    Ok(cells)
}

/// A parsed XML element; only names, attributes, and children matter for these parts.
#[derive(Debug, Default)]
struct Node {
    name: String,
    attrs: Vec<(String, String)>,
    children: Vec<Node>,
}

impl Node {
    fn attr(&self, key: &str) -> Option<&str> {
        self.attrs.iter().find(|(k, _)| k == key).map(|(_, v)| v.as_str())
    }

    fn child(&self, name: &str) -> Option<&Node> {
        self.children.iter().find(|c| c.name == name)
    }

    /// All descendants with this local name, in document order.
    fn find_all<'a>(&'a self, name: &'a str) -> impl Iterator<Item = &'a Node> + 'a {
        let mut stack: Vec<&Node> = vec![self];
        std::iter::from_fn(move || {
            while let Some(node) = stack.pop() {
                stack.extend(node.children.iter().rev());
                if node.name == name {
                    return Some(node);
                }
            }
            None
        })
    }
}

/// Parse into a tree keyed by local names (namespace prefixes dropped). Malformed
/// trailing content just ends the parse; these parts are machine-written.
fn parse(xml: &str) -> Node {
    fn element(e: &BytesStart) -> Node {
        let attrs = e
            .attributes()
            .flatten()
            .map(|a| {
                let key = a.key.local_name().as_ref().to_owned();
                let value = a.normalized_value(XmlVersion::Implicit1_0).map(|v| v.into_owned()).unwrap_or_default();
                (key, value)
            })
            .collect();
        Node { name: e.local_name().as_ref().to_owned(), attrs, children: Vec::new() }
    }

    let mut reader = quick_xml::Reader::from_str(xml);
    let mut stack = vec![Node::default()];
    loop {
        match reader.read_event() {
            Ok(Event::Start(e)) => stack.push(element(&e)),
            Ok(Event::Empty(e)) => {
                let node = element(&e);
                stack.last_mut().expect("root").children.push(node);
            }
            Ok(Event::End(_)) if stack.len() > 1 => {
                let node = stack.pop().expect("open element");
                stack.last_mut().expect("root").children.push(node);
            }
            Ok(Event::Eof) | Err(_) => break,
            _ => {}
        }
    }
    while stack.len() > 1 {
        let node = stack.pop().expect("open element");
        stack.last_mut().expect("root").children.push(node);
    }
    stack.pop().expect("root")
}

struct Font {
    color: String,
    bold: bool,
    italic: bool,
}

struct Styles {
    /// Per cellXfs entry: (numFmtId, fontId, fillId, borderId).
    xfs: Vec<(u32, usize, usize, usize)>,
    num_fmts: HashMap<u32, String>,
    fonts: Vec<Font>,
    fills: Vec<Option<String>>,
    borders: Vec<Borders>,
}

impl Styles {
    fn parse(styles: &Node, theme: &[String]) -> Self {
        let num_fmts = styles
            .find_all("numFmt")
            .filter_map(|n| Some((n.attr("numFmtId")?.parse().ok()?, n.attr("formatCode")?.to_string())))
            .collect();
        let sheet = styles.child("styleSheet").unwrap_or(styles);
        let section = |name: &str| sheet.child(name).map(|s| s.children.as_slice()).unwrap_or_default();
        let fonts = section("fonts")
            .iter()
            .map(|f| Font {
                color: f.child("color").and_then(|c| color(c, theme)).unwrap_or_else(|| "#000000".into()),
                bold: f.child("b").is_some_and(flag),
                italic: f.child("i").is_some_and(flag),
            })
            .collect();
        let fills = section("fills")
            .iter()
            .map(|f| {
                let pattern = f.child("patternFill")?;
                match pattern.attr("patternType") {
                    None | Some("none") | Some("gray125") => None,
                    // White fills just hide gridlines; treat them as no fill.
                    Some(_) => pattern.child("fgColor").and_then(|c| color(c, theme)).filter(|c| c != "#FFFFFF"),
                }
            })
            .collect();
        let borders = section("borders")
            .iter()
            .map(|b| {
                let edge = |name: &str| {
                    b.child(name).and_then(|e| e.attr("style")).filter(|s| *s != "none").map(str::to_string)
                };
                Borders { top: edge("top"), bottom: edge("bottom"), left: edge("left"), right: edge("right") }
            })
            .collect();
        let xfs = section("cellXfs")
            .iter()
            .map(|x| {
                let n = |key: &str| x.attr(key).and_then(|v| v.parse().ok()).unwrap_or(0);
                (n("numFmtId") as u32, n("fontId"), n("fillId"), n("borderId"))
            })
            .collect();
        Styles { xfs, num_fmts, fonts, fills, borders }
    }

    fn format(&self, xf: usize) -> Format {
        let (num_fmt, font, fill, border) = self.xfs.get(xf).copied().unwrap_or_default();
        let font = self.fonts.get(font);
        Format {
            font_color: font.map_or_else(|| "#000000".into(), |f| f.color.clone()),
            bold: font.is_some_and(|f| f.bold),
            italic: font.is_some_and(|f| f.italic),
            fill: self.fills.get(fill).cloned().flatten(),
            number_format: self
                .num_fmts
                .get(&num_fmt)
                .cloned()
                .unwrap_or_else(|| builtin_number_format(num_fmt).map_or_else(|| format!("builtin #{num_fmt}"), str::to_string)),
            borders: self.borders.get(border).cloned().unwrap_or_default(),
        }
    }
}

/// `<b/>` means bold; `<b val="0"/>` explicitly turns it off.
fn flag(node: &Node) -> bool {
    !matches!(node.attr("val"), Some("0") | Some("false"))
}

fn builtin_number_format(id: u32) -> Option<&'static str> {
    Some(match id {
        0 => "General",
        1 => "0",
        2 => "0.00",
        3 => "#,##0",
        4 => "#,##0.00",
        9 => "0%",
        10 => "0.00%",
        11 => "0.00E+00",
        12 => "# ?/?",
        13 => "# ??/??",
        14 => "m/d/yyyy",
        15 => "d-mmm-yy",
        16 => "d-mmm",
        17 => "mmm-yy",
        18 => "h:mm AM/PM",
        19 => "h:mm:ss AM/PM",
        20 => "h:mm",
        21 => "h:mm:ss",
        22 => "m/d/yyyy h:mm",
        37 => "#,##0 ;(#,##0)",
        38 => "#,##0 ;[Red](#,##0)",
        39 => "#,##0.00;(#,##0.00)",
        40 => "#,##0.00;[Red](#,##0.00)",
        45 => "mm:ss",
        46 => "[h]:mm:ss",
        47 => "mmss.0",
        48 => "##0.0E+0",
        49 => "@",
        _ => return None,
    })
}

/// Excel's default indexed palette (indices 0–63; 64/65 are system text/background).
const PALETTE: [u32; 64] = [
    0x000000, 0xFFFFFF, 0xFF0000, 0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF, 0x000000, 0xFFFFFF, 0xFF0000,
    0x00FF00, 0x0000FF, 0xFFFF00, 0xFF00FF, 0x00FFFF, 0x800000, 0x008000, 0x000080, 0x808000, 0x800080, 0x008080,
    0xC0C0C0, 0x808080, 0x9999FF, 0x993366, 0xFFFFCC, 0xCCFFFF, 0x660066, 0xFF8080, 0x0066CC, 0xCCCCFF, 0x000080,
    0xFF00FF, 0xFFFF00, 0x00FFFF, 0x800080, 0x800000, 0x008080, 0x0000FF, 0x00CCFF, 0xCCFFFF, 0xCCFFCC, 0xFFFF99,
    0x99CCFF, 0xFF99CC, 0xCC99FF, 0xFFCC99, 0x3366FF, 0x33CCCC, 0x99CC00, 0xFFCC00, 0xFF9900, 0xFF6600, 0x666699,
    0x969696, 0x003366, 0x339966, 0x003300, 0x333300, 0x993300, 0x993366, 0x333399, 0x333333,
];

/// Theme colors in the order `theme="n"` indexes them. The file lists dk1, lt1, dk2,
/// lt2, accents…; Excel's index swaps each dark/light pair (0 = lt1, 1 = dk1, ...).
fn theme_colors(theme: &Node) -> Vec<String> {
    let Some(scheme) = theme.find_all("clrScheme").next() else { return Vec::new() };
    let mut colors: Vec<String> = scheme
        .children
        .iter()
        .map(|c| {
            let hex = c
                .child("srgbClr")
                .and_then(|s| s.attr("val"))
                .or_else(|| c.child("sysClr").and_then(|s| s.attr("lastClr")))
                .unwrap_or("000000");
            format!("#{}", hex.to_ascii_uppercase())
        })
        .collect();
    if colors.len() >= 4 {
        colors.swap(0, 1);
        colors.swap(2, 3);
    }
    colors
}

/// Resolve a `<color>`-style element to `#RRGGBB`; `None` for automatic.
fn color(node: &Node, theme: &[String]) -> Option<String> {
    let base = if let Some(rgb) = node.attr("rgb") {
        let hex = &rgb[rgb.len().saturating_sub(6)..];
        format!("#{}", hex.to_ascii_uppercase())
    } else if let Some(i) = node.attr("indexed").and_then(|i| i.parse::<usize>().ok()) {
        match i {
            64 => "#000000".into(),
            65 => "#FFFFFF".into(),
            _ => format!("#{:06X}", PALETTE.get(i)?),
        }
    } else if let Some(t) = node.attr("theme").and_then(|t| t.parse::<usize>().ok()) {
        theme.get(t)?.clone()
    } else {
        return None;
    };
    match node.attr("tint").and_then(|t| t.parse::<f64>().ok()) {
        Some(tint) if tint != 0.0 => Some(apply_tint(&base, tint)),
        _ => Some(base),
    }
}

/// Excel's tint: lighten (tint > 0) or darken (tint < 0) the color's HSL lightness.
fn apply_tint(hex: &str, tint: f64) -> String {
    let n = u32::from_str_radix(hex.trim_start_matches('#'), 16).unwrap_or(0);
    let [r, g, b] = [(n >> 16) & 255, (n >> 8) & 255, n & 255].map(|c| f64::from(c) / 255.0);
    let (max, min) = (r.max(g).max(b), r.min(g).min(b));
    let l = (max + min) / 2.0;
    let d = max - min;
    let s = if d == 0.0 { 0.0 } else { d / (1.0 - (2.0 * l - 1.0).abs()) };
    let h = if d == 0.0 {
        0.0
    } else if max == r {
        60.0 * (((g - b) / d).rem_euclid(6.0))
    } else if max == g {
        60.0 * ((b - r) / d + 2.0)
    } else {
        60.0 * ((r - g) / d + 4.0)
    };
    let l = if tint < 0.0 { l * (1.0 + tint) } else { l * (1.0 - tint) + tint };
    let c = (1.0 - (2.0 * l - 1.0).abs()) * s;
    let x = c * (1.0 - ((h / 60.0).rem_euclid(2.0) - 1.0).abs());
    let m = l - c / 2.0;
    let (r, g, b) = match h as u32 / 60 {
        0 => (c, x, 0.0),
        1 => (x, c, 0.0),
        2 => (0.0, c, x),
        3 => (0.0, x, c),
        4 => (x, 0.0, c),
        _ => (c, 0.0, x),
    };
    let byte = |v: f64| ((v + m) * 255.0).round().clamp(0.0, 255.0) as u32;
    format!("#{:02X}{:02X}{:02X}", byte(r), byte(g), byte(b))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn colors_resolve() {
        let theme = vec!["#FFFFFF".into(), "#000000".into(), "#E7E6E6".into(), "#44546A".into(), "#4472C4".into()];
        let node = |attrs: &[(&str, &str)]| Node {
            name: "color".into(),
            attrs: attrs.iter().map(|(k, v)| (k.to_string(), v.to_string())).collect(),
            children: vec![],
        };
        assert_eq!(color(&node(&[("rgb", "FF0000FF")]), &theme).as_deref(), Some("#0000FF"));
        assert_eq!(color(&node(&[("indexed", "12")]), &theme).as_deref(), Some("#0000FF"));
        assert_eq!(color(&node(&[("theme", "1")]), &theme).as_deref(), Some("#000000"));
        assert_eq!(color(&node(&[("auto", "1")]), &theme), None);
        // White darkened by 15% is Excel's familiar light gray.
        assert_eq!(apply_tint("#FFFFFF", -0.1499984740745262), "#D9D9D9");
    }

    #[test]
    fn parses_nested_xml() {
        let root = parse(r#"<a xmlns:x="u"><x:b n="1"><c/></x:b><b n="2"/></a>"#);
        let found: Vec<_> = root.find_all("b").map(|b| b.attr("n").unwrap()).collect();
        assert_eq!(found, ["1", "2"]);
        assert_eq!(root.find_all("c").count(), 1);
    }
}
