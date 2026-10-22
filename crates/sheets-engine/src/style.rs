//! Cell formatting: font, colours, alignment, borders and number format.

pub use crate::borders::{BorderLine, BorderStyle, Borders, Side};

/// Formatting of one cell. `Style::default()` is an unformatted cell.
#[derive(Debug, Clone, Default, PartialEq, Eq, Hash)]
pub struct Style {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strike: bool,
    /// Text colour as `#RRGGBB`.
    pub color: Option<String>,
    /// Background colour as `#RRGGBB`.
    pub fill: Option<String>,
    /// `None` means automatic: numbers right, text left, booleans and errors centred.
    pub align: Option<Align>,
    pub number_format: NumberFormat,
    /// Font family, e.g. "Arial". `None` is the workbook default (Calibri).
    pub font: Option<String>,
    /// Font size in tenths of a point (`110` = 11 pt). `None` is the default, 11 pt.
    pub font_size: Option<u16>,
    /// Wrap text onto several lines inside the cell (Excel's "Wrap Text").
    pub wrap: bool,
    /// Vertical alignment; Excel's default is bottom.
    pub valign: VAlign,
    /// Lines drawn on the cell's edges (see `borders.rs`).
    pub borders: Borders,
}

impl Style {
    pub fn is_default(&self) -> bool {
        *self == Style::default()
    }

    /// A copy with one toolbar change applied.
    pub fn patched(&self, patch: &StylePatch) -> Style {
        let mut s = self.clone();
        match patch {
            StylePatch::Bold(on) => s.bold = *on,
            StylePatch::Italic(on) => s.italic = *on,
            StylePatch::Underline(on) => s.underline = *on,
            StylePatch::Strike(on) => s.strike = *on,
            StylePatch::Color(c) => s.color = c.clone(),
            StylePatch::Fill(c) => s.fill = c.clone(),
            StylePatch::Align(a) => s.align = *a,
            StylePatch::NumberFormat(nf) => s.number_format = nf.clone(),
            StylePatch::Decimals(delta) => s.number_format = s.number_format.with_decimals_changed(*delta),
            StylePatch::Font(name) => s.font = name.clone().filter(|n| !n.is_empty() && n != DEFAULT_FONT),
            StylePatch::FontSize(size) => s.font_size = size.map(clamp_font_size).filter(|t| *t != DEFAULT_FONT_SIZE),
            StylePatch::FontSizeStep(up) => {
                let next = step_font_size(s.font_size.unwrap_or(DEFAULT_FONT_SIZE), *up);
                s.font_size = Some(next).filter(|t| *t != DEFAULT_FONT_SIZE);
            }
            StylePatch::Wrap(on) => s.wrap = *on,
            StylePatch::VAlign(v) => s.valign = *v,
            StylePatch::Clear => s = Style::default(),
        }
        s
    }
}

/// One formatting change from the toolbar, applied to every selected cell.
#[derive(Debug, Clone, PartialEq)]
pub enum StylePatch {
    Bold(bool),
    Italic(bool),
    Underline(bool),
    Strike(bool),
    Color(Option<String>),
    Fill(Option<String>),
    Align(Option<Align>),
    NumberFormat(NumberFormat),
    /// Add or remove decimal places (the ".0" / ".00" buttons).
    Decimals(i8),
    /// Font family; `None` resets to the default font.
    Font(Option<String>),
    /// Font size in tenths of a point; `None` resets to 11 pt.
    FontSize(Option<u16>),
    /// Excel's "Increase Font Size" (`true`) / "Decrease Font Size" (`false`) buttons.
    FontSizeStep(bool),
    Wrap(bool),
    VAlign(VAlign),
    Clear,
}

/// Excel's default font for new workbooks.
pub const DEFAULT_FONT: &str = "Calibri";
/// 11 pt, in tenths of a point.
pub const DEFAULT_FONT_SIZE: u16 = 110;
/// The sizes Excel's font-size box lists, in tenths of a point.
pub const FONT_SIZES: [u16; 17] = [80, 90, 100, 105, 110, 120, 140, 160, 180, 200, 220, 240, 260, 280, 360, 480, 720];

/// Excel accepts font sizes from 1 to 409 pt.
fn clamp_font_size(tenths: u16) -> u16 {
    tenths.clamp(10, 4090)
}

/// The next size up or down, like Excel's Increase/Decrease Font Size buttons:
/// through the standard list (without 10.5), then in 10 pt steps above 72 and
/// 1 pt steps below 8.
pub fn step_font_size(tenths: u16, up: bool) -> u16 {
    let steps = || FONT_SIZES.iter().copied().filter(|t| *t != 105);
    let next = if up {
        steps().find(|t| *t > tenths).unwrap_or((tenths / 100 + 1) * 100)
    } else if tenths > 720 {
        (tenths.div_ceil(100) * 100 - 100).max(720)
    } else {
        steps().rfind(|t| *t < tenths).unwrap_or(tenths.div_ceil(10).saturating_sub(1) * 10)
    };
    clamp_font_size(next)
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Hash)]
pub enum VAlign {
    Top,
    Center,
    #[default]
    Bottom,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Align {
    Left,
    Center,
    Right,
}

#[derive(Debug, Clone, Default, PartialEq, Eq, Hash)]
pub enum NumberFormat {
    #[default]
    General,
    /// `1234.5` -> `1,234.50` with `decimals: 2, thousands: true`.
    Number { decimals: u8, thousands: bool },
    /// `₹12,34,567.00`. The rupee uses Indian digit grouping; other symbols use groups of three.
    Currency { symbol: String, decimals: u8 },
    /// `0.125` -> `12.5%` with `decimals: 1`.
    Percent { decimals: u8 },
    /// `dd-mm-yyyy`
    Date,
    /// `hh:mm`
    Time,
    /// `dd-mm-yyyy hh:mm`
    DateTime,
    /// Show the input as typed, never as a number.
    Text,
    /// Any other Excel format code, e.g. `dd-mmm-yyyy`, `#,##0.0,"K"`, `[Red]0.00` (see `numfmt.rs`).
    Custom(String),
}

impl NumberFormat {
    fn with_decimals_changed(&self, delta: i8) -> NumberFormat {
        let change = |d: u8| (d as i16 + delta as i16).clamp(0, 10) as u8;
        match self {
            NumberFormat::General => NumberFormat::Number { decimals: change(0), thousands: false },
            NumberFormat::Number { decimals, thousands } => {
                NumberFormat::Number { decimals: change(*decimals), thousands: *thousands }
            }
            NumberFormat::Currency { symbol, decimals } => {
                NumberFormat::Currency { symbol: symbol.clone(), decimals: change(*decimals) }
            }
            NumberFormat::Percent { decimals } => NumberFormat::Percent { decimals: change(*decimals) },
            NumberFormat::Custom(code) => NumberFormat::from_code(&custom_decimals(code, delta)),
            other => other.clone(),
        }
    }

    /// Excel format code, used when saving `.xlsx`.
    pub fn to_code(&self) -> String {
        let decimals = |d: u8| if d == 0 { String::new() } else { format!(".{}", "0".repeat(d as usize)) };
        match self {
            NumberFormat::General => "General".into(),
            NumberFormat::Number { decimals: d, thousands: true } => format!("#,##0{}", decimals(*d)),
            NumberFormat::Number { decimals: d, thousands: false } => format!("0{}", decimals(*d)),
            NumberFormat::Currency { symbol, decimals: d } if symbol == "₹" => {
                format!("[$₹-4009]#,##,##0{}", decimals(*d))
            }
            NumberFormat::Currency { symbol, decimals: d } => format!("\"{symbol}\"#,##0{}", decimals(*d)),
            NumberFormat::Percent { decimals: d } => format!("0{}%", decimals(*d)),
            NumberFormat::Date => "dd-mm-yyyy".into(),
            NumberFormat::Time => "hh:mm".into(),
            NumberFormat::DateTime => "dd-mm-yyyy hh:mm".into(),
            NumberFormat::Text => "@".into(),
            NumberFormat::Custom(code) => code.clone(),
        }
    }

    /// The format for an Excel format code (from `.xlsx` files or the Format
    /// Cells dialog). Codes our toolbar formats write map back to them; Excel's
    /// locale-dependent short date (built-in format 14) is our date; anything
    /// else is kept exactly as `Custom`.
    pub fn from_code(code: &str) -> NumberFormat {
        let code = code.trim();
        if code.is_empty() || code.eq_ignore_ascii_case("general") {
            return NumberFormat::General;
        }
        if code == "@" {
            return NumberFormat::Text;
        }
        if matches!(
            code.to_ascii_lowercase().as_str(),
            "mm-dd-yy" | "m/d/yyyy" | "m/d/yy" | "dd-mm-yyyy" | "dd/mm/yyyy"
        ) {
            return NumberFormat::Date;
        }
        let simple = [
            NumberFormat::Date,
            NumberFormat::Time,
            NumberFormat::DateTime,
            NumberFormat::Percent { decimals: 0 },
            NumberFormat::Percent { decimals: 1 },
            NumberFormat::Percent { decimals: 2 },
        ];
        if let Some(f) = simple.into_iter().find(|f| f.to_code() == code) {
            return f;
        }
        for decimals in 0..=10u8 {
            for f in [
                NumberFormat::Number { decimals, thousands: true },
                NumberFormat::Number { decimals, thousands: false },
                NumberFormat::Currency { symbol: "₹".into(), decimals },
                NumberFormat::Currency { symbol: "$".into(), decimals },
                NumberFormat::Currency { symbol: "€".into(), decimals },
                NumberFormat::Currency { symbol: "£".into(), decimals },
                NumberFormat::Currency { symbol: "¥".into(), decimals },
                NumberFormat::Percent { decimals },
            ] {
                if f.to_code() == code {
                    return f;
                }
            }
        }
        NumberFormat::Custom(code.to_string())
    }
}

/// A custom code with one more (`delta > 0`) or one fewer decimal place in each
/// numeric section: `#,##0` -> `#,##0.0`, `0.00"kg"` -> `0.0"kg"`. Date codes are left alone.
fn custom_decimals(code: &str, delta: i8) -> String {
    let lower = code.to_ascii_lowercase();
    if ["y", "d", "h", "s", "mm"].iter().any(|t| lower.replace("\"", "").contains(t)) && !lower.contains('0') {
        return code.to_string();
    }
    let sections: Vec<String> = split_unquoted(code, ';')
        .into_iter()
        .map(|sec| {
            let chars: Vec<char> = sec.chars().collect();
            // Last digit placeholder outside quotes.
            let mut quote = false;
            let mut last_digit = None;
            let mut point = None;
            for (i, c) in chars.iter().enumerate() {
                match c {
                    '"' => quote = !quote,
                    '0' | '#' | '?' if !quote => last_digit = Some(i),
                    '.' if !quote => point = Some(i),
                    _ => {}
                }
            }
            let Some(last) = last_digit else { return sec };
            let mut out = chars.clone();
            if delta > 0 {
                if point.is_some() {
                    out.insert(last + 1, '0');
                } else {
                    out.splice(last + 1..last + 1, ['.', '0']);
                }
            } else if let Some(p) = point {
                if last > p {
                    out.remove(last);
                    if last == p + 1 {
                        out.remove(p);
                    }
                }
            }
            out.into_iter().collect()
        })
        .collect();
    sections.join(";")
}

fn split_unquoted(s: &str, sep: char) -> Vec<String> {
    let mut out = vec![String::new()];
    let mut quote = false;
    for c in s.chars() {
        if c == '"' {
            quote = !quote;
        }
        if c == sep && !quote {
            out.push(String::new());
        } else {
            out.last_mut().unwrap().push(c);
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn patches() {
        let s = Style::default().patched(&StylePatch::Bold(true)).patched(&StylePatch::Decimals(2));
        assert!(s.bold);
        assert_eq!(s.number_format, NumberFormat::Number { decimals: 2, thousands: false });
        let s = s.patched(&StylePatch::Decimals(-5));
        assert_eq!(s.number_format, NumberFormat::Number { decimals: 0, thousands: false });
        assert!(s.patched(&StylePatch::Clear).is_default());
    }

    #[test]
    fn font_patches_normalise_defaults() {
        let s =
            Style::default().patched(&StylePatch::Font(Some("Arial".into()))).patched(&StylePatch::FontSize(Some(140)));
        assert_eq!(s.font.as_deref(), Some("Arial"));
        assert_eq!(s.font_size, Some(140));
        // Choosing Calibri 11 again is the same as the default.
        let s = s.patched(&StylePatch::Font(Some("Calibri".into()))).patched(&StylePatch::FontSize(Some(110)));
        assert!(s.is_default());
        assert_eq!(Style::default().patched(&StylePatch::FontSize(Some(5000))).font_size, Some(4090));
    }

    #[test]
    fn font_size_steps_like_excel() {
        let up = |t| step_font_size(t, true);
        let down = |t| step_font_size(t, false);
        assert_eq!(up(110), 120);
        assert_eq!(up(100), 110); // skips 10.5
        assert_eq!(up(105), 110);
        assert_eq!(up(130), 140);
        assert_eq!(up(720), 800);
        assert_eq!(up(800), 900);
        assert_eq!(down(110), 100);
        assert_eq!(down(80), 70);
        assert_eq!(down(75), 70);
        assert_eq!(down(10), 10);
        assert_eq!(down(900), 800);
        assert_eq!(down(800), 720);
        assert_eq!(down(750), 720);
        let s = Style::default().patched(&StylePatch::FontSizeStep(true));
        assert_eq!(s.font_size, Some(120));
        assert!(s.patched(&StylePatch::FontSizeStep(false)).is_default());
    }

    #[test]
    fn wrap_and_vertical_alignment() {
        let s = Style::default().patched(&StylePatch::Wrap(true)).patched(&StylePatch::VAlign(VAlign::Top));
        assert!(s.wrap);
        assert_eq!(s.valign, VAlign::Top);
        let s = s.patched(&StylePatch::Wrap(false)).patched(&StylePatch::VAlign(VAlign::Bottom));
        assert!(s.is_default());
    }

    #[test]
    fn format_codes_round_trip() {
        let formats = [
            NumberFormat::General,
            NumberFormat::Number { decimals: 2, thousands: true },
            NumberFormat::Number { decimals: 0, thousands: false },
            NumberFormat::Currency { symbol: "₹".into(), decimals: 2 },
            NumberFormat::Currency { symbol: "$".into(), decimals: 0 },
            NumberFormat::Percent { decimals: 1 },
            NumberFormat::Date,
            NumberFormat::Time,
            NumberFormat::DateTime,
            NumberFormat::Text,
        ];
        for f in formats {
            assert_eq!(NumberFormat::from_code(&f.to_code()), f, "{}", f.to_code());
        }
        assert_eq!(NumberFormat::from_code("m/d/yy"), NumberFormat::Date);
        // Everything else is kept exactly, so it displays (and saves) as in Excel.
        for code in ["mm-dd-yy h:mm", "#,##0.00_);(#,##0.00)", "dd-mmm-yyyy", "[Red]0.0"] {
            assert_eq!(NumberFormat::from_code(code), NumberFormat::Custom(code.into()));
        }
    }

    #[test]
    fn decimals_change_in_custom_codes() {
        let step = |code: &str, d: i8| NumberFormat::Custom(code.into()).with_decimals_changed(d).to_code();
        assert_eq!(step("#,##0\" kg\"", 1), "#,##0.0\" kg\"");
        assert_eq!(step("0.00;[Red]-0.00", -1), "0.0;[Red]-0.0");
        assert_eq!(step("0.0\"x\"", -1), "0\"x\"");
        assert_eq!(step("dd-mmm-yyyy", 1), "dd-mmm-yyyy");
    }
}
