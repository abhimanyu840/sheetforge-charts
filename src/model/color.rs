//! Color, gradient, and fill types for DrawingML.
//!
//! ## Color resolution pipeline
//!
//! DrawingML colors are specified in one of three ways:
//!
//! 1. **`<a:srgbClr val="RRGGBB"/>`** — already an RGB hex literal.
//! 2. **`<a:sysClr lastClr="RRGGBB"/>`** — system color; `lastClr` is the
//!    fallback RGB observed when the file was saved.
//! 3. **`<a:schemeClr val="accent1"/>`** — a named slot in the workbook
//!    theme.  Requires a [`crate::model::theme::Theme`] to resolve.
//!
//! Color modifiers (`<a:lumMod>`, `<a:lumOff>`, `<a:tint>`, `<a:shade>`)
//! can be chained after any of the above to shift the final value.
//! [`Rgb::apply_mods`] handles this in HSL space.

use serde::{Deserialize, Serialize};

// ── Rgb ───────────────────────────────────────────────────────────────────────

/// A 24-bit sRGB color.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub struct Rgb {
    pub r: u8,
    pub g: u8,
    pub b: u8,
}

impl Rgb {
    pub const BLACK: Self = Self { r: 0, g: 0, b: 0 };
    pub const WHITE: Self = Self {
        r: 255,
        g: 255,
        b: 255,
    };

    /// Parse a 6-character uppercase hex string (e.g. `"4472C4"`).
    pub fn from_hex(s: &str) -> Option<Self> {
        let s = s.trim_start_matches('#');
        if s.len() != 6 {
            return None;
        }
        let r = u8::from_str_radix(&s[0..2], 16).ok()?;
        let g = u8::from_str_radix(&s[2..4], 16).ok()?;
        let b = u8::from_str_radix(&s[4..6], 16).ok()?;
        Some(Self { r, g, b })
    }

    /// Format as `"RRGGBB"` uppercase hex.
    pub fn to_hex(self) -> String {
        format!("{:02X}{:02X}{:02X}", self.r, self.g, self.b)
    }

    /// Apply a chain of [`ColorMod`]s in HSL space.
    ///
    /// DrawingML processes modifiers in document order, each one operating
    /// on the output of the previous.  We convert to HSL, apply, then back.
    pub fn apply_mods(self, mods: &[ColorMod]) -> Self {
        if mods.is_empty() {
            return self;
        }
        let (h, s, mut l) = rgb_to_hsl(self);
        let mut alpha_f = 1.0f64; // not stored in Rgb but tracked so tint/shade compose correctly

        for m in mods {
            match m {
                // lumMod val is in 1/1000ths of a percent (100 000 = 100%)
                ColorMod::LumMod(v) => {
                    l = (l * (*v as f64 / 100_000.0)).clamp(0.0, 1.0);
                }
                ColorMod::LumOff(v) => {
                    l = (l + (*v as f64 / 100_000.0)).clamp(0.0, 1.0);
                }
                // tint: blend toward white  (val/100000 = fraction of tint)
                ColorMod::Tint(v) => {
                    l = l + (1.0 - l) * (*v as f64 / 100_000.0);
                    l = l.clamp(0.0, 1.0);
                }
                // shade: blend toward black
                ColorMod::Shade(v) => {
                    l = l * (*v as f64 / 100_000.0);
                    l = l.clamp(0.0, 1.0);
                }
                ColorMod::Alpha(v) => {
                    alpha_f = *v as f64 / 100_000.0;
                }
                // satMod / satOff — future extension
                ColorMod::SatMod(_) | ColorMod::SatOff(_) => {}
            }
        }
        let _ = alpha_f; // not stored in Rgb, but tracking prevents misuse
        hsl_to_rgb(h, s, l)
    }
}

impl std::fmt::Display for Rgb {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "#{}", self.to_hex())
    }
}

// ── HSL helpers ───────────────────────────────────────────────────────────────

fn rgb_to_hsl(c: Rgb) -> (f64, f64, f64) {
    let r = c.r as f64 / 255.0;
    let g = c.g as f64 / 255.0;
    let b = c.b as f64 / 255.0;
    let max = r.max(g).max(b);
    let min = r.min(g).min(b);
    let l = (max + min) / 2.0;
    if (max - min).abs() < 1e-10 {
        return (0.0, 0.0, l);
    }
    let d = max - min;
    let s = if l > 0.5 {
        d / (2.0 - max - min)
    } else {
        d / (max + min)
    };
    let h = if max == r {
        (g - b) / d + if g < b { 6.0 } else { 0.0 }
    } else if max == g {
        (b - r) / d + 2.0
    } else {
        (r - g) / d + 4.0
    } / 6.0;
    (h, s, l)
}

fn hsl_to_rgb(h: f64, s: f64, l: f64) -> Rgb {
    if s < 1e-10 {
        let v = (l * 255.0).round() as u8;
        return Rgb { r: v, g: v, b: v };
    }
    let q = if l < 0.5 {
        l * (1.0 + s)
    } else {
        l + s - l * s
    };
    let p = 2.0 * l - q;
    let r = hue_to_rgb(p, q, h + 1.0 / 3.0);
    let g = hue_to_rgb(p, q, h);
    let b = hue_to_rgb(p, q, h - 1.0 / 3.0);
    Rgb {
        r: (r * 255.0).round() as u8,
        g: (g * 255.0).round() as u8,
        b: (b * 255.0).round() as u8,
    }
}

fn hue_to_rgb(p: f64, q: f64, mut t: f64) -> f64 {
    if t < 0.0 {
        t += 1.0;
    }
    if t > 1.0 {
        t -= 1.0;
    }
    if t < 1.0 / 6.0 {
        return p + (q - p) * 6.0 * t;
    }
    if t < 1.0 / 2.0 {
        return q;
    }
    if t < 2.0 / 3.0 {
        return p + (q - p) * (2.0 / 3.0 - t) * 6.0;
    }
    p
}

// ── ColorMod ──────────────────────────────────────────────────────────────────

/// A single DrawingML color modifier.  Values use the DrawingML integer scale:
/// `100 000` = 100 %.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub enum ColorMod {
    LumMod(i32),
    LumOff(i32),
    Tint(i32),
    Shade(i32),
    Alpha(i32),
    SatMod(i32),
    SatOff(i32),
}

impl ColorMod {
    /// Parse from a DrawingML modifier element name + `val` attribute.
    pub fn from_tag_val(tag: &str, val: i32) -> Option<Self> {
        match tag {
            "lumMod" => Some(Self::LumMod(val)),
            "lumOff" => Some(Self::LumOff(val)),
            "tint" => Some(Self::Tint(val)),
            "shade" => Some(Self::Shade(val)),
            "alpha" => Some(Self::Alpha(val)),
            "satMod" => Some(Self::SatMod(val)),
            "satOff" => Some(Self::SatOff(val)),
            _ => None,
        }
    }
}

// ── ThemeColorSlot ────────────────────────────────────────────────────────────

/// The 12 standard DrawingML theme color slots.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum ThemeColorSlot {
    Dk1,
    Lt1,
    Dk2,
    Lt2,
    Accent1,
    Accent2,
    Accent3,
    Accent4,
    Accent5,
    Accent6,
    HLink,
    FolHLink,
}

impl ThemeColorSlot {
    pub fn from_str(s: &str) -> Option<Self> {
        match s {
            "dk1" => Some(Self::Dk1),
            "lt1" => Some(Self::Lt1),
            "dk2" => Some(Self::Dk2),
            "lt2" => Some(Self::Lt2),
            "accent1" => Some(Self::Accent1),
            "accent2" => Some(Self::Accent2),
            "accent3" => Some(Self::Accent3),
            "accent4" => Some(Self::Accent4),
            "accent5" => Some(Self::Accent5),
            "accent6" => Some(Self::Accent6),
            "hlink" => Some(Self::HLink),
            "folHlink" => Some(Self::FolHLink),
            _ => None,
        }
    }

    pub fn as_str(self) -> &'static str {
        match self {
            Self::Dk1 => "dk1",
            Self::Lt1 => "lt1",
            Self::Dk2 => "dk2",
            Self::Lt2 => "lt2",
            Self::Accent1 => "accent1",
            Self::Accent2 => "accent2",
            Self::Accent3 => "accent3",
            Self::Accent4 => "accent4",
            Self::Accent5 => "accent5",
            Self::Accent6 => "accent6",
            Self::HLink => "hlink",
            Self::FolHLink => "folHlink",
        }
    }
}

// ── ColorSpec ─────────────────────────────────────────────────────────────────

/// A color as it appears in the XML — before theme resolution.
///
/// Use [`ColorSpec::resolve`] to get a concrete [`Rgb`].
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub enum ColorSpec {
    /// `<a:srgbClr val="RRGGBB"/>` — direct hex.
    Srgb(Rgb, Vec<ColorMod>),

    /// `<a:sysClr lastClr="RRGGBB"/>` — system color, fallback via `lastClr`.
    Sys(Rgb, Vec<ColorMod>),

    /// `<a:schemeClr val="accent1"/>` — requires theme to resolve.
    Scheme(ThemeColorSlot, Vec<ColorMod>),

    /// `<a:prstClr val="black"/>` — preset named color.
    Preset(String, Vec<ColorMod>),
}

impl ColorSpec {
    /// Resolve to a concrete [`Rgb`], optionally using a theme lookup.
    ///
    /// Returns `None` only when the spec is a `Scheme` variant and no theme
    /// was provided (or the slot is absent from the theme).
    pub fn resolve(&self, theme: Option<&crate::model::theme::Theme>) -> Option<Rgb> {
        match self {
            Self::Srgb(rgb, mods) => Some(rgb.apply_mods(mods)),
            Self::Sys(rgb, mods) => Some(rgb.apply_mods(mods)),
            Self::Scheme(slot, mods) => {
                let base = theme?.color(*slot)?;
                Some(base.apply_mods(mods))
            }
            Self::Preset(name, mods) => preset_color(name).map(|rgb| rgb.apply_mods(mods)),
        }
    }

    /// Return the mods slice regardless of variant.
    pub fn mods(&self) -> &[ColorMod] {
        match self {
            Self::Srgb(_, m) | Self::Sys(_, m) | Self::Scheme(_, m) | Self::Preset(_, m) => m,
        }
    }
}

// ── Preset color table (subset) ───────────────────────────────────────────────

fn preset_color(name: &str) -> Option<Rgb> {
    // CSS/SVG color names used in DrawingML prstClr
    let hex = match name {
        "black" => "000000",
        "white" => "FFFFFF",
        "red" => "FF0000",
        "green" => "008000",
        "blue" => "0000FF",
        "yellow" => "FFFF00",
        "cyan" => "00FFFF",
        "magenta" => "FF00FF",
        "orange" => "FFA500",
        "purple" => "800080",
        "gray" => "808080",
        "silver" => "C0C0C0",
        _ => return None,
    };
    Rgb::from_hex(hex)
}

// ── GradientStop ──────────────────────────────────────────────────────────────

/// One `<a:gs>` entry in a gradient stop list.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct GradientStop {
    /// Stop position, 0–100 000 (maps to 0 %–100 %).
    pub position: u32,

    /// Color at this stop.  May reference the theme.
    pub color: ColorSpec,
}

impl GradientStop {
    /// Resolve to `(position_0_to_1, Rgb)`.
    pub fn resolve(&self, theme: Option<&crate::model::theme::Theme>) -> Option<(f64, Rgb)> {
        let rgb = self.color.resolve(theme)?;
        Some((self.position as f64 / 100_000.0, rgb))
    }
}

// ── GradientAngle ─────────────────────────────────────────────────────────────

/// Linear gradient direction or path type.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum GradientDirection {
    /// `<a:lin ang="N" scaled="0|1"/>` — angle in 1/60 000ths of a degree.
    Linear { angle_deg: f64, scaled: bool },
    /// `<a:path path="circle|rect|shape"/>` — radial or shape path.
    Path(String),
}

// ── Gradient ─────────────────────────────────────────────────────────────────

/// A resolved `<a:gradFill>` element.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct Gradient {
    /// Ordered gradient stops.
    pub stops: Vec<GradientStop>,
    /// Direction of the gradient (if `<a:lin>` or `<a:path>` is present).
    pub direction: Option<GradientDirection>,
    /// Whether the fill tiles (`<a:tileRect/>`).
    pub tile: bool,
}

impl Gradient {
    /// Resolve all stops to concrete `(position, Rgb)` pairs.
    pub fn resolve_stops(&self, theme: Option<&crate::model::theme::Theme>) -> Vec<(f64, Rgb)> {
        self.stops.iter().filter_map(|s| s.resolve(theme)).collect()
    }
}

// ── Fill ──────────────────────────────────────────────────────────────────────

/// The fill applied to a chart element (series bar, plot area background, etc.).
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub enum Fill {
    /// `<a:solidFill>` — single flat color.
    Solid(ColorSpec),
    /// `<a:gradFill>` — gradient with 2+ stops.
    Gradient(Gradient),
    /// `<a:noFill/>` — explicitly transparent.
    None,
    /// `<a:pattFill>` — pattern fill (not fully parsed).
    Pattern,
}

impl Fill {
    /// If this is a `Solid` fill, resolve and return the RGB.
    pub fn solid_rgb(&self, theme: Option<&crate::model::theme::Theme>) -> Option<Rgb> {
        match self {
            Self::Solid(spec) => spec.resolve(theme),
            _ => None,
        }
    }
}

// ── Tests ─────────────────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn rgb_from_hex_valid() {
        let rgb = Rgb::from_hex("4472C4").unwrap();
        assert_eq!(
            rgb,
            Rgb {
                r: 0x44,
                g: 0x72,
                b: 0xC4
            }
        );
    }

    #[test]
    fn rgb_from_hex_lowercase() {
        assert!(Rgb::from_hex("4472c4").is_some());
    }

    #[test]
    fn rgb_from_hex_invalid() {
        assert!(Rgb::from_hex("ZZZZZZ").is_none());
        assert!(Rgb::from_hex("123").is_none());
    }

    #[test]
    fn rgb_to_hex_roundtrip() {
        let rgb = Rgb {
            r: 0x44,
            g: 0x72,
            b: 0xC4,
        };
        assert_eq!(rgb.to_hex(), "4472C4");
    }

    #[test]
    fn rgb_display() {
        assert_eq!(
            format!(
                "{}",
                Rgb {
                    r: 0xFF,
                    g: 0x00,
                    b: 0x80
                }
            ),
            "#FF0080"
        );
    }

    #[test]
    fn lum_mod_darkens() {
        // lumMod 75000 = 75% → should darken
        let white = Rgb::WHITE;
        let dark = white.apply_mods(&[ColorMod::LumMod(75_000)]);
        // white with lumMod 75% → 75% luminance → not white
        assert_ne!(dark, Rgb::WHITE);
    }

    #[test]
    fn lum_mod_100000_is_identity() {
        let rgb = Rgb {
            r: 100,
            g: 150,
            b: 200,
        };
        let same = rgb.apply_mods(&[ColorMod::LumMod(100_000)]);
        // Within 1 unit rounding error
        assert!((rgb.r as i16 - same.r as i16).abs() <= 1);
        assert!((rgb.g as i16 - same.g as i16).abs() <= 1);
        assert!((rgb.b as i16 - same.b as i16).abs() <= 1);
    }

    #[test]
    fn no_mods_is_identity() {
        let rgb = Rgb {
            r: 70,
            g: 114,
            b: 196,
        };
        assert_eq!(rgb.apply_mods(&[]), rgb);
    }

    #[test]
    fn theme_color_slot_roundtrip() {
        for s in [
            "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "dk1", "lt1", "dk2",
            "lt2",
        ] {
            let slot = ThemeColorSlot::from_str(s).unwrap();
            assert_eq!(slot.as_str(), s);
        }
    }

    #[test]
    fn color_spec_srgb_resolves_without_theme() {
        let spec = ColorSpec::Srgb(Rgb::from_hex("FF0000").unwrap(), vec![]);
        assert_eq!(spec.resolve(None), Some(Rgb { r: 255, g: 0, b: 0 }));
    }

    #[test]
    fn color_spec_scheme_needs_theme() {
        let spec = ColorSpec::Scheme(ThemeColorSlot::Accent1, vec![]);
        assert_eq!(spec.resolve(None), None);
    }

    #[test]
    fn gradient_stop_position_range() {
        let stop = GradientStop {
            position: 50_000,
            color: ColorSpec::Srgb(Rgb::WHITE, vec![]),
        };
        let (pos, _rgb) = stop.resolve(None).unwrap();
        assert!((pos - 0.5).abs() < 1e-9);
    }

    #[test]
    fn fill_solid_rgb_resolves() {
        let fill = Fill::Solid(ColorSpec::Srgb(Rgb { r: 1, g: 2, b: 3 }, vec![]));
        assert_eq!(fill.solid_rgb(None), Some(Rgb { r: 1, g: 2, b: 3 }));
    }

    #[test]
    fn fill_gradient_solid_returns_none() {
        let fill = Fill::Gradient(Gradient {
            stops: vec![],
            direction: None,
            tile: false,
        });
        assert!(fill.solid_rgb(None).is_none());
    }
}
