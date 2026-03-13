//! [`Theme`] — workbook-level color scheme extracted from `xl/theme/theme1.xml`.

use std::collections::HashMap;

use serde::{Deserialize, Serialize};

use crate::model::color::{Rgb, ThemeColorSlot};

// ── Theme ─────────────────────────────────────────────────────────────────────

/// Workbook theme — holds the 12 standard color slots.
///
/// Built by [`crate::openxml::theme_parser`].  Passed to color resolvers so
/// that `<a:schemeClr>` references can be turned into concrete [`Rgb`] values.
#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct Theme {
    /// Map from slot → resolved base RGB.
    slots: HashMap<String, Rgb>,

    /// Theme name from `<a:clrScheme name="…"/>`.
    pub name: Option<String>,
}

impl Theme {
    /// Insert a slot → RGB mapping.
    pub(crate) fn set(&mut self, slot: ThemeColorSlot, rgb: Rgb) {
        self.slots.insert(slot.as_str().to_owned(), rgb);
    }

    /// Look up a slot by [`ThemeColorSlot`] enum value.
    pub fn color(&self, slot: ThemeColorSlot) -> Option<Rgb> {
        self.slots.get(slot.as_str()).copied()
    }

    /// Look up a slot by its string name (e.g. `"accent1"`).
    pub fn color_by_name(&self, name: &str) -> Option<Rgb> {
        ThemeColorSlot::from_str(name).and_then(|s| self.color(s))
    }

    // ── Convenience accessors ─────────────────────────────────────────────────

    pub fn accent1(&self) -> Option<Rgb> {
        self.color(ThemeColorSlot::Accent1)
    }
    pub fn accent2(&self) -> Option<Rgb> {
        self.color(ThemeColorSlot::Accent2)
    }
    pub fn accent3(&self) -> Option<Rgb> {
        self.color(ThemeColorSlot::Accent3)
    }
    pub fn accent4(&self) -> Option<Rgb> {
        self.color(ThemeColorSlot::Accent4)
    }
    pub fn accent5(&self) -> Option<Rgb> {
        self.color(ThemeColorSlot::Accent5)
    }
    pub fn accent6(&self) -> Option<Rgb> {
        self.color(ThemeColorSlot::Accent6)
    }
    pub fn dk1(&self) -> Option<Rgb> {
        self.color(ThemeColorSlot::Dk1)
    }
    pub fn lt1(&self) -> Option<Rgb> {
        self.color(ThemeColorSlot::Lt1)
    }
    pub fn dk2(&self) -> Option<Rgb> {
        self.color(ThemeColorSlot::Dk2)
    }
    pub fn lt2(&self) -> Option<Rgb> {
        self.color(ThemeColorSlot::Lt2)
    }

    /// Return all populated slots as `(name, Rgb)` pairs, sorted.
    pub fn all_colors(&self) -> Vec<(&str, Rgb)> {
        let mut v: Vec<(&str, Rgb)> = self.slots.iter().map(|(k, v)| (k.as_str(), *v)).collect();
        v.sort_by_key(|(k, _)| *k);
        v
    }

    /// Returns `true` when no color slots have been populated.
    pub fn is_empty(&self) -> bool {
        self.slots.is_empty()
    }
}

// ── Tests ─────────────────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;

    fn make_theme() -> Theme {
        let mut t = Theme::default();
        t.set(ThemeColorSlot::Accent1, Rgb::from_hex("4472C4").unwrap());
        t.set(ThemeColorSlot::Accent2, Rgb::from_hex("ED7D31").unwrap());
        t.set(ThemeColorSlot::Dk1, Rgb::from_hex("000000").unwrap());
        t.set(ThemeColorSlot::Lt1, Rgb::from_hex("FFFFFF").unwrap());
        t
    }

    #[test]
    fn accent1_lookup() {
        let t = make_theme();
        assert_eq!(t.accent1(), Some(Rgb::from_hex("4472C4").unwrap()));
    }

    #[test]
    fn missing_slot_returns_none() {
        let t = make_theme();
        assert!(t.accent6().is_none());
    }

    #[test]
    fn color_by_name() {
        let t = make_theme();
        assert_eq!(t.color_by_name("accent1"), t.accent1());
    }

    #[test]
    fn unknown_name_returns_none() {
        let t = make_theme();
        assert!(t.color_by_name("notaslot").is_none());
    }

    #[test]
    fn all_colors_sorted() {
        let t = make_theme();
        let names: Vec<_> = t.all_colors().iter().map(|(n, _)| *n).collect();
        let mut sorted = names.clone();
        sorted.sort();
        assert_eq!(names, sorted);
    }

    #[test]
    fn empty_theme() {
        assert!(Theme::default().is_empty());
    }
}
