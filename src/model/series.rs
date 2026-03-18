//! [`Series`] model — one data series within a chart.

use serde::{Deserialize, Serialize};

use crate::model::color::Fill;

// ── DataReference ─────────────────────────────────────────────────────────────

/// A raw cell-range formula (e.g. `"Sheet1!$B$2:$B$12"`).
/// Range resolution happens in a future phase.
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq)]
pub struct DataReference {
    pub formula: String,
}

// ── DataValues ────────────────────────────────────────────────────────────────

/// Inline numeric cache from `<c:numCache>`.
///
/// Points are stored **in index order** even when the source XML uses sparse
/// `<c:pt idx="N">` addressing (e.g. when some points are hidden or deleted).
/// Gaps are filled with `f64::NAN` so `values[i]` always corresponds to
/// data-point index `i`.
#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct DataValues {
    /// Declared point count from `<c:ptCount val="N"/>`.
    /// `None` when the element is absent (older files).
    pub pt_count: Option<usize>,

    /// Numeric values, index-aligned.  Gaps are `f64::NAN`.
    pub values: Vec<f64>,

    /// Number format string (e.g. `"0.00%"`).
    pub format_code: Option<String>,
}

impl DataValues {
    /// Pre-allocate storage when `<c:ptCount>` is known.
    pub(crate) fn with_capacity(pt_count: usize) -> Self {
        Self {
            pt_count: Some(pt_count),
            values: vec![f64::NAN; pt_count],
            format_code: None,
        }
    }

    /// Insert a value at a specific index.  Extends the vec if needed.
    pub(crate) fn set(&mut self, idx: usize, value: f64) {
        if idx >= self.values.len() {
            self.values.resize(idx + 1, f64::NAN);
        }
        self.values[idx] = value;
    }

    /// Return `true` when every point is present (non-NaN) and count matches.
    pub fn is_complete(&self) -> bool {
        match self.pt_count {
            Some(n) => self.values.len() == n && self.values.iter().all(|v| !v.is_nan()),
            None => !self.values.is_empty(),
        }
    }
}

// ── StringValues ──────────────────────────────────────────────────────────────

/// Inline string cache from `<c:strCache>`.
///
/// Like [`DataValues`], entries are index-aligned.  Missing indices are stored
/// as empty strings.
#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct StringValues {
    /// Declared point count from `<c:ptCount val="N"/>`.
    pub pt_count: Option<usize>,

    /// String values, index-aligned.
    pub values: Vec<String>,
}

impl StringValues {
    pub(crate) fn with_capacity(pt_count: usize) -> Self {
        Self {
            pt_count: Some(pt_count),
            values: vec![String::new(); pt_count],
        }
    }

    pub(crate) fn set(&mut self, idx: usize, value: String) {
        if idx >= self.values.len() {
            self.values.resize(idx + 1, String::new());
        }
        self.values[idx] = value;
    }

    pub fn is_complete(&self) -> bool {
        match self.pt_count {
            Some(n) => self.values.len() == n,
            None => !self.values.is_empty(),
        }
    }
}

// ── CacheState ────────────────────────────────────────────────────────────────

/// Indicates whether cached data is available and can be used instead of
/// resolving worksheet ranges.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "snake_case")]
pub enum CacheState {
    /// No cache present — worksheet must be read.
    #[default]
    None,
    /// Cache is present and complete — safe to use directly.
    Complete,
    /// Cache present but incomplete (sparse gaps or truncated).
    Partial,
}

// ── Series ────────────────────────────────────────────────────────────────────

/// One `<c:ser>` element.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Series {
    /// `<c:idx val="N"/>` — zero-based series index.
    pub index: u32,

    /// `<c:order val="N"/>` — plot order (may differ from index in combo charts).
    pub order: u32,

    /// Series name: literal string from `<c:v>` inside `<c:tx>`.
    pub name: Option<String>,

    /// Formula reference for the series name (e.g. `"Sheet1!$B$1"`).
    pub name_ref: Option<DataReference>,

    // ── Category (X-axis) ─────────────────────────────────────────────────────
    /// Category cell-range formula.
    pub category_ref: Option<DataReference>,

    /// Inline category string cache from `<c:strCache>`.
    pub category_values: Option<StringValues>,

    /// Inline category numeric cache (date axes use `<c:numCache>` for cats).
    pub category_num_cache: Option<DataValues>,

    /// Cache freshness for category data.
    pub category_cache_state: CacheState,

    // ── Values (Y-axis) ───────────────────────────────────────────────────────
    /// Value cell-range formula.
    pub value_ref: Option<DataReference>,

    /// Inline numeric value cache from `<c:numCache>`.
    pub value_cache: Option<DataValues>,

    /// Cache freshness for value data.
    pub value_cache_state: CacheState,

    // ── Scatter / Bubble extras ───────────────────────────────────────────────
    /// X-value formula (scatter / bubble charts use `<c:xVal>`).
    pub x_value_ref: Option<DataReference>,

    /// X-value numeric cache.
    pub x_value_cache: Option<DataValues>,

    /// Bubble-size formula.
    pub bubble_size_ref: Option<DataReference>,

    /// Bubble-size numeric cache.
    pub bubble_size_cache: Option<DataValues>,

    // ── Appearance ────────────────────────────────────────────────────────────
    /// Fill applied to this series bar/marker/slice, parsed from `<c:spPr>`.
    /// `None` means Excel uses its automatic series color (theme accent cycle).
    pub fill: Option<Fill>,

    // ── Axis association ──────────────────────────────────────────────────────
    /// The numeric ID of the value (or date) axis this series plots against.
    ///
    /// Set during [`crate::parser::chart_parser`] `finish()` by examining the
    /// `<c:axId>` elements inside the series' chart-type element and matching
    /// them against the parsed [`crate::model::axis::Axis`] list.
    ///
    /// `None` when:
    /// * The chart has no axes (e.g. pie charts), or
    /// * The axis ID could not be resolved (malformed XML).
    pub axis_id: Option<u32>,

    /// `true` when this series is plotted against the **secondary** value axis.
    ///
    /// A value axis is considered secondary when its
    /// [`crate::model::axis::AxisPosition`] is `Right` (secondary Y) or
    /// `Top` (secondary X).  Primary axes use `Left` (Y) or `Bottom` (X).
    ///
    /// This field is always `false` when [`axis_id`](Series::axis_id) is
    /// `None`.
    pub is_secondary_axis: bool,
}

impl Series {
    pub fn new(index: u32) -> Self {
        Self {
            index,
            order: index,
            name: None,
            name_ref: None,
            category_ref: None,
            category_values: None,
            category_num_cache: None,
            category_cache_state: CacheState::None,
            value_ref: None,
            value_cache: None,
            value_cache_state: CacheState::None,
            x_value_ref: None,
            x_value_cache: None,
            bubble_size_ref: None,
            bubble_size_cache: None,
            fill: None,
            axis_id: None,
            is_secondary_axis: false,
        }
    }

    /// Returns `true` when values can be read from cache without touching
    /// the worksheet.
    pub fn has_value_cache(&self) -> bool {
        self.value_cache_state == CacheState::Complete
    }

    /// Returns `true` when category labels can be read from cache.
    pub fn has_category_cache(&self) -> bool {
        self.category_cache_state == CacheState::Complete
    }

    /// Returns `true` when this series is plotted on the secondary axis.
    ///
    /// Convenience wrapper around [`is_secondary_axis`](Series::is_secondary_axis).
    pub fn is_on_secondary_axis(&self) -> bool {
        self.is_secondary_axis
    }
}

// ── Unit tests ────────────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn new_series_axis_id_is_none() {
        let s = Series::new(0);
        assert!(s.axis_id.is_none());
    }

    #[test]
    fn new_series_is_secondary_false() {
        let s = Series::new(0);
        assert!(!s.is_secondary_axis);
    }

    #[test]
    fn is_on_secondary_axis_matches_field() {
        let mut s = Series::new(0);
        assert!(!s.is_on_secondary_axis());
        s.is_secondary_axis = true;
        assert!(s.is_on_secondary_axis());
    }

    #[test]
    fn axis_id_roundtrip() {
        let mut s = Series::new(0);
        s.axis_id = Some(42);
        assert_eq!(s.axis_id, Some(42));
    }

    #[test]
    fn secondary_false_when_axis_id_none() {
        let s = Series::new(0);
        // By invariant: is_secondary_axis is false when axis_id is None
        assert!(s.axis_id.is_none());
        assert!(!s.is_secondary_axis);
    }
}
