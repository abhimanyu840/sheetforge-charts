//! Top-level workbook model returned by [`crate::extract_charts`].

use serde::{Deserialize, Serialize};

use super::chart::Chart;
use crate::model::theme::Theme;

// ── WorkbookCharts ────────────────────────────────────────────────────────────

/// The root value returned by [`crate::extract_charts`].
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct WorkbookCharts {
    /// Absolute path of the `.xlsx` file that was parsed.
    pub source_path: String,

    /// Ordered list of sheets (matches the tab order in Excel).
    pub sheets: Vec<SheetCharts>,

    /// Theme extracted from `xl/theme/theme1.xml`.
    /// `None` if the file contains no theme part (rare but valid).
    pub theme: Option<Theme>,
}

impl WorkbookCharts {
    /// Convenience: iterate every chart across all sheets.
    pub fn all_charts(&self) -> impl Iterator<Item = &Chart> {
        self.sheets.iter().flat_map(|s| s.charts.iter())
    }

    /// Total number of charts across all sheets.
    pub fn chart_count(&self) -> usize {
        self.sheets.iter().map(|s| s.charts.len()).sum()
    }
}

// ── SheetCharts ───────────────────────────────────────────────────────────────

/// One worksheet and all the charts embedded in it.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SheetCharts {
    /// Sheet display name (e.g. `"Sales Q1"`).
    pub name: String,

    /// The `r:id` used to reference this sheet in `workbook.xml`
    /// (e.g. `"rId1"`).
    pub relationship_id: String,

    /// Sheet index (0-based, matches Excel's left-to-right tab order).
    pub index: usize,

    /// Resolved ZIP path of the worksheet part
    /// (e.g. `"xl/worksheets/sheet1.xml"`).
    ///
    /// Populated by Phase 2 after the relationship chain is walked.
    pub part_path: Option<String>,

    /// Resolved ZIP paths of every drawing part attached to this sheet
    /// (e.g. `["xl/drawings/drawing1.xml"]`).
    ///
    /// Populated by Phase 2 after the relationship chain is walked.
    pub drawing_paths: Vec<String>,

    /// Every chart found in this sheet's drawing parts.
    pub charts: Vec<Chart>,
}

impl SheetCharts {
    /// Create a skeleton sheet with no charts (populated in later phases).
    pub fn new(name: impl Into<String>, relationship_id: impl Into<String>, index: usize) -> Self {
        Self {
            name: name.into(),
            relationship_id: relationship_id.into(),
            index,
            part_path: None,
            drawing_paths: Vec::new(),
            charts: Vec::new(),
        }
    }

    /// Record the resolved worksheet part path.
    pub fn set_part_path(&mut self, path: impl Into<String>) {
        self.part_path = Some(path.into());
    }

    /// Add a drawing path discovered via the relationship chain.
    pub fn add_drawing_path(&mut self, path: impl Into<String>) {
        self.drawing_paths.push(path.into());
    }
}
