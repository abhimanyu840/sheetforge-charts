//! [`Chart`] model, [`ChartType`], [`LegendPosition`], [`PlotArea`],
//! [`ChartAnchor`], [`Chart3DView`], [`Chart3DSurface`], and [`ChartPosition`].

use serde::{Deserialize, Serialize};

use super::{axis::Axis, series::Series};
use crate::model::color::Fill;
use crate::model::pivot::PivotTableMeta;

// ── ChartPosition ─────────────────────────────────────────────────────────────

/// Human-readable placement of a chart on its worksheet.
///
/// Derived from the raw [`ChartAnchor`] after the sheet name is known.
/// Cell addresses use standard **A1 notation** (column letters + 1-based row).
///
/// ## Example
/// ```
/// # use sheetforge_charts::model::chart::ChartPosition;
/// let pos = ChartPosition {
///     sheet:        "Sales".to_owned(),
///     top_left:     "B2".to_owned(),   // col=1, row=1  (0-based → B, 2)
///     bottom_right: "J17".to_owned(),  // col=9, row=16 (0-based → J, 17)
///     width_emu:    None,
///     height_emu:   None,
/// };
/// assert_eq!(pos.top_left, "B2");
/// ```
///
/// For `twoCellAnchor` charts, both corners are exact.
/// For `oneCellAnchor` charts, `bottom_right` is computed from the
/// top-left corner plus the `<xdr:ext cx cy/>` EMU dimensions; it is
/// an approximation because column/row pixel sizes vary.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct ChartPosition {
    /// Display name of the worksheet this chart lives on.
    pub sheet: String,

    /// Top-left cell in A1 notation (e.g. `"B2"`).
    ///
    /// Derived from `ChartAnchor::col_start` and `ChartAnchor::row_start`,
    /// both of which are **0-based** in the raw XML.
    pub top_left: String,

    /// Bottom-right cell in A1 notation (e.g. `"J17"`).
    ///
    /// Derived from `ChartAnchor::col_end` and `ChartAnchor::row_end`.
    /// For `oneCellAnchor` charts this is estimated from the EMU extent.
    pub bottom_right: String,

    /// Chart width in **English Metric Units** (EMU).
    ///
    /// `Some` for `oneCellAnchor` charts (from `<xdr:ext cx="…"/>`).
    /// `None` for `twoCellAnchor` charts (width is implicit from the corner delta).
    pub width_emu: Option<i64>,

    /// Chart height in **English Metric Units** (EMU).
    ///
    /// `Some` for `oneCellAnchor` charts (from `<xdr:ext cy="…"/>`).
    /// `None` for `twoCellAnchor` charts.
    pub height_emu: Option<i64>,
}

impl ChartPosition {
    /// Convert a 0-based column index to an Excel column letter string.
    ///
    /// ```
    /// # use sheetforge_charts::model::chart::ChartPosition;
    /// assert_eq!(ChartPosition::col_to_letter(0),  "A");
    /// assert_eq!(ChartPosition::col_to_letter(25), "Z");
    /// assert_eq!(ChartPosition::col_to_letter(26), "AA");
    /// assert_eq!(ChartPosition::col_to_letter(701), "ZZ");
    /// ```
    pub fn col_to_letter(mut col: u32) -> String {
        let mut result = Vec::new();
        loop {
            result.push((b'A' + (col % 26) as u8) as char);
            if col < 26 {
                break;
            }
            col = col / 26 - 1;
        }
        result.iter().rev().collect()
    }

    /// Build an A1-notation cell address from 0-based column and row indices.
    ///
    /// ```
    /// # use sheetforge_charts::model::chart::ChartPosition;
    /// assert_eq!(ChartPosition::cell_address(0, 0), "A1");
    /// assert_eq!(ChartPosition::cell_address(1, 1), "B2");
    /// assert_eq!(ChartPosition::cell_address(25, 9), "Z10");
    /// ```
    pub fn cell_address(col: u32, row: u32) -> String {
        format!("{}{}", Self::col_to_letter(col), row + 1)
    }

    /// Build a `ChartPosition` from a [`ChartAnchor`] and a sheet name.
    ///
    /// `width_emu` and `height_emu` are `None` for `twoCellAnchor` sources.
    pub fn from_anchor(anchor: &ChartAnchor, sheet: impl Into<String>) -> Self {
        Self {
            sheet: sheet.into(),
            top_left: Self::cell_address(anchor.col_start, anchor.row_start),
            bottom_right: Self::cell_address(anchor.col_end, anchor.row_end),
            width_emu: None,
            height_emu: None,
        }
    }

    /// Build a `ChartPosition` from a `oneCellAnchor` source where only the
    /// top-left corner and EMU dimensions are known.
    ///
    /// `bottom_right` is set equal to `top_left` (the extent in rows/cols
    /// would require knowing column-width / row-height pixel values, which are
    /// sheet-specific and not parsed here).
    pub fn from_one_cell(col: u32, row: u32, cx: i64, cy: i64, sheet: impl Into<String>) -> Self {
        let tl = Self::cell_address(col, row);
        Self {
            sheet: sheet.into(),
            top_left: tl.clone(),
            bottom_right: tl, // extent unknown without column-width data
            width_emu: Some(cx),
            height_emu: Some(cy),
        }
    }
}

// ── ChartAnchor ───────────────────────────────────────────────────────────────

/// Position of a chart on its worksheet, extracted from the
/// `<xdr:twoCellAnchor>` element in `xl/drawings/drawingN.xml`.
///
/// Row and column indices are **0-based** (matching the raw XML values).
/// The pixel/EMU offsets are in English Metric Units (1 inch = 914 400 EMU).
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct ChartAnchor {
    pub col_start: u32,
    pub col_off: i64,
    pub row_start: u32,
    pub row_off: i64,
    pub col_end: u32,
    pub col_end_off: i64,
    pub row_end: u32,
    pub row_end_off: i64,
}

impl ChartAnchor {
    pub fn col_span(&self) -> u32 {
        self.col_end.saturating_sub(self.col_start)
    }
    pub fn row_span(&self) -> u32 {
        self.row_end.saturating_sub(self.row_start)
    }
}

#[cfg(test)]
mod anchor_tests {
    use super::*;
    fn make(cs: u32, rs: u32, ce: u32, re: u32) -> ChartAnchor {
        ChartAnchor {
            col_start: cs,
            col_off: 0,
            row_start: rs,
            row_off: 0,
            col_end: ce,
            col_end_off: 0,
            row_end: re,
            row_end_off: 0,
        }
    }
    #[test]
    fn col_span_basic() {
        assert_eq!(make(0, 0, 8, 15).col_span(), 8);
    }
    #[test]
    fn row_span_basic() {
        assert_eq!(make(0, 0, 8, 15).row_span(), 15);
    }
    #[test]
    fn col_span_zero_equal() {
        assert_eq!(make(3, 0, 3, 5).col_span(), 0);
    }
    #[test]
    fn col_span_saturates() {
        assert_eq!(make(5, 0, 2, 5).col_span(), 0);
    }
}

// ── Chart3DView ───────────────────────────────────────────────────────────────

/// Camera / perspective configuration extracted from `<c:view3D>`.
///
/// ```xml
/// <c:view3D>
///   <c:rotX val="30"/>
///   <c:rotY val="20"/>
///   <c:rAngAx val="1"/>
///   <c:perspective val="30"/>
/// </c:view3D>
/// ```
///
/// All fields are `Option` — Excel omits elements whose value matches the
/// application default.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
pub struct Chart3DView {
    /// X-axis rotation in degrees (`<c:rotX val="…"/>`), range −90 to 90.
    /// Controls the elevation (tilt) angle.
    pub rotation_x: Option<i32>,

    /// Y-axis rotation in degrees (`<c:rotY val="…"/>`), range 0 to 359.
    /// Controls the horizontal spin around the vertical axis.
    pub rotation_y: Option<i32>,

    /// Right-angle axes flag (`<c:rAngAx val="1"/>`).
    /// `true`  → orthographic projection, right-angle axes.
    /// `false` → perspective projection (controlled by `perspective`).
    pub right_angle_axes: Option<bool>,

    /// Perspective depth (`<c:perspective val="…"/>`), range 0–240.
    /// Only meaningful when `right_angle_axes` is `Some(false)`.
    pub perspective: Option<u32>,
}

impl Chart3DView {
    /// `true` when no field carries a value (element was empty or absent).
    pub fn is_empty(&self) -> bool {
        self.rotation_x.is_none()
            && self.rotation_y.is_none()
            && self.right_angle_axes.is_none()
            && self.perspective.is_none()
    }
}

// ── Chart3DSurface ────────────────────────────────────────────────────────────

/// Fill formatting for the three geometry surfaces of a 3-D chart.
///
/// These are extracted from the sibling elements of `<c:plotArea>` that appear
/// only in 3-D chart XML:
///
/// ```xml
/// <c:floor>
///   <c:spPr><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></c:spPr>
/// </c:floor>
/// <c:sideWall>
///   <c:spPr><a:gradFill>…</a:gradFill></c:spPr>
/// </c:sideWall>
/// <c:backWall>
///   <c:spPr><a:solidFill><a:schemeClr val="accent2"/></a:solidFill></c:spPr>
/// </c:backWall>
/// ```
///
/// Each fill is `None` when the corresponding element is absent or carries no
/// `<c:spPr>` child, which is common — Excel omits these elements when the
/// default (automatic) formatting is used.
///
/// Scheme-color fills require a [`crate::model::theme::Theme`] to resolve to a
/// concrete [`Rgb`]; call [`Fill::solid_rgb`] or [`crate::model::color::ColorSpec::resolve`]
/// with the workbook theme.
#[derive(Debug, Clone, PartialEq, Serialize, Deserialize, Default)]
pub struct Chart3DSurface {
    /// Fill of `<c:floor>` — the horizontal base plane.
    pub floor_fill: Option<Fill>,

    /// Fill of `<c:sideWall>` — the left/right vertical wall.
    pub side_wall_fill: Option<Fill>,

    /// Fill of `<c:backWall>` — the rear vertical wall.
    pub back_wall_fill: Option<Fill>,
}

impl Chart3DSurface {
    /// `true` when every fill field is `None`.
    ///
    /// Used by `finish()` to avoid attaching an all-`None` struct to a chart —
    /// callers can rely on `chart.surface.is_none()` meaning "no surface
    /// formatting was present in the XML".
    pub fn is_empty(&self) -> bool {
        self.floor_fill.is_none() && self.side_wall_fill.is_none() && self.back_wall_fill.is_none()
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "snake_case")]
pub enum ChartType {
    Bar,
    HorizontalBar,
    /// `<c:bar3DChart>` with `<c:barDir val="col"/>` (vertical columns, 3-D)
    Bar3D,
    /// `<c:bar3DChart>` with `<c:barDir val="bar"/>` (horizontal bars, 3-D)
    HorizontalBar3D,
    Line,
    Line3D,
    Pie,
    Pie3D,
    Doughnut,
    Area,
    Area3D,
    Scatter,
    Bubble,
    Radar,
    Stock,
    Surface,
    Surface3D,
    Combo,
    #[default]
    Unknown,
}

impl ChartType {
    /// Map a DrawingML XML local element name → [`ChartType`].
    ///
    /// 3-D variants map to their own arms. `bar3DChart` → `Bar3D` initially;
    /// the parser later upgrades to `HorizontalBar3D` when `<c:barDir val="bar"/>` is seen.
    pub fn from_xml_tag(tag: &str) -> Self {
        match tag {
            "barChart" => Self::Bar,
            "bar3DChart" => Self::Bar3D,
            "lineChart" => Self::Line,
            "line3DChart" => Self::Line3D,
            "pieChart" => Self::Pie,
            "pie3DChart" => Self::Pie3D,
            "doughnutChart" => Self::Doughnut,
            "areaChart" => Self::Area,
            "area3DChart" => Self::Area3D,
            "scatterChart" => Self::Scatter,
            "bubbleChart" => Self::Bubble,
            "radarChart" => Self::Radar,
            "stockChart" => Self::Stock,
            "surfaceChart" => Self::Surface,
            "surface3DChart" => Self::Surface3D,
            _ => Self::Unknown,
        }
    }

    /// `true` when `tag` is the name of a plot-area chart element.
    pub fn is_chart_tag(tag: &str) -> bool {
        matches!(
            tag,
            "barChart"
                | "bar3DChart"
                | "lineChart"
                | "line3DChart"
                | "pieChart"
                | "pie3DChart"
                | "doughnutChart"
                | "areaChart"
                | "area3DChart"
                | "scatterChart"
                | "bubbleChart"
                | "radarChart"
                | "stockChart"
                | "surfaceChart"
                | "surface3DChart"
        )
    }

    /// `true` for chart types with a 3-D projection element (`<c:view3D>`).
    pub fn is_3d(&self) -> bool {
        matches!(
            self,
            Self::Bar3D
                | Self::HorizontalBar3D
                | Self::Line3D
                | Self::Pie3D
                | Self::Area3D
                | Self::Surface3D
        )
    }
}

// ── LegendPosition ────────────────────────────────────────────────────────────

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum LegendPosition {
    Bottom,
    Top,
    Left,
    Right,
    TopRight,
}

impl LegendPosition {
    pub fn from_val(val: &str) -> Option<Self> {
        match val {
            "b" => Some(Self::Bottom),
            "t" => Some(Self::Top),
            "l" => Some(Self::Left),
            "r" => Some(Self::Right),
            "tr" => Some(Self::TopRight),
            _ => None,
        }
    }
}

// ── Grouping ──────────────────────────────────────────────────────────────────

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum Grouping {
    Clustered,
    Stacked,
    PercentStacked,
    Standard,
}

impl Grouping {
    pub fn from_val(val: &str) -> Option<Self> {
        match val {
            "clustered" => Some(Self::Clustered),
            "stacked" => Some(Self::Stacked),
            "percentStacked" => Some(Self::PercentStacked),
            "standard" => Some(Self::Standard),
            _ => None,
        }
    }
}

// ── ChartLayer ────────────────────────────────────────────────────────────────

/// One chart-type element inside `<c:plotArea>` (e.g. one `<c:barChart>` or
/// `<c:lineChart>`).
///
/// A standard single-type chart has exactly one layer.  A combo chart has two
/// or more layers, each with its own type and series subset.
///
/// ```xml
/// <c:plotArea>
///   <c:barChart>            ← layer 0: ChartType::Bar
///     <c:ser>…</c:ser>
///     <c:ser>…</c:ser>
///   </c:barChart>
///   <c:lineChart>           ← layer 1: ChartType::Line
///     <c:ser>…</c:ser>
///   </c:lineChart>
/// </c:plotArea>
/// ```
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ChartLayer {
    /// The chart type for this layer.
    pub chart_type: ChartType,
    /// Series belonging to this layer, in document order.
    pub series: Vec<Series>,
    /// Grouping for this layer (`clustered`, `stacked`, etc.).
    /// `None` when the `<c:grouping>` element is absent.
    pub grouping: Option<Grouping>,
    /// `true` when this is a horizontal-bar layer (`<c:barDir val="bar"/>`).
    pub bar_horizontal: bool,
    /// Axis IDs referenced by this layer, from `<c:axId val="N"/>` elements
    /// that are **direct children of the chart-type element** (not inside
    /// `<c:ser>`).  Typically two entries: category-axis ID and value-axis ID.
    ///
    /// Used to determine which axes each series in this layer plots against,
    /// and therefore whether they are on the primary or secondary axis.
    pub axis_ids: Vec<u32>,
}

// ── PlotArea ─────────────────────────────────────────────────────────────────

/// Everything inside `<c:plotArea>`.
#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct PlotArea {
    /// Primary chart type (type of the first layer, or `Combo` when >1 layers).
    pub chart_type: ChartType,
    pub grouping: Option<Grouping>,
    /// `true` for horizontal bars (`<c:barDir val="bar"/>`).
    pub bar_horizontal: bool,
    /// All series across all layers, in document order (flat convenience view).
    pub series: Vec<Series>,
    pub axes: Vec<Axis>,
    pub fill: Option<Fill>,
}

// ── Chart ─────────────────────────────────────────────────────────────────────

/// Fully-parsed chart metadata.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Chart {
    /// ZIP-relative path (`"xl/charts/chart1.xml"`).
    pub chart_path: String,
    /// Primary chart type (mirrors `plot_area.chart_type`).
    pub chart_type: ChartType,
    /// Rich-text title (runs concatenated, formatting stripped).
    pub title: Option<String>,
    pub legend_position: Option<LegendPosition>,
    /// Style index from `<c:style val="N"/>`, 1–48.
    pub style: Option<u32>,
    pub plot_area: PlotArea,
    /// Flat series list — mirrors `plot_area.series` for convenience.
    pub series: Vec<Series>,
    /// Flat axis list — mirrors `plot_area.axes` for convenience.
    pub axes: Vec<Axis>,
    /// Chart-space background fill (`<c:chartSpace><c:spPr>`).
    pub chart_fill: Option<Fill>,
    /// Worksheet position from `<xdr:twoCellAnchor>`. `None` for chartsheets.
    pub anchor: Option<ChartAnchor>,
    /// 3-D camera/perspective configuration from `<c:view3D>`.
    /// `None` for 2-D charts or when the element is absent.
    pub view_3d: Option<Chart3DView>,

    /// Fill formatting for the floor, side-wall, and back-wall surfaces of a
    /// 3-D chart, extracted from `<c:floor>`, `<c:sideWall>`, `<c:backWall>`.
    /// `None` when all three elements are absent or carry no `<c:spPr>`.
    pub surface: Option<Chart3DSurface>,

    /// `true` when `<c:pivotSource>` is present in the chart XML.
    ///
    /// Pivot charts are backed by a PivotTable rather than a direct worksheet
    /// range.  Their series references point into the PivotTable cache instead
    /// of raw cell ranges, so value resolution requires the pivot cache.
    pub is_pivot_chart: bool,

    /// The pivot table name from `<c:pivotSource><c:name>…</c:name>`.
    ///
    /// Excel writes this as `"SheetName!PivotTableName"` (e.g.
    /// `"Sheet1!PivotTable1"`).  `None` when [`is_pivot_chart`](Chart::is_pivot_chart)
    /// is `false`.
    pub pivot_table_name: Option<String>,

    /// Full pivot table metadata resolved by following the relationship chain
    /// `chart → pivotTable → pivotCacheDefinition`.
    ///
    /// `None` when:
    /// * [`is_pivot_chart`](Chart::is_pivot_chart) is `false`, or
    /// * the chart `.rels` has no `pivotTable` relationship, or
    /// * either the `pivotTableDefinition` or `pivotCacheDefinition` XML could
    ///   not be parsed.
    pub pivot_meta: Option<PivotTableMeta>,

    /// Per-layer breakdown of the chart's plot area.
    ///
    /// Each [`ChartLayer`] corresponds to one chart-type element inside
    /// `<c:plotArea>` (e.g. `<c:barChart>`, `<c:lineChart>`).
    ///
    /// * **Single-type chart** — one layer whose `chart_type` matches
    ///   `self.chart_type`.
    /// * **Combo chart** — two or more layers with different types.
    ///   `self.chart_type` is set to [`ChartType::Combo`] in this case.
    ///
    /// `self.series` (flat) and `self.plot_area.series` remain fully populated
    /// as a convenience view; `layers` gives the per-type breakdown.
    pub layers: Vec<ChartLayer>,

    /// Human-readable placement of this chart on its worksheet.
    ///
    /// Populated after Phase 2 relationship resolution when the sheet name and
    /// drawing anchor are both available.  `None` for chart-sheets (charts that
    /// occupy an entire sheet rather than being embedded in a worksheet).
    pub position: Option<ChartPosition>,
}

impl Chart {
    pub fn new_skeleton(chart_path: impl Into<String>) -> Self {
        Self {
            chart_path: chart_path.into(),
            chart_type: ChartType::Unknown,
            title: None,
            legend_position: None,
            style: None,
            plot_area: PlotArea::default(),
            series: Vec::new(),
            axes: Vec::new(),
            chart_fill: None,
            anchor: None,
            view_3d: None,
            surface: None,
            is_pivot_chart: false,
            pivot_table_name: None,
            pivot_meta: None,
            layers: Vec::new(),
            position: None,
        }
    }
}

// ── Unit tests ────────────────────────────────────────────────────────────────

#[cfg(test)]
mod chart_type_tests {
    use super::*;

    // ChartType::from_xml_tag
    #[test]
    fn bar3d_tag() {
        assert_eq!(ChartType::from_xml_tag("bar3DChart"), ChartType::Bar3D);
    }
    #[test]
    fn bar_tag() {
        assert_eq!(ChartType::from_xml_tag("barChart"), ChartType::Bar);
    }
    #[test]
    fn line3d_tag() {
        assert_eq!(ChartType::from_xml_tag("line3DChart"), ChartType::Line3D);
    }
    #[test]
    fn pie3d_tag() {
        assert_eq!(ChartType::from_xml_tag("pie3DChart"), ChartType::Pie3D);
    }
    #[test]
    fn area3d_tag() {
        assert_eq!(ChartType::from_xml_tag("area3DChart"), ChartType::Area3D);
    }
    #[test]
    fn surface3d_tag() {
        assert_eq!(
            ChartType::from_xml_tag("surface3DChart"),
            ChartType::Surface3D
        );
    }
    #[test]
    fn surface_tag() {
        assert_eq!(ChartType::from_xml_tag("surfaceChart"), ChartType::Surface);
    }
    #[test]
    fn unknown_tag() {
        assert_eq!(ChartType::from_xml_tag("fooChart"), ChartType::Unknown);
    }

    // ChartType::is_3d
    #[test]
    fn bar3d_is_3d() {
        assert!(ChartType::Bar3D.is_3d());
    }
    #[test]
    fn hbar3d_is_3d() {
        assert!(ChartType::HorizontalBar3D.is_3d());
    }
    #[test]
    fn line3d_is_3d() {
        assert!(ChartType::Line3D.is_3d());
    }
    #[test]
    fn pie3d_is_3d() {
        assert!(ChartType::Pie3D.is_3d());
    }
    #[test]
    fn area3d_is_3d() {
        assert!(ChartType::Area3D.is_3d());
    }
    #[test]
    fn surface3d_is_3d() {
        assert!(ChartType::Surface3D.is_3d());
    }
    #[test]
    fn bar_not_3d() {
        assert!(!ChartType::Bar.is_3d());
    }
    #[test]
    fn line_not_3d() {
        assert!(!ChartType::Line.is_3d());
    }
    #[test]
    fn pie_not_3d() {
        assert!(!ChartType::Pie.is_3d());
    }
    #[test]
    fn surface_not_3d() {
        assert!(!ChartType::Surface.is_3d());
    }

    // Chart3DView
    #[test]
    fn view3d_empty_default() {
        assert!(Chart3DView::default().is_empty());
    }
    #[test]
    fn view3d_not_empty_rotx() {
        assert!(!Chart3DView {
            rotation_x: Some(30),
            ..Default::default()
        }
        .is_empty());
    }
    #[test]
    fn view3d_not_empty_perspective() {
        assert!(!Chart3DView {
            perspective: Some(15),
            ..Default::default()
        }
        .is_empty());
    }

    // Chart3DSurface
    #[test]
    fn surface_empty_default() {
        assert!(Chart3DSurface::default().is_empty());
    }
    #[test]
    fn surface_not_empty_floor() {
        let s = Chart3DSurface {
            floor_fill: Some(Fill::None),
            ..Default::default()
        };
        assert!(!s.is_empty());
    }
    #[test]
    fn surface_not_empty_side_wall() {
        let s = Chart3DSurface {
            side_wall_fill: Some(Fill::Pattern),
            ..Default::default()
        };
        assert!(!s.is_empty());
    }
    #[test]
    fn surface_not_empty_back_wall() {
        use crate::model::color::{ColorSpec, Rgb};
        let s = Chart3DSurface {
            back_wall_fill: Some(Fill::Solid(ColorSpec::Srgb(Rgb::BLACK, vec![]))),
            ..Default::default()
        };
        assert!(!s.is_empty());
    }

    // Chart pivot fields — defaults
    #[test]
    fn chart_not_pivot_by_default() {
        let c = Chart::new_skeleton("xl/charts/chart1.xml");
        assert!(!c.is_pivot_chart);
    }
    #[test]
    fn chart_pivot_name_none_by_default() {
        let c = Chart::new_skeleton("xl/charts/chart1.xml");
        assert!(c.pivot_table_name.is_none());
    }
    #[test]
    fn chart_pivot_meta_none_by_default() {
        let c = Chart::new_skeleton("xl/charts/chart1.xml");
        assert!(c.pivot_meta.is_none());
    }
    #[test]
    fn chart_layers_empty_by_default() {
        let c = Chart::new_skeleton("xl/charts/chart1.xml");
        assert!(c.layers.is_empty());
    }
    #[test]
    fn chart_position_none_by_default() {
        let c = Chart::new_skeleton("xl/charts/chart1.xml");
        assert!(c.position.is_none());
    }

    // ChartLayer basics
    #[test]
    fn chart_layer_fields() {
        let layer = ChartLayer {
            chart_type: ChartType::Bar,
            series: vec![],
            grouping: Some(Grouping::Clustered),
            bar_horizontal: false,
            axis_ids: vec![],
        };
        assert_eq!(layer.chart_type, ChartType::Bar);
        assert!(layer.series.is_empty());
        assert_eq!(layer.grouping, Some(Grouping::Clustered));
        assert!(!layer.bar_horizontal);
    }
    #[test]
    fn chart_layer_horizontal_bar() {
        let layer = ChartLayer {
            chart_type: ChartType::HorizontalBar,
            series: vec![],
            grouping: None,
            bar_horizontal: true,
            axis_ids: vec![1, 2],
        };
        assert!(layer.bar_horizontal);
    }
}

#[cfg(test)]
mod position_tests {
    use super::*;

    // ── col_to_letter ─────────────────────────────────────────────────────────
    #[test]
    fn col_a() {
        assert_eq!(ChartPosition::col_to_letter(0), "A");
    }
    #[test]
    fn col_b() {
        assert_eq!(ChartPosition::col_to_letter(1), "B");
    }
    #[test]
    fn col_z() {
        assert_eq!(ChartPosition::col_to_letter(25), "Z");
    }
    #[test]
    fn col_aa() {
        assert_eq!(ChartPosition::col_to_letter(26), "AA");
    }
    #[test]
    fn col_ab() {
        assert_eq!(ChartPosition::col_to_letter(27), "AB");
    }
    #[test]
    fn col_az() {
        assert_eq!(ChartPosition::col_to_letter(51), "AZ");
    }
    #[test]
    fn col_ba() {
        assert_eq!(ChartPosition::col_to_letter(52), "BA");
    }
    #[test]
    fn col_zz() {
        assert_eq!(ChartPosition::col_to_letter(701), "ZZ");
    }

    // ── cell_address ──────────────────────────────────────────────────────────
    #[test]
    fn cell_a1() {
        assert_eq!(ChartPosition::cell_address(0, 0), "A1");
    }
    #[test]
    fn cell_b2() {
        assert_eq!(ChartPosition::cell_address(1, 1), "B2");
    }
    #[test]
    fn cell_z10() {
        assert_eq!(ChartPosition::cell_address(25, 9), "Z10");
    }
    #[test]
    fn cell_j17() {
        assert_eq!(ChartPosition::cell_address(9, 16), "J17");
    }
    #[test]
    fn cell_aa1() {
        assert_eq!(ChartPosition::cell_address(26, 0), "AA1");
    }

    // ── from_anchor ───────────────────────────────────────────────────────────
    #[test]
    fn from_anchor_sheet_name() {
        let a = ChartAnchor {
            col_start: 0,
            col_off: 0,
            row_start: 0,
            row_off: 0,
            col_end: 8,
            col_end_off: 0,
            row_end: 15,
            row_end_off: 0,
        };
        let p = ChartPosition::from_anchor(&a, "Sales");
        assert_eq!(p.sheet, "Sales");
    }
    #[test]
    fn from_anchor_top_left() {
        let a = ChartAnchor {
            col_start: 1,
            col_off: 0,
            row_start: 1,
            row_off: 0,
            col_end: 9,
            col_end_off: 0,
            row_end: 16,
            row_end_off: 0,
        };
        let p = ChartPosition::from_anchor(&a, "S");
        assert_eq!(p.top_left, "B2");
    }
    #[test]
    fn from_anchor_bottom_right() {
        let a = ChartAnchor {
            col_start: 1,
            col_off: 0,
            row_start: 1,
            row_off: 0,
            col_end: 9,
            col_end_off: 0,
            row_end: 16,
            row_end_off: 0,
        };
        let p = ChartPosition::from_anchor(&a, "S");
        assert_eq!(p.bottom_right, "J17");
    }
    #[test]
    fn from_anchor_width_height_none() {
        let a = ChartAnchor {
            col_start: 0,
            col_off: 0,
            row_start: 0,
            row_off: 0,
            col_end: 8,
            col_end_off: 0,
            row_end: 15,
            row_end_off: 0,
        };
        let p = ChartPosition::from_anchor(&a, "S");
        assert!(p.width_emu.is_none());
        assert!(p.height_emu.is_none());
    }
    #[test]
    fn from_anchor_a1_origin() {
        let a = ChartAnchor {
            col_start: 0,
            col_off: 0,
            row_start: 0,
            row_off: 0,
            col_end: 5,
            col_end_off: 0,
            row_end: 10,
            row_end_off: 0,
        };
        let p = ChartPosition::from_anchor(&a, "Sheet1");
        assert_eq!(p.top_left, "A1");
        assert_eq!(p.bottom_right, "F11");
    }

    // ── from_one_cell ─────────────────────────────────────────────────────────
    #[test]
    fn from_one_cell_top_left() {
        let p = ChartPosition::from_one_cell(5, 1, 3000000, 2000000, "Sheet2");
        assert_eq!(p.top_left, "F2");
    }
    #[test]
    fn from_one_cell_bottom_right_equals_top_left() {
        let p = ChartPosition::from_one_cell(5, 1, 3000000, 2000000, "Sheet2");
        assert_eq!(p.top_left, p.bottom_right);
    }
    #[test]
    fn from_one_cell_width_height() {
        let p = ChartPosition::from_one_cell(0, 0, 4572000, 2743200, "S");
        assert_eq!(p.width_emu, Some(4572000));
        assert_eq!(p.height_emu, Some(2743200));
    }
    #[test]
    fn from_one_cell_sheet_name() {
        let p = ChartPosition::from_one_cell(0, 0, 1, 1, "MySheet");
        assert_eq!(p.sheet, "MySheet");
    }
}
