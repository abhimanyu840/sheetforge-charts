//! [`Chart`] model, [`ChartType`], [`LegendPosition`], [`PlotArea`],
//! [`ChartAnchor`], [`Chart3DView`], and [`Chart3DSurface`].

use serde::{Deserialize, Serialize};

use super::{axis::Axis, series::Series};
use crate::model::color::Fill;

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

// ── PlotArea ─────────────────────────────────────────────────────────────────

/// Everything inside `<c:plotArea>`.
#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct PlotArea {
    pub chart_type: ChartType,
    pub grouping: Option<Grouping>,
    /// `true` for horizontal bars (`<c:barDir val="bar"/>`).
    pub bar_horizontal: bool,
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
}
