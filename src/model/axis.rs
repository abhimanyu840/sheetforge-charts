//! [`Axis`] model — one axis on a chart.

use serde::{Deserialize, Serialize};

// ── AxisType ──────────────────────────────────────────────────────────────────

/// The four DrawingML axis types.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum AxisType {
    /// Category axis (`<c:catAx>`).
    Category,
    /// Value axis (`<c:valAx>`).
    Value,
    /// Date axis (`<c:dateAx>`).
    Date,
    /// Series axis (`<c:serAx>`) — used in 3-D charts.
    Series,
}

impl AxisType {
    /// Map the DrawingML XML local element name to an [`AxisType`].
    pub fn from_xml_tag(tag: &str) -> Option<Self> {
        match tag {
            "catAx"  => Some(Self::Category),
            "valAx"  => Some(Self::Value),
            "dateAx" => Some(Self::Date),
            "serAx"  => Some(Self::Series),
            _        => None,
        }
    }
}

// ── AxisPosition ──────────────────────────────────────────────────────────────

/// Where the axis is drawn relative to the plot area.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum AxisPosition {
    Bottom,
    Top,
    Left,
    Right,
}

// ── Axis ──────────────────────────────────────────────────────────────────────

/// Metadata for a single axis in a chart.
///
/// Maps to a `<c:catAx>`, `<c:valAx>`, `<c:dateAx>`, or `<c:serAx>` element.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Axis {
    /// Numeric axis ID (`<c:axId val="..."/>`).
    pub id: u32,

    /// The axis type — determines which XML element this was parsed from.
    pub axis_type: AxisType,

    /// Optional human-readable title for the axis.
    pub title: Option<String>,

    /// Position of the axis relative to the plot area.
    pub position: Option<AxisPosition>,

    /// Number format string, e.g. `"0.00"`.
    pub number_format: Option<String>,

    /// The ID of the *crossing* axis (the axis this one intersects).
    pub cross_axis_id: Option<u32>,
}

impl Axis {
    /// Minimum constructor.
    pub fn new(id: u32, axis_type: AxisType) -> Self {
        Self {
            id,
            axis_type,
            title: None,
            position: None,
            number_format: None,
            cross_axis_id: None,
        }
    }
}
