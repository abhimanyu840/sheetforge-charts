//! PyO3 bindings for sheetforge-charts — full color/fill/theme exposure.
//! Activated with --features python (via maturin).

use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;

use crate::model::color::{ColorSpec, Fill, Gradient, Rgb};
use crate::model::theme::Theme;

// ── Color helpers ─────────────────────────────────────────────────────────────

fn fill_to_py(fill: &Fill, theme: Option<&Theme>) -> Option<PyFill> {
    match fill {
        Fill::Solid(spec) => Some(PyFill {
            fill_type: "solid".to_owned(),
            color: spec.resolve(theme).map(|rgb| format!("#{}", rgb.to_hex())),
            color_raw: colorspec_raw(spec),
            gradient_stops: vec![],
            gradient_angle: None,
        }),
        Fill::Gradient(grad) => {
            let stops = grad
                .stops
                .iter()
                .map(|s| PyGradientStop {
                    position: s.position as f64 / 100_000.0,
                    color: s
                        .color
                        .resolve(theme)
                        .map(|rgb| format!("#{}", rgb.to_hex())),
                    color_raw: colorspec_raw(&s.color),
                })
                .collect();
            let angle = match &grad.direction {
                Some(crate::model::color::GradientDirection::Linear { angle_deg, .. }) => {
                    Some(*angle_deg)
                }
                _ => None,
            };
            Some(PyFill {
                fill_type: "gradient".to_owned(),
                color: None,
                color_raw: None,
                gradient_stops: stops,
                gradient_angle: angle,
            })
        }
        Fill::None => Some(PyFill {
            fill_type: "none".to_owned(),
            color: None,
            color_raw: None,
            gradient_stops: vec![],
            gradient_angle: None,
        }),
        Fill::Pattern => None,
    }
}

fn colorspec_raw(spec: &ColorSpec) -> Option<String> {
    match spec {
        ColorSpec::Srgb(rgb, _) => Some(format!("srgb:#{}", rgb.to_hex())),
        ColorSpec::Sys(rgb, _) => Some(format!("sys:#{}", rgb.to_hex())),
        ColorSpec::Scheme(slot, _) => Some(format!("theme:{}", slot.as_str())),
        ColorSpec::Preset(name, _) => Some(format!("preset:{}", name)),
    }
}

// ── PyGradientStop ────────────────────────────────────────────────────────────

#[pyclass(name = "GradientStop")]
#[derive(Clone)]
pub struct PyGradientStop {
    #[pyo3(get)]
    pub position: f64,
    #[pyo3(get)]
    pub color: Option<String>,
    #[pyo3(get)]
    pub color_raw: Option<String>,
}

#[pymethods]
impl PyGradientStop {
    fn __repr__(&self) -> String {
        format!(
            "GradientStop(pos={:.0}%, {})",
            self.position * 100.0,
            self.color.as_deref().unwrap_or("?")
        )
    }
}

// ── PyFill ────────────────────────────────────────────────────────────────────

#[pyclass(name = "Fill")]
#[derive(Clone)]
pub struct PyFill {
    #[pyo3(get)]
    pub fill_type: String,
    #[pyo3(get)]
    pub color: Option<String>,
    #[pyo3(get)]
    pub color_raw: Option<String>,
    #[pyo3(get)]
    pub gradient_stops: Vec<PyGradientStop>,
    #[pyo3(get)]
    pub gradient_angle: Option<f64>,
}

#[pymethods]
impl PyFill {
    fn __repr__(&self) -> String {
        match self.fill_type.as_str() {
            "solid" => format!("Fill(solid, {})", self.color.as_deref().unwrap_or("?")),
            "gradient" => format!("Fill(gradient, {} stops)", self.gradient_stops.len()),
            _ => format!("Fill({})", self.fill_type),
        }
    }
}

// ── PyTheme ───────────────────────────────────────────────────────────────────

#[pyclass(name = "Theme")]
pub struct PyTheme {
    inner: Theme,
}

#[pymethods]
impl PyTheme {
    #[getter]
    fn name(&self) -> Option<&str> {
        self.inner.name.as_deref()
    }

    fn colors(&self) -> std::collections::HashMap<String, String> {
        self.inner
            .all_colors()
            .into_iter()
            .map(|(k, rgb)| (k.to_owned(), format!("#{}", rgb.to_hex())))
            .collect()
    }

    fn color(&self, slot: &str) -> Option<String> {
        self.inner
            .color_by_name(slot)
            .map(|rgb| format!("#{}", rgb.to_hex()))
    }

    #[getter]
    fn accent1(&self) -> Option<String> {
        self.inner.accent1().map(|c| format!("#{}", c.to_hex()))
    }
    #[getter]
    fn accent2(&self) -> Option<String> {
        self.inner.accent2().map(|c| format!("#{}", c.to_hex()))
    }
    #[getter]
    fn accent3(&self) -> Option<String> {
        self.inner.accent3().map(|c| format!("#{}", c.to_hex()))
    }
    #[getter]
    fn accent4(&self) -> Option<String> {
        self.inner.accent4().map(|c| format!("#{}", c.to_hex()))
    }
    #[getter]
    fn accent5(&self) -> Option<String> {
        self.inner.accent5().map(|c| format!("#{}", c.to_hex()))
    }
    #[getter]
    fn accent6(&self) -> Option<String> {
        self.inner.accent6().map(|c| format!("#{}", c.to_hex()))
    }
    #[getter]
    fn dk1(&self) -> Option<String> {
        self.inner.dk1().map(|c| format!("#{}", c.to_hex()))
    }
    #[getter]
    fn lt1(&self) -> Option<String> {
        self.inner.lt1().map(|c| format!("#{}", c.to_hex()))
    }
    #[getter]
    fn dk2(&self) -> Option<String> {
        self.inner.dk2().map(|c| format!("#{}", c.to_hex()))
    }
    #[getter]
    fn lt2(&self) -> Option<String> {
        self.inner.lt2().map(|c| format!("#{}", c.to_hex()))
    }

    fn __repr__(&self) -> String {
        format!(
            "Theme(name={:?}, slots={})",
            self.inner.name,
            self.inner.all_colors().len()
        )
    }
}

// ── PyWorkbookCharts ──────────────────────────────────────────────────────────

#[pyclass(name = "WorkbookCharts")]
pub struct PyWorkbookCharts {
    inner: crate::model::workbook::WorkbookCharts,
}

#[pymethods]
impl PyWorkbookCharts {
    #[getter]
    fn source_path(&self) -> &str {
        &self.inner.source_path
    }
    #[getter]
    fn chart_count(&self) -> usize {
        self.inner.chart_count()
    }

    #[getter]
    fn sheets(&self) -> Vec<PySheetCharts> {
        self.inner
            .sheets
            .iter()
            .map(|s| PySheetCharts { inner: s.clone() })
            .collect()
    }

    #[getter]
    fn theme(&self) -> Option<PyTheme> {
        self.inner
            .theme
            .as_ref()
            .map(|t| PyTheme { inner: t.clone() })
    }

    fn __repr__(&self) -> String {
        format!(
            "WorkbookCharts(source={:?}, sheets={}, charts={})",
            self.inner.source_path,
            self.inner.sheets.len(),
            self.inner.chart_count()
        )
    }
}

// ── PySheetCharts ─────────────────────────────────────────────────────────────

#[pyclass(name = "SheetCharts")]
pub struct PySheetCharts {
    inner: crate::model::workbook::SheetCharts,
}

#[pymethods]
impl PySheetCharts {
    #[getter]
    fn name(&self) -> &str {
        &self.inner.name
    }
    #[getter]
    fn index(&self) -> usize {
        self.inner.index
    }

    #[getter]
    fn charts(&self) -> Vec<PyChart> {
        self.inner
            .charts
            .iter()
            .map(|c| PyChart { inner: c.clone() })
            .collect()
    }

    fn __repr__(&self) -> String {
        format!(
            "SheetCharts(name={:?}, charts={})",
            self.inner.name,
            self.inner.charts.len()
        )
    }
}

// ── PyChart ───────────────────────────────────────────────────────────────────

#[pyclass(name = "Chart")]
pub struct PyChart {
    inner: crate::model::chart::Chart,
}

#[pymethods]
impl PyChart {
    #[getter]
    fn chart_path(&self) -> &str {
        &self.inner.chart_path
    }
    #[getter]
    fn chart_type(&self) -> String {
        format!("{:?}", self.inner.chart_type)
    }
    #[getter]
    fn title(&self) -> Option<&str> {
        self.inner.title.as_deref()
    }
    #[getter]
    fn style(&self) -> Option<u32> {
        self.inner.style
    }
    #[getter]
    fn is_pivot_chart(&self) -> bool {
        self.inner.is_pivot_chart
    }
    #[getter]
    fn pivot_table_name(&self) -> Option<&str> {
        self.inner.pivot_table_name.as_deref()
    }

    #[getter]
    fn layers(&self) -> Vec<PyChartLayer> {
        self.inner
            .layers
            .iter()
            .map(|l| PyChartLayer { inner: l.clone() })
            .collect()
    }

    #[getter]
    fn series(&self) -> Vec<PySeries> {
        self.inner
            .series
            .iter()
            .map(|s| PySeries { inner: s.clone() })
            .collect()
    }

    #[getter]
    fn position(&self) -> Option<PyChartPosition> {
        self.inner
            .position
            .as_ref()
            .map(|p| PyChartPosition { inner: p.clone() })
    }

    #[getter]
    fn anchor(&self) -> Option<PyChartAnchor> {
        self.inner
            .anchor
            .as_ref()
            .map(|a| PyChartAnchor { inner: a.clone() })
    }

    fn chart_fill(&self, theme: Option<&PyTheme>) -> Option<PyFill> {
        self.inner
            .chart_fill
            .as_ref()
            .and_then(|f| fill_to_py(f, theme.map(|t| &t.inner)))
    }

    fn plot_area_fill(&self, theme: Option<&PyTheme>) -> Option<PyFill> {
        self.inner
            .plot_area
            .fill
            .as_ref()
            .and_then(|f| fill_to_py(f, theme.map(|t| &t.inner)))
    }

    fn __repr__(&self) -> String {
        format!(
            "Chart(type={:?}, title={:?}, series={}, layers={})",
            self.inner.chart_type,
            self.inner.title,
            self.inner.series.len(),
            self.inner.layers.len()
        )
    }
}

// ── PyChartLayer ──────────────────────────────────────────────────────────────

#[pyclass(name = "ChartLayer")]
pub struct PyChartLayer {
    inner: crate::model::chart::ChartLayer,
}

#[pymethods]
impl PyChartLayer {
    #[getter]
    fn chart_type(&self) -> String {
        format!("{:?}", self.inner.chart_type)
    }
    #[getter]
    fn bar_horizontal(&self) -> bool {
        self.inner.bar_horizontal
    }
    #[getter]
    fn axis_ids(&self) -> Vec<u32> {
        self.inner.axis_ids.clone()
    }

    #[getter]
    fn grouping(&self) -> Option<String> {
        self.inner.grouping.as_ref().map(|g| format!("{:?}", g))
    }

    #[getter]
    fn series(&self) -> Vec<PySeries> {
        self.inner
            .series
            .iter()
            .map(|s| PySeries { inner: s.clone() })
            .collect()
    }

    fn __repr__(&self) -> String {
        format!(
            "ChartLayer(type={:?}, series={}, bar_horizontal={})",
            self.inner.chart_type,
            self.inner.series.len(),
            self.inner.bar_horizontal
        )
    }
}

// ── PySeries ──────────────────────────────────────────────────────────────────

#[pyclass(name = "Series")]
pub struct PySeries {
    inner: crate::model::series::Series,
}

#[pymethods]
impl PySeries {
    #[getter]
    fn index(&self) -> u32 {
        self.inner.index
    }
    #[getter]
    fn order(&self) -> u32 {
        self.inner.order
    }
    #[getter]
    fn name(&self) -> Option<&str> {
        self.inner.name.as_deref()
    }
    #[getter]
    fn is_secondary_axis(&self) -> bool {
        self.inner.is_secondary_axis
    }
    #[getter]
    fn axis_id(&self) -> Option<u32> {
        self.inner.axis_id
    }

    #[getter]
    fn values(&self) -> Vec<f64> {
        self.inner
            .value_cache
            .as_ref()
            .map(|c| {
                c.values
                    .iter()
                    .map(|v| if v.is_nan() { 0.0 } else { *v })
                    .collect()
            })
            .unwrap_or_default()
    }

    #[getter]
    fn categories(&self) -> Vec<String> {
        self.inner
            .category_values
            .as_ref()
            .map(|c| c.values.clone())
            .unwrap_or_default()
    }

    #[getter]
    fn value_ref(&self) -> Option<&str> {
        self.inner.value_ref.as_ref().map(|r| r.formula.as_str())
    }

    #[getter]
    fn category_ref(&self) -> Option<&str> {
        self.inner.category_ref.as_ref().map(|r| r.formula.as_str())
    }

    fn fill(&self, theme: Option<&PyTheme>) -> Option<PyFill> {
        self.inner
            .fill
            .as_ref()
            .and_then(|f| fill_to_py(f, theme.map(|t| &t.inner)))
    }

    fn __repr__(&self) -> String {
        format!(
            "Series(index={}, name={:?}, values={}, secondary={})",
            self.inner.index,
            self.inner.name,
            self.inner
                .value_cache
                .as_ref()
                .map(|c| c.values.len())
                .unwrap_or(0),
            self.inner.is_secondary_axis
        )
    }
}

// ── PyChartPosition ───────────────────────────────────────────────────────────

#[pyclass(name = "ChartPosition")]
pub struct PyChartPosition {
    inner: crate::model::chart::ChartPosition,
}

#[pymethods]
impl PyChartPosition {
    #[getter]
    fn sheet(&self) -> &str {
        &self.inner.sheet
    }
    #[getter]
    fn top_left(&self) -> &str {
        &self.inner.top_left
    }
    #[getter]
    fn bottom_right(&self) -> &str {
        &self.inner.bottom_right
    }
    #[getter]
    fn width_emu(&self) -> Option<i64> {
        self.inner.width_emu
    }
    #[getter]
    fn height_emu(&self) -> Option<i64> {
        self.inner.height_emu
    }

    fn __repr__(&self) -> String {
        format!(
            "ChartPosition(sheet={:?}, {}:{})",
            self.inner.sheet, self.inner.top_left, self.inner.bottom_right
        )
    }
}

// ── PyChartAnchor ─────────────────────────────────────────────────────────────

#[pyclass(name = "ChartAnchor")]
pub struct PyChartAnchor {
    inner: crate::model::chart::ChartAnchor,
}

#[pymethods]
impl PyChartAnchor {
    #[getter]
    fn col_start(&self) -> u32 {
        self.inner.col_start
    }
    #[getter]
    fn row_start(&self) -> u32 {
        self.inner.row_start
    }
    #[getter]
    fn col_end(&self) -> u32 {
        self.inner.col_end
    }
    #[getter]
    fn row_end(&self) -> u32 {
        self.inner.row_end
    }
    #[getter]
    fn col_span(&self) -> u32 {
        self.inner.col_span()
    }
    #[getter]
    fn row_span(&self) -> u32 {
        self.inner.row_span()
    }

    fn __repr__(&self) -> String {
        format!(
            "ChartAnchor(col={}:{}, row={}:{})",
            self.inner.col_start, self.inner.col_end, self.inner.row_start, self.inner.row_end
        )
    }
}

// ── Module entry point ────────────────────────────────────────────────────────

#[pyfunction]
pub fn extract_charts(path: &str) -> PyResult<PyWorkbookCharts> {
    crate::extract_charts(path)
        .map(|wb| PyWorkbookCharts { inner: wb })
        .map_err(|e| PyValueError::new_err(format!("{e:#}")))
}

#[pymodule]
pub fn _core(_py: Python<'_>, m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(extract_charts, m)?)?;
    m.add_class::<PyWorkbookCharts>()?;
    m.add_class::<PySheetCharts>()?;
    m.add_class::<PyChart>()?;
    m.add_class::<PyChartLayer>()?;
    m.add_class::<PySeries>()?;
    m.add_class::<PyChartPosition>()?;
    m.add_class::<PyChartAnchor>()?;
    m.add_class::<PyTheme>()?;
    m.add_class::<PyFill>()?;
    m.add_class::<PyGradientStop>()?;
    Ok(())
}
