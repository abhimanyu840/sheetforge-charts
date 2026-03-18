//! Streaming parser for `xl/charts/chartN.xml`.
//!
//! ## Phase 4 improvements over Phase 3
//!
//! | Problem | Fix |
//! |---|---|
//! | numCache always wrote to value_cache regardless of slot | Slot-aware dispatch: xVal->x_value_cache, bubbleSize->bubble_size_cache |
//! | strCache always wrote to category_values even inside tx | Slot guard: strCache only lands in category when ref_slot == Category |
//! | pt index ignored -> sparse data silently reordered | ptCount pre-allocates; pt idx is read as write index |
//! | No way for callers to skip worksheet I/O | CacheState::Complete/Partial/None set on every series |
//! | parse() did read_to_string (UTF-8 scan + alloc) before parsing | Uses read_entry_bytes -> Reader::from_reader(&bytes[..]) |

use anyhow::{Context, Result};
use quick_xml::{events::Event, Reader};

use crate::{
    archive::zip_reader::{read_entry_bytes, XlsxArchive},
    model::{
        axis::{Axis, AxisPosition, AxisType},
        chart::{Chart, Chart3DSurface, ChartType, Grouping, LegendPosition, PlotArea},
        color::{
            ColorMod, ColorSpec, Fill, Gradient, GradientDirection, GradientStop, Rgb,
            ThemeColorSlot,
        },
        series::{CacheState, DataReference, DataValues, Series, StringValues},
    },
};

// ── Public entry points ───────────────────────────────────────────────────────

pub fn parse(archive: &mut XlsxArchive, chart_path: &str) -> Result<Chart> {
    let bytes = read_entry_bytes(archive, chart_path)
        .with_context(|| format!("Cannot read chart part: {chart_path}"))?;
    parse_bytes(&bytes, chart_path)
        .with_context(|| format!("Failed to parse chart XML: {chart_path}"))
}

pub fn parse_bytes(bytes: &[u8], chart_path: &str) -> Result<Chart> {
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().trim_text(true);
    let mut state = ParseState::new(chart_path);
    loop {
        match reader.read_event()? {
            Event::Start(ref e) => {
                let ln = e.local_name();
                let tag = lname(ln.as_ref());
                let dec = reader.decoder();
                state.on_start(tag, e, dec)?;
            }
            Event::Empty(ref e) => {
                let ln = e.local_name();
                let tag = lname(ln.as_ref());
                let dec = reader.decoder();
                state.on_start(tag, e, dec)?;
                state.on_end(tag);
            }
            Event::Text(ref e) => {
                let cow = e.unescape()?;
                state.on_text(cow.as_ref());
            }
            Event::End(ref e) => {
                let ln = e.local_name();
                let tag = lname(ln.as_ref());
                state.on_end(tag);
            }
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(state.finish())
}

#[cfg(test)]
pub(crate) fn parse_xml(xml: &str, chart_path: &str) -> Result<Chart> {
    parse_bytes(xml.as_bytes(), chart_path)
}

fn lname(bytes: &[u8]) -> &str {
    std::str::from_utf8(bytes).unwrap_or("")
}

// ── State machine ─────────────────────────────────────────────────────────────

struct ParseState {
    chart_path: String,
    style: Option<u32>,
    title_text: Option<String>,
    legend_position: Option<LegendPosition>,
    plot_area: PlotArea,
    current_series: Option<Series>,
    current_axis: Option<Axis>,

    in_chart_title: bool,
    in_axis_title: bool,
    in_ser: bool,
    in_ser_tx: bool,
    in_cat: bool,
    in_val: bool,
    in_bubble: bool,
    in_num_cache: bool,
    in_str_cache: bool,
    in_formula: bool,
    in_text_run: bool,
    in_value_elem: bool,
    in_format_code: bool,
    ser_depth: u32,
    in_plot_area: bool,
    chart_tag_depth: u32,

    pending_num_cache: DataValues,
    pending_str_cache: StringValues,
    pending_pt_idx: usize,

    ref_slot: RefSlot,
    title_buf: String,
    formula_buf: String,
    value_buf: String,
    format_buf: String,

    // ── Fill / spPr parsing ───────────────────────────────────────────────────
    /// Depth inside `<c:spPr>` (0 = outside).
    sppr_depth: u32,
    /// Depth inside `<a:spPr>` plot-area background spPr (0 = outside).
    /// We track ser vs plot_area context via `in_ser`.
    fill_ctx: FillContext,
    // Solid fill accumulation
    pending_solid: Option<ColorSpec>,
    // Gradient accumulation
    in_grad_fill: bool,
    pending_grad_stops: Vec<GradientStop>,
    pending_grad_dir: Option<GradientDirection>,
    pending_grad_tile: bool,
    pending_gs_pos: u32,
    // Active color spec being built (inside solidFill or gs)
    pending_color: Option<PendingColor>,
    // Chart-space level fill
    chart_fill: Option<Fill>,

    // ── 3-D view parsing ──────────────────────────────────────────────────────
    /// Set to `true` when the parser is inside `<c:view3D>…</c:view3D>`.
    in_view_3d: bool,
    /// Accumulates fields from `<c:view3D>` child elements.
    pending_view_3d: crate::model::chart::Chart3DView,

    // ── 3-D surface parsing ───────────────────────────────────────────────────
    /// Which surface element we are currently inside (or `None`).
    /// Drives the `FillContext` when `<c:spPr>` opens.
    surface_slot: SurfaceSlot,
    /// Accumulated fill for `<c:floor>`.
    pending_floor_fill: Option<Fill>,
    /// Accumulated fill for `<c:sideWall>`.
    pending_side_wall_fill: Option<Fill>,
    /// Accumulated fill for `<c:backWall>`.
    pending_back_wall_fill: Option<Fill>,

    // ── Pivot chart detection ─────────────────────────────────────────────────
    /// Set to `true` while the parser is inside `<c:pivotSource>…</c:pivotSource>`.
    /// Guards the `<c:name>` text handler so it only fires for pivot names,
    /// not for series-title `<c:name>` elements elsewhere in the document.
    in_pivot_source: bool,
    /// Text accumulator for `<c:pivotSource><c:name>…</c:name>`.
    /// Flushed to `pending_pivot_name` on `</c:name>`.
    pivot_name_buf: String,
    /// Completed pivot table name, set when `</c:name>` closes inside
    /// `<c:pivotSource>`.  Transferred to `Chart.pivot_table_name` in `finish()`.
    pending_pivot_name: Option<String>,
}

#[derive(Debug, Clone, PartialEq)]
enum RefSlot {
    None,
    SeriesName,
    Category,
    Value,
    XValue,
    BubbleSize,
}

/// Which element owns the spPr currently being parsed.
#[derive(Debug, Clone, PartialEq)]
enum FillContext {
    None,
    Series,
    PlotArea,
    Chart,
    /// `<c:floor><c:spPr>…`
    Floor,
    /// `<c:sideWall><c:spPr>…`
    SideWall,
    /// `<c:backWall><c:spPr>…`
    BackWall,
}

/// Which 3-D surface element the parser is currently inside.
/// Used to set the correct [`FillContext`] when `<c:spPr>` opens.
#[derive(Debug, Clone, PartialEq)]
enum SurfaceSlot {
    None,
    Floor,
    SideWall,
    BackWall,
}

/// A color spec that is being accumulated (waiting for its closing tag).
#[derive(Debug, Clone)]
struct PendingColor {
    spec: ColorSpec,
}

impl PendingColor {
    fn srgb(rgb: Rgb) -> Self {
        Self {
            spec: ColorSpec::Srgb(rgb, vec![]),
        }
    }
    fn sys(rgb: Rgb) -> Self {
        Self {
            spec: ColorSpec::Sys(rgb, vec![]),
        }
    }
    fn scheme(slot: ThemeColorSlot) -> Self {
        Self {
            spec: ColorSpec::Scheme(slot, vec![]),
        }
    }
    fn preset(name: &str) -> Self {
        Self {
            spec: ColorSpec::Preset(name.to_owned(), vec![]),
        }
    }

    fn push_mod(&mut self, m: ColorMod) {
        match &mut self.spec {
            ColorSpec::Srgb(_, mods)
            | ColorSpec::Sys(_, mods)
            | ColorSpec::Scheme(_, mods)
            | ColorSpec::Preset(_, mods) => mods.push(m),
        }
    }
}

impl ParseState {
    fn new(chart_path: &str) -> Self {
        Self {
            chart_path: chart_path.to_owned(),
            style: None,
            title_text: None,
            legend_position: None,
            plot_area: PlotArea::default(),
            current_series: None,
            current_axis: None,
            in_chart_title: false,
            in_axis_title: false,
            in_ser: false,
            in_ser_tx: false,
            in_cat: false,
            in_val: false,
            in_bubble: false,
            in_num_cache: false,
            in_str_cache: false,
            in_formula: false,
            in_text_run: false,
            in_value_elem: false,
            in_format_code: false,
            ser_depth: 0,
            in_plot_area: false,
            chart_tag_depth: 0,
            pending_num_cache: DataValues::default(),
            pending_str_cache: StringValues::default(),
            pending_pt_idx: 0,
            ref_slot: RefSlot::None,
            title_buf: String::new(),
            formula_buf: String::new(),
            value_buf: String::new(),
            format_buf: String::new(),
            sppr_depth: 0,
            fill_ctx: FillContext::None,
            pending_solid: None,
            in_grad_fill: false,
            pending_grad_stops: Vec::new(),
            pending_grad_dir: None,
            pending_grad_tile: false,
            pending_gs_pos: 0,
            pending_color: None,
            chart_fill: None,
            in_view_3d: false,
            pending_view_3d: crate::model::chart::Chart3DView::default(),
            surface_slot: SurfaceSlot::None,
            pending_floor_fill: None,
            pending_side_wall_fill: None,
            pending_back_wall_fill: None,
            in_pivot_source: false,
            pivot_name_buf: String::new(),
            pending_pivot_name: None,
        }
    }

    /// Route a completed [`Fill`] to the correct target based on current context.
    fn commit_fill(&mut self, fill: Fill) {
        match self.fill_ctx {
            FillContext::Series => {
                if let Some(s) = self.current_series.as_mut() {
                    s.fill = Some(fill);
                }
            }
            FillContext::PlotArea => {
                self.plot_area.fill = Some(fill);
            }
            FillContext::Chart => {
                self.chart_fill = Some(fill);
            }
            FillContext::Floor => {
                self.pending_floor_fill = Some(fill);
            }
            FillContext::SideWall => {
                self.pending_side_wall_fill = Some(fill);
            }
            FillContext::BackWall => {
                self.pending_back_wall_fill = Some(fill);
            }
            FillContext::None => {}
        }
    }

    fn on_start(
        &mut self,
        tag: &str,
        e: &quick_xml::events::BytesStart<'_>,
        dec: quick_xml::Decoder,
    ) -> Result<()> {
        match tag {
            "style" => {
                if let Some(v) = attr(e, b"val", dec)? {
                    self.style = v.parse().ok();
                }
            }

            // ── Pivot chart detection ─────────────────────────────────────────
            // <c:pivotSource> is a direct child of <c:chartSpace>, sibling of
            // <c:chart>.  Its presence marks this as a pivot-backed chart.
            // We only need the text of its <c:name> child; all other children
            // (<c:fmtId>, etc.) are ignored.
            "pivotSource" => {
                self.in_pivot_source = true;
                self.pivot_name_buf.clear();
            }

            // ── 3-D view ─────────────────────────────────────────────────────
            "view3D" => {
                self.in_view_3d = true;
                self.pending_view_3d = crate::model::chart::Chart3DView::default();
            }
            "rotX" if self.in_view_3d => {
                if let Some(v) = attr(e, b"val", dec)? {
                    self.pending_view_3d.rotation_x = v.parse().ok();
                }
            }
            "rotY" if self.in_view_3d => {
                if let Some(v) = attr(e, b"val", dec)? {
                    self.pending_view_3d.rotation_y = v.parse().ok();
                }
            }
            "rAngAx" if self.in_view_3d => {
                if let Some(v) = attr(e, b"val", dec)? {
                    // "1" = right-angle axes on, "0" = off
                    self.pending_view_3d.right_angle_axes = Some(v.trim() == "1");
                }
            }
            "perspective" if self.in_view_3d => {
                if let Some(v) = attr(e, b"val", dec)? {
                    self.pending_view_3d.perspective = v.parse().ok();
                }
            }

            "plotArea" => {
                self.in_plot_area = true;
            }

            t if ChartType::is_chart_tag(t) && self.in_plot_area => {
                if self.chart_tag_depth == 0 {
                    self.plot_area.chart_type = ChartType::from_xml_tag(t);
                }
                self.chart_tag_depth += 1;
            }
            "barDir" => {
                if let Some(v) = attr(e, b"val", dec)? {
                    self.plot_area.bar_horizontal = v == "bar";
                }
            }
            "grouping" => {
                if let Some(v) = attr(e, b"val", dec)? {
                    self.plot_area.grouping = Grouping::from_val(&v);
                }
            }

            // Series
            "ser" => {
                self.ser_depth += 1;
                if self.ser_depth == 1 {
                    self.in_ser = true;
                    self.current_series = Some(Series::new(0));
                }
            }
            "idx" if self.in_ser => {
                if let Some(v) = attr(e, b"val", dec)? {
                    if let Some(s) = self.current_series.as_mut() {
                        s.index = v.parse().unwrap_or(0);
                    }
                }
            }
            "order" if self.in_ser => {
                if let Some(v) = attr(e, b"val", dec)? {
                    if let Some(s) = self.current_series.as_mut() {
                        s.order = v.parse().unwrap_or(0);
                    }
                }
            }
            "tx" if self.in_ser => {
                self.in_ser_tx = true;
                self.ref_slot = RefSlot::SeriesName;
            }
            "cat" if self.in_ser => {
                self.in_cat = true;
                self.ref_slot = RefSlot::Category;
            }
            "xVal" if self.in_ser => {
                self.in_cat = true;
                self.ref_slot = RefSlot::XValue;
            }
            "val" if self.in_ser => {
                self.in_val = true;
                self.ref_slot = RefSlot::Value;
            }
            "yVal" if self.in_ser => {
                self.in_val = true;
                self.ref_slot = RefSlot::Value;
            }
            "bubbleSize" if self.in_ser => {
                self.in_bubble = true;
                self.ref_slot = RefSlot::BubbleSize;
            }

            "f" => {
                self.in_formula = true;
                self.formula_buf.clear();
            }

            // Cache open
            "numCache" => {
                self.in_num_cache = true;
                self.pending_num_cache = DataValues::default();
            }
            "strCache" => {
                self.in_str_cache = true;
                self.pending_str_cache = StringValues::default();
            }

            // ptCount — pre-allocate
            "ptCount" => {
                if let Some(v) = attr(e, b"val", dec)? {
                    if let Ok(n) = v.parse::<usize>() {
                        if self.in_num_cache {
                            self.pending_num_cache = DataValues::with_capacity(n);
                        } else if self.in_str_cache {
                            self.pending_str_cache = StringValues::with_capacity(n);
                        }
                    }
                }
            }

            // pt — capture idx attribute
            "pt" => {
                self.pending_pt_idx = 0;
                if let Some(v) = attr(e, b"idx", dec)? {
                    self.pending_pt_idx = v.parse().unwrap_or(0);
                }
                self.value_buf.clear();
            }

            "v" => {
                self.in_value_elem = true;
                self.value_buf.clear();
            }
            "formatCode" => {
                self.in_format_code = true;
                self.format_buf.clear();
            }

            // Axes
            t if AxisType::from_xml_tag(t).is_some() => {
                self.current_axis = Some(Axis::new(0, AxisType::from_xml_tag(t).unwrap()));
            }
            "axId" if self.current_axis.is_some() && !self.in_ser => {
                if let Some(v) = attr(e, b"val", dec)? {
                    if let Some(ax) = self.current_axis.as_mut() {
                        ax.id = v.parse().unwrap_or(0);
                    }
                }
            }
            "axPos" if self.current_axis.is_some() => {
                if let Some(v) = attr(e, b"val", dec)? {
                    if let Some(ax) = self.current_axis.as_mut() {
                        ax.position = match v.as_str() {
                            "b" => Some(AxisPosition::Bottom),
                            "t" => Some(AxisPosition::Top),
                            "l" => Some(AxisPosition::Left),
                            "r" => Some(AxisPosition::Right),
                            _ => None,
                        };
                    }
                }
            }
            "crossAx" if self.current_axis.is_some() => {
                if let Some(v) = attr(e, b"val", dec)? {
                    if let Some(ax) = self.current_axis.as_mut() {
                        ax.cross_axis_id = v.parse().ok();
                    }
                }
            }
            "numFmt" if self.current_axis.is_some() => {
                if let Some(v) = attr(e, b"formatCode", dec)? {
                    if let Some(ax) = self.current_axis.as_mut() {
                        ax.number_format = Some(v);
                    }
                }
            }
            "title" if self.current_axis.is_some() => {
                self.in_axis_title = true;
                self.title_buf.clear();
            }
            "title" if !self.in_ser && self.current_axis.is_none() => {
                self.in_chart_title = true;
                self.title_buf.clear();
            }
            "legendPos" => {
                if let Some(v) = attr(e, b"val", dec)? {
                    self.legend_position = LegendPosition::from_val(&v);
                }
            }
            "t" => {
                self.in_text_run = true;
            }

            // ── 3-D surface geometry elements ─────────────────────────────────
            // Each of these may contain a <c:spPr> with fill children.
            // We record which slot we're in so the spPr handler below can pick
            // the right FillContext.  The elements themselves carry no attributes
            // we need — all meaningful data is in their <c:spPr> child.
            "floor" => {
                self.surface_slot = SurfaceSlot::Floor;
            }
            "sideWall" => {
                self.surface_slot = SurfaceSlot::SideWall;
            }
            "backWall" => {
                self.surface_slot = SurfaceSlot::BackWall;
            }

            // ── spPr (shape properties) — series fill or plot area background ─
            // We track depth so nested spPr elements don't confuse context.
            "spPr" => {
                self.sppr_depth += 1;
                if self.sppr_depth == 1 {
                    // Copy fields to locals first to satisfy the borrow checker:
                    // we need an immutable borrow for `surface_slot`/`in_ser`/
                    // `in_plot_area` and a mutable one to write `fill_ctx`.
                    let slot = self.surface_slot.clone();
                    let in_ser = self.in_ser;
                    let in_pa = self.in_plot_area;
                    self.fill_ctx = match slot {
                        SurfaceSlot::Floor => FillContext::Floor,
                        SurfaceSlot::SideWall => FillContext::SideWall,
                        SurfaceSlot::BackWall => FillContext::BackWall,
                        SurfaceSlot::None if in_ser => FillContext::Series,
                        SurfaceSlot::None if in_pa => FillContext::PlotArea,
                        SurfaceSlot::None => FillContext::Chart,
                    };
                }
            }

            // ── solidFill ────────────────────────────────────────────────────
            "solidFill" if self.sppr_depth > 0 => {
                self.pending_solid = None;
                self.pending_color = None;
            }

            // ── gradFill ─────────────────────────────────────────────────────
            "gradFill" if self.sppr_depth > 0 => {
                self.in_grad_fill = true;
                self.pending_grad_stops.clear();
                self.pending_grad_dir = None;
                self.pending_grad_tile = false;
            }

            // gradient stop
            "gs" if self.in_grad_fill => {
                self.pending_gs_pos = 0;
                if let Some(v) = attr(e, b"pos", dec)? {
                    self.pending_gs_pos = v.parse().unwrap_or(0);
                }
                self.pending_color = None;
            }

            // linear gradient direction
            "lin" if self.in_grad_fill => {
                if let Some(v) = attr(e, b"ang", dec)? {
                    let angle_deg = v.parse::<f64>().unwrap_or(0.0) / 60_000.0;
                    let scaled = attr(e, b"scaled", dec)?.as_deref() == Some("1");
                    self.pending_grad_dir = Some(GradientDirection::Linear { angle_deg, scaled });
                }
            }

            // path gradient direction
            "path" if self.in_grad_fill => {
                if let Some(v) = attr(e, b"path", dec)? {
                    self.pending_grad_dir = Some(GradientDirection::Path(v));
                }
            }

            "tileRect" if self.in_grad_fill => {
                self.pending_grad_tile = true;
            }

            // ── Color elements (inside solidFill or gs) ───────────────────────
            "srgbClr" if self.sppr_depth > 0 => {
                if let Some(v) = attr(e, b"val", dec)? {
                    if let Some(rgb) = Rgb::from_hex(&v) {
                        self.pending_color = Some(PendingColor::srgb(rgb));
                    }
                }
            }
            "sysClr" if self.sppr_depth > 0 => {
                if let Some(v) = attr(e, b"lastClr", dec)? {
                    if let Some(rgb) = Rgb::from_hex(&v) {
                        self.pending_color = Some(PendingColor::sys(rgb));
                    }
                }
            }
            "schemeClr" if self.sppr_depth > 0 => {
                if let Some(v) = attr(e, b"val", dec)? {
                    if let Some(slot) = ThemeColorSlot::from_str(&v) {
                        self.pending_color = Some(PendingColor::scheme(slot));
                    }
                }
            }
            "prstClr" if self.sppr_depth > 0 => {
                if let Some(v) = attr(e, b"val", dec)? {
                    self.pending_color = Some(PendingColor::preset(&v));
                }
            }

            // ── Color modifiers ───────────────────────────────────────────────
            t if self.sppr_depth > 0 && self.pending_color.is_some() => {
                if let Some(v) = attr(e, b"val", dec)? {
                    if let Ok(n) = v.parse::<i32>() {
                        if let Some(m) = ColorMod::from_tag_val(t, n) {
                            if let Some(pc) = self.pending_color.as_mut() {
                                pc.push_mod(m);
                            }
                        }
                    }
                }
            }

            _ => {}
        }
        Ok(())
    }

    fn on_text(&mut self, text: &str) {
        if text.is_empty() {
            return;
        }
        if self.in_formula {
            self.formula_buf.push_str(text);
        } else if self.in_value_elem {
            self.value_buf.push_str(text);
        } else if self.in_format_code {
            self.format_buf.push_str(text);
        }
        // Pivot source name — <c:pivotSource><c:name>text</c:name>.
        // Must be checked before in_text_run because <c:name> has no <a:t> wrapper.
        else if self.in_pivot_source {
            self.pivot_name_buf.push_str(text);
        } else if self.in_text_run {
            if self.in_chart_title || self.in_axis_title {
                if !self.title_buf.is_empty() {
                    self.title_buf.push(' ');
                }
                self.title_buf.push_str(text);
            } else if self.in_ser_tx {
                if let Some(s) = self.current_series.as_mut() {
                    s.name = Some(text.to_owned());
                }
            }
        }
    }

    fn on_end(&mut self, tag: &str) {
        match tag {
            // ── 3-D view ──────────────────────────────────────────────────────
            "view3D" => {
                self.in_view_3d = false;
                // view_3d is collected by finish(); nothing to flush here.
            }

            // ── Pivot chart detection ──────────────────────────────────────────
            // </c:name> fires for many elements (series tx, axis titles, etc.).
            // The `in_pivot_source` guard ensures we only capture the name that
            // is a direct child of <c:pivotSource>.
            "name" if self.in_pivot_source => {
                let name = std::mem::take(&mut self.pivot_name_buf);
                if !name.is_empty() {
                    self.pending_pivot_name = Some(name);
                }
            }
            "pivotSource" => {
                self.in_pivot_source = false;
            }

            // ── 3-D surface elements ──────────────────────────────────────────
            // The fills were already committed to pending_*_fill by commit_fill()
            // when their <c:spPr> closed.  We only need to clear surface_slot.
            "floor" => {
                self.surface_slot = SurfaceSlot::None;
            }
            "sideWall" => {
                self.surface_slot = SurfaceSlot::None;
            }
            "backWall" => {
                self.surface_slot = SurfaceSlot::None;
            }

            "plotArea" => {
                self.in_plot_area = false;
            }

            t if ChartType::is_chart_tag(t) && self.chart_tag_depth > 0 => {
                self.chart_tag_depth -= 1;
            }

            "ser" => {
                if self.ser_depth > 0 {
                    self.ser_depth -= 1;
                }
                if self.ser_depth == 0 {
                    self.in_ser = false;
                    if let Some(s) = self.current_series.take() {
                        self.plot_area.series.push(s);
                    }
                }
            }

            "tx" => {
                self.in_ser_tx = false;
            }
            "cat" | "xVal" => {
                self.in_cat = false;
                self.ref_slot = RefSlot::None;
            }
            "val" | "yVal" => {
                self.in_val = false;
                self.ref_slot = RefSlot::None;
            }
            "bubbleSize" => {
                self.in_bubble = false;
                self.ref_slot = RefSlot::None;
            }

            "f" => {
                self.in_formula = false;
                let formula = std::mem::take(&mut self.formula_buf);
                if formula.is_empty() {
                    return;
                }
                let dr = DataReference { formula };
                match &self.ref_slot {
                    RefSlot::SeriesName => {
                        if let Some(s) = self.current_series.as_mut() {
                            s.name_ref = Some(dr);
                        }
                    }
                    RefSlot::Category => {
                        if let Some(s) = self.current_series.as_mut() {
                            s.category_ref = Some(dr);
                        }
                    }
                    RefSlot::Value => {
                        if let Some(s) = self.current_series.as_mut() {
                            s.value_ref = Some(dr);
                        }
                    }
                    RefSlot::XValue => {
                        if let Some(s) = self.current_series.as_mut() {
                            s.x_value_ref = Some(dr);
                        }
                    }
                    RefSlot::BubbleSize => {
                        if let Some(s) = self.current_series.as_mut() {
                            s.bubble_size_ref = Some(dr);
                        }
                    }
                    RefSlot::None => {}
                }
            }

            // <c:v> closed — write value into pending cache at pending_pt_idx
            "v" => {
                self.in_value_elem = false;
                let raw = std::mem::take(&mut self.value_buf);
                if raw.is_empty() {
                    return;
                }
                let idx = self.pending_pt_idx;
                if self.in_num_cache {
                    if let Ok(f) = raw.parse::<f64>() {
                        self.pending_num_cache.set(idx, f);
                    }
                } else if self.in_str_cache {
                    self.pending_str_cache.set(idx, raw);
                }
            }

            "pt" => { /* idx already consumed in on_start */ }

            "formatCode" => {
                self.in_format_code = false;
                let fc = std::mem::take(&mut self.format_buf);
                if self.in_num_cache && !fc.is_empty() {
                    self.pending_num_cache.format_code = Some(fc);
                }
            }

            // numCache closed — route to correct Series field by slot
            "numCache" => {
                self.in_num_cache = false;
                let cache = std::mem::take(&mut self.pending_num_cache);
                if let Some(s) = self.current_series.as_mut() {
                    let state = num_cache_state(&cache);
                    match &self.ref_slot {
                        RefSlot::XValue => {
                            s.x_value_cache = Some(cache);
                        }
                        RefSlot::BubbleSize => {
                            s.bubble_size_cache = Some(cache);
                        }
                        _ => {
                            s.value_cache_state = state;
                            s.value_cache = Some(cache);
                        }
                    }
                }
            }

            // strCache closed — route by slot
            // SeriesName -> series.name (first entry)
            // Category / XValue -> category_values
            "strCache" => {
                self.in_str_cache = false;
                let cache = std::mem::take(&mut self.pending_str_cache);
                if let Some(s) = self.current_series.as_mut() {
                    match &self.ref_slot {
                        RefSlot::SeriesName => {
                            if let Some(name) = cache.values.into_iter().find(|v| !v.is_empty()) {
                                s.name = Some(name);
                            }
                        }
                        _ => {
                            let state = str_cache_state(&cache);
                            s.category_cache_state = state;
                            s.category_values = Some(cache);
                        }
                    }
                }
            }

            "catAx" | "valAx" | "dateAx" | "serAx" => {
                if let Some(ax) = self.current_axis.take() {
                    self.plot_area.axes.push(ax);
                }
            }
            "title" if self.in_axis_title => {
                self.in_axis_title = false;
                let t = std::mem::take(&mut self.title_buf);
                if let Some(ax) = self.current_axis.as_mut() {
                    if !t.is_empty() {
                        ax.title = Some(t);
                    }
                }
            }
            "title" if self.in_chart_title => {
                self.in_chart_title = false;
                let t = std::mem::take(&mut self.title_buf);
                if !t.is_empty() {
                    self.title_text = Some(t);
                }
            }
            "t" => {
                self.in_text_run = false;
            }

            // ── Color element closed — store spec into pending_solid or grad stop ─
            "srgbClr" | "sysClr" | "schemeClr" | "prstClr" => {
                // Color closed: if we're inside a gs, it stays in pending_color
                // until </gs>. If inside solidFill, promote to pending_solid.
                // (We distinguish by in_grad_fill flag.)
                if !self.in_grad_fill {
                    // solidFill context — color is fully built
                    if let Some(pc) = self.pending_color.take() {
                        self.pending_solid = Some(pc.spec);
                    }
                }
                // In grad context, pending_color stays until </gs>
            }

            // gradient stop closed
            "gs" if self.in_grad_fill => {
                if let Some(pc) = self.pending_color.take() {
                    self.pending_grad_stops.push(GradientStop {
                        position: self.pending_gs_pos,
                        color: pc.spec,
                    });
                }
            }

            // solidFill closed — build Fill::Solid and route to context
            "solidFill" if self.sppr_depth > 0 => {
                if let Some(spec) = self.pending_solid.take() {
                    self.commit_fill(Fill::Solid(spec));
                }
            }

            // gradFill closed — build Fill::Gradient and route
            "gradFill" if self.in_grad_fill => {
                self.in_grad_fill = false;
                let grad = Gradient {
                    stops: std::mem::take(&mut self.pending_grad_stops),
                    direction: self.pending_grad_dir.take(),
                    tile: self.pending_grad_tile,
                };
                self.commit_fill(Fill::Gradient(grad));
            }

            // noFill — explicit transparent
            "noFill" if self.sppr_depth > 0 => {
                self.commit_fill(Fill::None);
            }

            // spPr closed
            "spPr" => {
                if self.sppr_depth > 0 {
                    self.sppr_depth -= 1;
                }
                if self.sppr_depth == 0 {
                    self.fill_ctx = FillContext::None;
                }
            }

            _ => {}
        }
    }

    fn finish(self) -> Chart {
        let series = self.plot_area.series.clone();
        let axes = self.plot_area.axes.clone();
        let chart_type = self.plot_area.chart_type.clone();

        // Resolve horizontal-bar variants (both 2-D and 3-D).
        let chart_type = match (&chart_type, self.plot_area.bar_horizontal) {
            (ChartType::Bar, true) => ChartType::HorizontalBar,
            (ChartType::Bar3D, true) => ChartType::HorizontalBar3D,
            _ => chart_type,
        };

        // Only attach view_3d when it actually carried data.
        let view_3d = if self.pending_view_3d.is_empty() {
            None
        } else {
            Some(self.pending_view_3d)
        };

        // Only attach surface when at least one fill was parsed.
        let surface = {
            let s = Chart3DSurface {
                floor_fill: self.pending_floor_fill,
                side_wall_fill: self.pending_side_wall_fill,
                back_wall_fill: self.pending_back_wall_fill,
            };
            if s.is_empty() {
                None
            } else {
                Some(s)
            }
        };

        Chart {
            chart_path: self.chart_path,
            chart_type: chart_type.clone(),
            title: self.title_text,
            legend_position: self.legend_position,
            style: self.style,
            plot_area: PlotArea {
                chart_type,
                ..self.plot_area
            },
            series,
            axes,
            chart_fill: self.chart_fill,
            anchor: None, // set by sheet_parser after drawing anchor resolved
            view_3d,
            surface,
            is_pivot_chart: self.pending_pivot_name.is_some(),
            pivot_table_name: self.pending_pivot_name,
        }
    }
}

// ── Cache completeness ────────────────────────────────────────────────────────

fn num_cache_state(c: &DataValues) -> CacheState {
    if c.values.is_empty() {
        CacheState::None
    } else if c.is_complete() {
        CacheState::Complete
    } else {
        CacheState::Partial
    }
}

fn str_cache_state(c: &StringValues) -> CacheState {
    if c.values.is_empty() {
        CacheState::None
    } else if c.is_complete() {
        CacheState::Complete
    } else {
        CacheState::Partial
    }
}

// ── Attribute helper ──────────────────────────────────────────────────────────

fn attr(
    e: &quick_xml::events::BytesStart<'_>,
    name: &[u8],
    dec: quick_xml::Decoder,
) -> Result<Option<String>> {
    for a in e.attributes() {
        let a = a.context("Malformed XML attribute")?;
        if a.key.local_name().as_ref() == name {
            return Ok(Some(a.decode_and_unescape_value(dec)?.into_owned()));
        }
    }
    Ok(None)
}

// ── Tests ─────────────────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::{
        axis::{AxisPosition, AxisType},
        chart::{ChartType, Grouping, LegendPosition},
        color::{ColorMod, ColorSpec, Fill, GradientDirection, Rgb, ThemeColorSlot},
        series::CacheState,
    };

    const BAR_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:style val="2"/>
  <c:chart>
    <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/>
      <a:p><a:r><a:t>Sales Overview</a:t></a:r></a:p>
    </c:rich></c:tx><c:overlay val="0"/></c:title>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:tx>
            <c:strRef><c:f>Sales!$B$1</c:f>
              <c:strCache>
                <c:ptCount val="1"/>
                <c:pt idx="0"><c:v>Revenue</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:tx>
          <c:cat>
            <c:strRef><c:f>Sales!$A$2:$A$4</c:f>
              <c:strCache>
                <c:ptCount val="3"/>
                <c:pt idx="0"><c:v>Jan</c:v></c:pt>
                <c:pt idx="2"><c:v>Mar</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef><c:f>Sales!$B$2:$B$4</c:f>
              <c:numCache>
                <c:formatCode>0.00</c:formatCode>
                <c:ptCount val="3"/>
                <c:pt idx="0"><c:v>1000</c:v></c:pt>
                <c:pt idx="1"><c:v>1500</c:v></c:pt>
                <c:pt idx="2"><c:v>1200</c:v></c:pt>
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
        <c:ser>
          <c:idx val="1"/><c:order val="1"/>
          <c:tx><c:strRef><c:f>Sales!$C$1</c:f></c:strRef></c:tx>
          <c:cat><c:strRef><c:f>Sales!$A$2:$A$4</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sales!$C$2:$C$4</c:f></c:numRef></c:val>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
      <c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
      <c:valAx>
        <c:axId val="2"/><c:axPos val="l"/>
        <c:numFmt formatCode="General" sourceLinked="1"/>
        <c:crossAx val="1"/>
      </c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="b"/></c:legend>
  </c:chart>
</c:chartSpace>"#;

    const LINE_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/>
      <a:p><a:r><a:t>Monthly Expenses</a:t></a:r></a:p>
    </c:rich></c:tx></c:title>
    <c:plotArea>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:tx><c:strRef><c:f>Expenses!$B$1</c:f></c:strRef></c:tx>
          <c:cat><c:strRef><c:f>Expenses!$A$2:$A$3</c:f></c:strRef></c:cat>
          <c:val>
            <c:numRef><c:f>Expenses!$B$2:$B$3</c:f>
              <c:numCache>
                <c:formatCode>General</c:formatCode>
                <c:ptCount val="2"/>
                <c:pt idx="0"><c:v>800</c:v></c:pt>
                <c:pt idx="1"><c:v>950</c:v></c:pt>
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:lineChart>
      <c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="r"/></c:legend>
  </c:chart>
</c:chartSpace>"#;

    const PIE_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:pieChart>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:tx><c:strRef><c:f>Data!$B$1</c:f></c:strRef></c:tx>
          <c:cat><c:strRef><c:f>Data!$A$2:$A$5</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Data!$B$2:$B$5</c:f></c:numRef></c:val>
        </c:ser>
      </c:pieChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

    const SCATTER_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:scatterChart>
        <c:scatterStyle val="lineMarker"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:tx><c:strRef><c:f>Sheet1!$A$1</c:f></c:strRef></c:tx>
          <c:xVal>
            <c:numRef><c:f>Sheet1!$A$2:$A$4</c:f>
              <c:numCache>
                <c:ptCount val="3"/>
                <c:pt idx="0"><c:v>1</c:v></c:pt>
                <c:pt idx="1"><c:v>2</c:v></c:pt>
                <c:pt idx="2"><c:v>3</c:v></c:pt>
              </c:numCache>
            </c:numRef>
          </c:xVal>
          <c:yVal>
            <c:numRef><c:f>Sheet1!$B$2:$B$4</c:f>
              <c:numCache>
                <c:ptCount val="3"/>
                <c:pt idx="0"><c:v>10</c:v></c:pt>
                <c:pt idx="1"><c:v>20</c:v></c:pt>
                <c:pt idx="2"><c:v>30</c:v></c:pt>
              </c:numCache>
            </c:numRef>
          </c:yVal>
        </c:ser>
      </c:scatterChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

    // chart type
    #[test]
    fn bar_type() {
        assert_eq!(
            parse_xml(BAR_XML, "c.xml").unwrap().chart_type,
            ChartType::Bar
        );
    }
    #[test]
    fn line_type() {
        assert_eq!(
            parse_xml(LINE_XML, "c.xml").unwrap().chart_type,
            ChartType::Line
        );
    }
    #[test]
    fn pie_type() {
        assert_eq!(
            parse_xml(PIE_XML, "c.xml").unwrap().chart_type,
            ChartType::Pie
        );
    }
    #[test]
    fn scatter_type() {
        assert_eq!(
            parse_xml(SCATTER_XML, "c.xml").unwrap().chart_type,
            ChartType::Scatter
        );
    }

    // title
    #[test]
    fn bar_title() {
        assert_eq!(
            parse_xml(BAR_XML, "c.xml").unwrap().title.as_deref(),
            Some("Sales Overview")
        );
    }
    #[test]
    fn line_title() {
        assert_eq!(
            parse_xml(LINE_XML, "c.xml").unwrap().title.as_deref(),
            Some("Monthly Expenses")
        );
    }
    #[test]
    fn pie_no_title() {
        assert!(parse_xml(PIE_XML, "c.xml").unwrap().title.is_none());
    }

    // style / legend / grouping
    #[test]
    fn style_present() {
        assert_eq!(parse_xml(BAR_XML, "c.xml").unwrap().style, Some(2));
    }
    #[test]
    fn style_absent() {
        assert!(parse_xml(LINE_XML, "c.xml").unwrap().style.is_none());
    }
    #[test]
    fn legend_bottom() {
        assert_eq!(
            parse_xml(BAR_XML, "c.xml").unwrap().legend_position,
            Some(LegendPosition::Bottom)
        );
    }
    #[test]
    fn legend_right() {
        assert_eq!(
            parse_xml(LINE_XML, "c.xml").unwrap().legend_position,
            Some(LegendPosition::Right)
        );
    }
    #[test]
    fn legend_none() {
        assert!(parse_xml(PIE_XML, "c.xml")
            .unwrap()
            .legend_position
            .is_none());
    }
    #[test]
    fn grouping_clustered() {
        assert_eq!(
            parse_xml(BAR_XML, "c.xml").unwrap().plot_area.grouping,
            Some(Grouping::Clustered)
        );
    }
    #[test]
    fn grouping_standard() {
        assert_eq!(
            parse_xml(LINE_XML, "c.xml").unwrap().plot_area.grouping,
            Some(Grouping::Standard)
        );
    }

    // series name from strCache inside <c:tx>
    #[test]
    fn series_name_from_strcache_in_tx() {
        let c = parse_xml(BAR_XML, "c.xml").unwrap();
        assert_eq!(c.series[0].name.as_deref(), Some("Revenue"));
    }
    #[test]
    fn series_name_ref_still_set() {
        let c = parse_xml(BAR_XML, "c.xml").unwrap();
        assert_eq!(
            c.series[0].name_ref.as_ref().map(|r| r.formula.as_str()),
            Some("Sales!$B$1")
        );
    }
    #[test]
    fn strcache_in_tx_does_not_pollute_category_values() {
        // category_values comes from <c:cat> strCache (3 entries), not <c:tx> (1 entry)
        let c = parse_xml(BAR_XML, "c.xml").unwrap();
        assert_eq!(
            c.series[0].category_values.as_ref().unwrap().values.len(),
            3
        );
    }

    // sparse pt idx
    #[test]
    fn sparse_cat_gap_is_empty_string() {
        let c = parse_xml(BAR_XML, "c.xml").unwrap();
        let cats = c.series[0].category_values.as_ref().unwrap();
        assert_eq!(cats.values.len(), 3);
        assert_eq!(cats.values[0], "Jan");
        assert_eq!(cats.values[1], ""); // gap
        assert_eq!(cats.values[2], "Mar");
    }
    #[test]
    fn ptcount_recorded_on_str_cache() {
        let c = parse_xml(BAR_XML, "c.xml").unwrap();
        assert_eq!(
            c.series[0].category_values.as_ref().unwrap().pt_count,
            Some(3)
        );
    }
    #[test]
    fn ptcount_recorded_on_num_cache() {
        let c = parse_xml(BAR_XML, "c.xml").unwrap();
        assert_eq!(c.series[0].value_cache.as_ref().unwrap().pt_count, Some(3));
    }

    // value cache correctness
    #[test]
    fn value_cache_bar() {
        assert_eq!(
            parse_xml(BAR_XML, "c.xml").unwrap().series[0]
                .value_cache
                .as_ref()
                .unwrap()
                .values,
            vec![1000.0, 1500.0, 1200.0]
        );
    }
    #[test]
    fn value_cache_line() {
        assert_eq!(
            parse_xml(LINE_XML, "c.xml").unwrap().series[0]
                .value_cache
                .as_ref()
                .unwrap()
                .values,
            vec![800.0, 950.0]
        );
    }
    #[test]
    fn value_cache_fmt() {
        assert_eq!(
            parse_xml(BAR_XML, "c.xml").unwrap().series[0]
                .value_cache
                .as_ref()
                .unwrap()
                .format_code
                .as_deref(),
            Some("0.00")
        );
    }

    // CacheState
    #[test]
    fn complete_cache_state() {
        assert_eq!(
            parse_xml(BAR_XML, "c.xml").unwrap().series[0].value_cache_state,
            CacheState::Complete
        );
    }
    #[test]
    fn has_value_cache_true() {
        assert!(parse_xml(BAR_XML, "c.xml").unwrap().series[0].has_value_cache());
    }
    #[test]
    fn no_cache_state_none() {
        assert_eq!(
            parse_xml(BAR_XML, "c.xml").unwrap().series[1].value_cache_state,
            CacheState::None
        );
    }
    #[test]
    fn no_cache_has_false() {
        assert!(!parse_xml(BAR_XML, "c.xml").unwrap().series[1].has_value_cache());
    }
    #[test]
    fn sparse_cat_is_partial() {
        assert_eq!(
            parse_xml(BAR_XML, "c.xml").unwrap().series[0].category_cache_state,
            CacheState::Partial
        );
    }
    #[test]
    fn partial_no_has_cache() {
        assert!(!parse_xml(BAR_XML, "c.xml").unwrap().series[0].has_category_cache());
    }

    // scatter: xVal / yVal routing
    #[test]
    fn scatter_x_cache() {
        assert_eq!(
            parse_xml(SCATTER_XML, "c.xml").unwrap().series[0]
                .x_value_cache
                .as_ref()
                .unwrap()
                .values,
            vec![1.0, 2.0, 3.0]
        );
    }
    #[test]
    fn scatter_y_cache() {
        assert_eq!(
            parse_xml(SCATTER_XML, "c.xml").unwrap().series[0]
                .value_cache
                .as_ref()
                .unwrap()
                .values,
            vec![10.0, 20.0, 30.0]
        );
    }
    #[test]
    fn scatter_x_ref() {
        assert_eq!(
            parse_xml(SCATTER_XML, "c.xml").unwrap().series[0]
                .x_value_ref
                .as_ref()
                .map(|r| r.formula.as_str()),
            Some("Sheet1!$A$2:$A$4")
        );
    }
    #[test]
    fn scatter_y_ref() {
        assert_eq!(
            parse_xml(SCATTER_XML, "c.xml").unwrap().series[0]
                .value_ref
                .as_ref()
                .map(|r| r.formula.as_str()),
            Some("Sheet1!$B$2:$B$4")
        );
    }
    #[test]
    fn scatter_x_not_in_value_cache() {
        // xVal should land in x_value_cache, NOT value_cache
        let c = parse_xml(SCATTER_XML, "c.xml").unwrap();
        let xc = c.series[0].x_value_cache.as_ref().unwrap();
        let vc = c.series[0].value_cache.as_ref().unwrap();
        assert_ne!(xc.values, vc.values);
    }

    // axes
    #[test]
    fn bar_two_axes() {
        assert_eq!(parse_xml(BAR_XML, "c.xml").unwrap().axes.len(), 2);
    }
    #[test]
    fn cat_axis_bottom() {
        let c = parse_xml(BAR_XML, "c.xml").unwrap();
        assert_eq!(
            c.axes
                .iter()
                .find(|a| a.axis_type == AxisType::Category)
                .unwrap()
                .position,
            Some(AxisPosition::Bottom)
        );
    }
    #[test]
    fn val_axis_numfmt() {
        let c = parse_xml(BAR_XML, "c.xml").unwrap();
        assert_eq!(
            c.axes
                .iter()
                .find(|a| a.axis_type == AxisType::Value)
                .unwrap()
                .number_format
                .as_deref(),
            Some("General")
        );
    }
    #[test]
    fn pie_no_axes() {
        assert!(parse_xml(PIE_XML, "c.xml").unwrap().axes.is_empty());
    }

    // DataValues/StringValues helpers
    #[test]
    fn data_values_is_complete_dense() {
        let mut dv = DataValues::with_capacity(3);
        dv.set(0, 1.0);
        dv.set(1, 2.0);
        dv.set(2, 3.0);
        assert!(dv.is_complete());
    }
    #[test]
    fn data_values_incomplete_with_nan() {
        let mut dv = DataValues::with_capacity(3);
        dv.set(0, 1.0);
        dv.set(2, 3.0);
        assert!(!dv.is_complete());
    }
    #[test]
    fn string_values_complete() {
        let mut sv = StringValues::with_capacity(2);
        sv.set(0, "A".into());
        sv.set(1, "B".into());
        assert!(sv.is_complete());
    }

    #[test]
    fn chart_path_preserved() {
        assert_eq!(
            parse_xml(BAR_XML, "xl/charts/chart1.xml")
                .unwrap()
                .chart_path,
            "xl/charts/chart1.xml"
        );
    }

    // ── Phase 5: Fill / color parsing ─────────────────────────────────────────

    const FILL_CHART_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:spPr>
    <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
  </c:spPr>
  <c:chart>
    <c:plotArea>
      <c:spPr>
        <a:solidFill><a:schemeClr val="accent2"/></a:solidFill>
      </c:spPr>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:tx><c:strRef><c:f>S!$A$1</c:f></c:strRef></c:tx>
          <c:cat><c:strRef><c:f>S!$A$2:$A$3</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>S!$B$2:$B$3</c:f></c:numRef></c:val>
          <c:spPr>
            <a:solidFill>
              <a:schemeClr val="accent1">
                <a:lumMod val="75000"/>
              </a:schemeClr>
            </a:solidFill>
          </c:spPr>
        </c:ser>
        <c:ser>
          <c:idx val="1"/><c:order val="1"/>
          <c:tx><c:strRef><c:f>S!$C$1</c:f></c:strRef></c:tx>
          <c:cat><c:strRef><c:f>S!$A$2:$A$3</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>S!$C$2:$C$3</c:f></c:numRef></c:val>
          <c:spPr>
            <a:gradFill>
              <a:gsLst>
                <a:gs pos="0">
                  <a:srgbClr val="FF0000"/>
                </a:gs>
                <a:gs pos="100000">
                  <a:srgbClr val="0000FF"/>
                </a:gs>
              </a:gsLst>
              <a:lin ang="5400000" scaled="0"/>
            </a:gradFill>
          </c:spPr>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
      <c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

    fn fill_chart() -> crate::model::chart::Chart {
        parse_xml(FILL_CHART_XML, "xl/charts/chart1.xml").unwrap()
    }

    // chart-space fill
    #[test]
    fn chart_fill_solid_white() {
        let c = fill_chart();
        match c.chart_fill.as_ref().unwrap() {
            Fill::Solid(ColorSpec::Srgb(rgb, mods)) => {
                assert_eq!(*rgb, Rgb::WHITE);
                assert!(mods.is_empty());
            }
            other => panic!("Expected Solid(Srgb(WHITE)), got {other:?}"),
        }
    }

    // plot area fill — schemeClr
    #[test]
    fn plot_area_fill_scheme() {
        let c = fill_chart();
        match c.plot_area.fill.as_ref().unwrap() {
            Fill::Solid(ColorSpec::Scheme(slot, mods)) => {
                assert_eq!(*slot, ThemeColorSlot::Accent2);
                assert!(mods.is_empty());
            }
            other => panic!("Expected Solid(Scheme(accent2)), got {other:?}"),
        }
    }

    // series 0 fill — schemeClr with lumMod
    #[test]
    fn series_solid_fill_scheme_with_mod() {
        let c = fill_chart();
        match c.series[0].fill.as_ref().unwrap() {
            Fill::Solid(ColorSpec::Scheme(slot, mods)) => {
                assert_eq!(*slot, ThemeColorSlot::Accent1);
                assert_eq!(mods, &[ColorMod::LumMod(75_000)]);
            }
            other => panic!("Expected Solid(Scheme(accent1, lumMod)), got {other:?}"),
        }
    }

    // series 0 fill resolves against theme
    #[test]
    fn series_fill_resolves_with_theme() {
        use crate::model::theme::Theme;
        let mut theme = Theme::default();
        theme.set(ThemeColorSlot::Accent1, Rgb::from_hex("4472C4").unwrap());

        let c = fill_chart();
        let spec = match c.series[0].fill.as_ref().unwrap() {
            Fill::Solid(s) => s,
            other => panic!("expected solid, got {other:?}"),
        };
        let resolved = spec.resolve(Some(&theme)).unwrap();
        // accent1 with lumMod 75% should be darker than the base
        let base = Rgb::from_hex("4472C4").unwrap();
        // Just verify it resolves successfully and is different from base
        assert_ne!(resolved, base);
    }

    // series 1 fill — gradient
    #[test]
    fn series_gradient_fill_two_stops() {
        let c = fill_chart();
        match c.series[1].fill.as_ref().unwrap() {
            Fill::Gradient(grad) => {
                assert_eq!(grad.stops.len(), 2);
            }
            other => panic!("Expected Gradient, got {other:?}"),
        }
    }

    #[test]
    fn gradient_stop_positions() {
        let c = fill_chart();
        let grad = match c.series[1].fill.as_ref().unwrap() {
            Fill::Gradient(g) => g,
            _ => panic!(),
        };
        assert_eq!(grad.stops[0].position, 0);
        assert_eq!(grad.stops[1].position, 100_000);
    }

    #[test]
    fn gradient_stop_colors() {
        let c = fill_chart();
        let grad = match c.series[1].fill.as_ref().unwrap() {
            Fill::Gradient(g) => g,
            _ => panic!(),
        };
        assert_eq!(
            grad.stops[0].color,
            ColorSpec::Srgb(Rgb::from_hex("FF0000").unwrap(), vec![])
        );
        assert_eq!(
            grad.stops[1].color,
            ColorSpec::Srgb(Rgb::from_hex("0000FF").unwrap(), vec![])
        );
    }

    #[test]
    fn gradient_linear_direction() {
        let c = fill_chart();
        let grad = match c.series[1].fill.as_ref().unwrap() {
            Fill::Gradient(g) => g,
            _ => panic!(),
        };
        match grad.direction.as_ref().unwrap() {
            GradientDirection::Linear { angle_deg, scaled } => {
                // 5400000 / 60000 = 90 degrees
                assert!((angle_deg - 90.0).abs() < 0.01);
                assert!(!scaled);
            }
            _ => panic!("expected Linear direction"),
        }
    }

    #[test]
    fn gradient_resolve_stops() {
        let c = fill_chart();
        let grad = match c.series[1].fill.as_ref().unwrap() {
            Fill::Gradient(g) => g,
            _ => panic!(),
        };
        let resolved = grad.resolve_stops(None);
        assert_eq!(resolved.len(), 2);
        assert!((resolved[0].0 - 0.0).abs() < 1e-9);
        assert!((resolved[1].0 - 1.0).abs() < 1e-9);
        assert_eq!(resolved[0].1, Rgb::from_hex("FF0000").unwrap());
        assert_eq!(resolved[1].1, Rgb::from_hex("0000FF").unwrap());
    }

    // no fill on series without spPr
    #[test]
    fn series_without_sppr_has_no_fill() {
        let c = parse_xml(BAR_XML, "c.xml").unwrap();
        assert!(c.series[0].fill.is_none());
        assert!(c.series[1].fill.is_none());
    }

    // scheme color resolution end-to-end: plot area + theme
    #[test]
    fn plot_area_scheme_fill_resolves_with_theme() {
        use crate::model::theme::Theme;
        let mut theme = Theme::default();
        theme.set(ThemeColorSlot::Accent2, Rgb::from_hex("ED7D31").unwrap());

        let c = fill_chart();
        let spec = match c.plot_area.fill.as_ref().unwrap() {
            Fill::Solid(s) => s,
            other => panic!("expected solid, got {other:?}"),
        };
        assert_eq!(
            spec.resolve(Some(&theme)),
            Some(Rgb::from_hex("ED7D31").unwrap())
        );
        assert_eq!(spec.resolve(None), None); // schemeClr needs theme
    }

    // ── 3-D view parsing ─────────────────────────────────────────────────────

    const BAR3D_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:view3D>
      <c:rotX val="30"/>
      <c:rotY val="20"/>
      <c:rAngAx val="1"/>
      <c:perspective val="30"/>
    </c:view3D>
    <c:plotArea>
      <c:bar3DChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:val><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f>
            <c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/>
              <c:pt idx="0"><c:v>100</c:v></c:pt>
              <c:pt idx="1"><c:v>200</c:v></c:pt>
              <c:pt idx="2"><c:v>150</c:v></c:pt>
            </c:numCache>
          </c:numRef></c:val>
        </c:ser>
      </c:bar3DChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

    const PIE3D_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:view3D>
      <c:rotX val="15"/>
      <c:rotY val="0"/>
      <c:rAngAx val="0"/>
      <c:perspective val="45"/>
    </c:view3D>
    <c:plotArea>
      <c:pie3DChart>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
      </c:pie3DChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

    const SURFACE3D_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:view3D>
      <c:rotX val="-30"/>
      <c:rotY val="180"/>
    </c:view3D>
    <c:plotArea>
      <c:surface3DChart>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
      </c:surface3DChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

    const HBAR3D_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:view3D><c:rotX val="15"/><c:rotY val="20"/></c:view3D>
    <c:plotArea>
      <c:bar3DChart>
        <c:barDir val="bar"/>
        <c:grouping val="clustered"/>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
      </c:bar3DChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

    const NO_VIEW3D_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

    // ── chart type detection ──────────────────────────────────────────────────

    #[test]
    fn bar3d_chart_type() {
        let c = parse_xml(BAR3D_XML, "c.xml").unwrap();
        assert_eq!(c.chart_type, ChartType::Bar3D);
    }

    #[test]
    fn pie3d_chart_type() {
        let c = parse_xml(PIE3D_XML, "c.xml").unwrap();
        assert_eq!(c.chart_type, ChartType::Pie3D);
    }

    #[test]
    fn surface3d_chart_type() {
        let c = parse_xml(SURFACE3D_XML, "c.xml").unwrap();
        assert_eq!(c.chart_type, ChartType::Surface3D);
    }

    #[test]
    fn hbar3d_chart_type() {
        let c = parse_xml(HBAR3D_XML, "c.xml").unwrap();
        assert_eq!(c.chart_type, ChartType::HorizontalBar3D);
    }

    #[test]
    fn no_view3d_on_2d_chart() {
        let c = parse_xml(NO_VIEW3D_XML, "c.xml").unwrap();
        assert_eq!(c.chart_type, ChartType::Bar);
        assert!(c.view_3d.is_none());
    }

    // ── view_3d presence ─────────────────────────────────────────────────────

    #[test]
    fn bar3d_has_view3d() {
        let c = parse_xml(BAR3D_XML, "c.xml").unwrap();
        assert!(c.view_3d.is_some(), "bar3DChart should have view_3d");
    }

    #[test]
    fn pie3d_has_view3d() {
        let c = parse_xml(PIE3D_XML, "c.xml").unwrap();
        assert!(c.view_3d.is_some());
    }

    #[test]
    fn surface3d_has_view3d() {
        let c = parse_xml(SURFACE3D_XML, "c.xml").unwrap();
        assert!(c.view_3d.is_some());
    }

    // ── rotX ─────────────────────────────────────────────────────────────────

    #[test]
    fn view3d_rotx_bar3d() {
        let v = parse_xml(BAR3D_XML, "c.xml").unwrap().view_3d.unwrap();
        assert_eq!(v.rotation_x, Some(30));
    }

    #[test]
    fn view3d_rotx_pie3d() {
        let v = parse_xml(PIE3D_XML, "c.xml").unwrap().view_3d.unwrap();
        assert_eq!(v.rotation_x, Some(15));
    }

    #[test]
    fn view3d_rotx_negative() {
        let v = parse_xml(SURFACE3D_XML, "c.xml").unwrap().view_3d.unwrap();
        assert_eq!(v.rotation_x, Some(-30));
    }

    // ── rotY ─────────────────────────────────────────────────────────────────

    #[test]
    fn view3d_roty_bar3d() {
        let v = parse_xml(BAR3D_XML, "c.xml").unwrap().view_3d.unwrap();
        assert_eq!(v.rotation_y, Some(20));
    }

    #[test]
    fn view3d_roty_180() {
        let v = parse_xml(SURFACE3D_XML, "c.xml").unwrap().view_3d.unwrap();
        assert_eq!(v.rotation_y, Some(180));
    }

    // ── rAngAx ───────────────────────────────────────────────────────────────

    #[test]
    fn view3d_right_angle_axes_true() {
        let v = parse_xml(BAR3D_XML, "c.xml").unwrap().view_3d.unwrap();
        assert_eq!(v.right_angle_axes, Some(true));
    }

    #[test]
    fn view3d_right_angle_axes_false() {
        let v = parse_xml(PIE3D_XML, "c.xml").unwrap().view_3d.unwrap();
        assert_eq!(v.right_angle_axes, Some(false));
    }

    #[test]
    fn view3d_right_angle_axes_absent() {
        // SURFACE3D_XML has no <c:rAngAx>
        let v = parse_xml(SURFACE3D_XML, "c.xml").unwrap().view_3d.unwrap();
        assert_eq!(v.right_angle_axes, None);
    }

    // ── perspective ──────────────────────────────────────────────────────────

    #[test]
    fn view3d_perspective_bar3d() {
        let v = parse_xml(BAR3D_XML, "c.xml").unwrap().view_3d.unwrap();
        assert_eq!(v.perspective, Some(30));
    }

    #[test]
    fn view3d_perspective_pie3d() {
        let v = parse_xml(PIE3D_XML, "c.xml").unwrap().view_3d.unwrap();
        assert_eq!(v.perspective, Some(45));
    }

    #[test]
    fn view3d_perspective_absent() {
        let v = parse_xml(SURFACE3D_XML, "c.xml").unwrap().view_3d.unwrap();
        assert_eq!(v.perspective, None);
    }

    // ── partial view3D (only rotX + rotY, rest absent) ────────────────────────

    #[test]
    fn view3d_partial_hbar3d_rotx() {
        let v = parse_xml(HBAR3D_XML, "c.xml").unwrap().view_3d.unwrap();
        assert_eq!(v.rotation_x, Some(15));
    }

    #[test]
    fn view3d_partial_hbar3d_roty() {
        let v = parse_xml(HBAR3D_XML, "c.xml").unwrap().view_3d.unwrap();
        assert_eq!(v.rotation_y, Some(20));
    }

    #[test]
    fn view3d_partial_hbar3d_no_rangax() {
        let v = parse_xml(HBAR3D_XML, "c.xml").unwrap().view_3d.unwrap();
        assert_eq!(v.right_angle_axes, None);
    }

    #[test]
    fn view3d_partial_hbar3d_no_perspective() {
        let v = parse_xml(HBAR3D_XML, "c.xml").unwrap().view_3d.unwrap();
        assert_eq!(v.perspective, None);
    }

    // ── cache still works inside bar3DChart ───────────────────────────────────

    #[test]
    fn bar3d_value_cache_intact() {
        let c = parse_xml(BAR3D_XML, "c.xml").unwrap();
        let cache = c.series[0].value_cache.as_ref().unwrap();
        assert_eq!(cache.values, vec![100.0, 200.0, 150.0]);
    }

    // ── 2D chart is_3d() flag ─────────────────────────────────────────────────

    #[test]
    fn bar_chart_is_not_3d() {
        let c = parse_xml(NO_VIEW3D_XML, "c.xml").unwrap();
        assert!(!c.chart_type.is_3d());
    }

    #[test]
    fn bar3d_chart_is_3d() {
        let c = parse_xml(BAR3D_XML, "c.xml").unwrap();
        assert!(c.chart_type.is_3d());
    }

    // ═════════════════════════════════════════════════════════════════════════
    // Phase 9 — 3-D geometry surface tests
    // ═════════════════════════════════════════════════════════════════════════

    // ── XML fixtures ─────────────────────────────────────────────────────────

    /// All three surfaces with explicit fills:
    ///   floor     → solid sRGB #D9D9D9
    ///   sideWall  → solid scheme accent1
    ///   backWall  → gradient (two sRGB stops, linear 90°)
    const ALL_SURFACES_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:view3D><c:rotX val="15"/><c:rotY val="20"/></c:view3D>
    <c:plotArea>
      <c:bar3DChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
      </c:bar3DChart>
    </c:plotArea>
    <c:floor>
      <c:spPr>
        <a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill>
      </c:spPr>
    </c:floor>
    <c:sideWall>
      <c:spPr>
        <a:solidFill>
          <a:schemeClr val="accent1">
            <a:lumMod val="75000"/>
          </a:schemeClr>
        </a:solidFill>
      </c:spPr>
    </c:sideWall>
    <c:backWall>
      <c:spPr>
        <a:gradFill>
          <a:gsLst>
            <a:gs pos="0"><a:srgbClr val="4472C4"/></a:gs>
            <a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs>
          </a:gsLst>
          <a:lin ang="5400000" scaled="0"/>
        </a:gradFill>
      </c:spPr>
    </c:backWall>
  </c:chart>
</c:chartSpace>"#;

    /// Only floor has a fill; sideWall and backWall are absent.
    const FLOOR_ONLY_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:view3D><c:rotX val="30"/></c:view3D>
    <c:plotArea>
      <c:bar3DChart>
        <c:barDir val="col"/>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
      </c:bar3DChart>
    </c:plotArea>
    <c:floor>
      <c:spPr>
        <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
      </c:spPr>
    </c:floor>
  </c:chart>
</c:chartSpace>"#;

    /// Explicit noFill on all three surfaces.
    const NO_FILL_SURFACES_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:view3D><c:rotX val="15"/></c:view3D>
    <c:plotArea>
      <c:bar3DChart>
        <c:barDir val="col"/>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
      </c:bar3DChart>
    </c:plotArea>
    <c:floor>   <c:spPr><a:noFill/></c:spPr></c:floor>
    <c:sideWall><c:spPr><a:noFill/></c:spPr></c:sideWall>
    <c:backWall><c:spPr><a:noFill/></c:spPr></c:backWall>
  </c:chart>
</c:chartSpace>"#;

    /// No surface elements at all — surface should be None.
    const NO_SURFACES_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:view3D><c:rotX val="30"/></c:view3D>
    <c:plotArea>
      <c:bar3DChart>
        <c:barDir val="col"/>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
      </c:bar3DChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

    /// sideWall carries a sysClr; backWall carries a prstClr.
    const SYS_PRESET_SURFACES_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:view3D><c:rotX val="15"/></c:view3D>
    <c:plotArea>
      <c:bar3DChart>
        <c:barDir val="col"/>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
      </c:bar3DChart>
    </c:plotArea>
    <c:sideWall>
      <c:spPr>
        <a:solidFill><a:sysClr lastClr="FFFFFF"/></a:solidFill>
      </c:spPr>
    </c:sideWall>
    <c:backWall>
      <c:spPr>
        <a:solidFill><a:prstClr val="blue"/></a:solidFill>
      </c:spPr>
    </c:backWall>
  </c:chart>
</c:chartSpace>"#;

    // ── surface presence / absence ────────────────────────────────────────────

    #[test]
    fn no_surfaces_gives_none() {
        let c = parse_xml(NO_SURFACES_XML, "c.xml").unwrap();
        assert!(
            c.surface.is_none(),
            "expected surface = None when no surface elements present"
        );
    }

    #[test]
    fn all_surfaces_gives_some() {
        let c = parse_xml(ALL_SURFACES_XML, "c.xml").unwrap();
        assert!(c.surface.is_some());
    }

    #[test]
    fn floor_only_gives_some() {
        let c = parse_xml(FLOOR_ONLY_XML, "c.xml").unwrap();
        assert!(c.surface.is_some());
    }

    #[test]
    fn no_fill_surfaces_gives_some() {
        // noFill is a valid Fill variant, so surface should be Some
        let c = parse_xml(NO_FILL_SURFACES_XML, "c.xml").unwrap();
        assert!(c.surface.is_some());
    }

    // ── floor fill ────────────────────────────────────────────────────────────

    #[test]
    fn floor_solid_srgb_color() {
        let c = parse_xml(ALL_SURFACES_XML, "c.xml").unwrap();
        let fill = c.surface.as_ref().unwrap().floor_fill.as_ref().unwrap();
        match fill {
            Fill::Solid(ColorSpec::Srgb(rgb, mods)) => {
                assert_eq!(*rgb, Rgb::from_hex("D9D9D9").unwrap());
                assert!(mods.is_empty());
            }
            other => panic!("expected Solid(Srgb), got {other:?}"),
        }
    }

    #[test]
    fn floor_only_side_and_back_absent() {
        let c = parse_xml(FLOOR_ONLY_XML, "c.xml").unwrap();
        let s = c.surface.as_ref().unwrap();
        assert!(s.floor_fill.is_some());
        assert!(s.side_wall_fill.is_none());
        assert!(s.back_wall_fill.is_none());
    }

    #[test]
    fn floor_solid_color_floor_only() {
        let c = parse_xml(FLOOR_ONLY_XML, "c.xml").unwrap();
        let fill = c.surface.as_ref().unwrap().floor_fill.as_ref().unwrap();
        match fill {
            Fill::Solid(ColorSpec::Srgb(rgb, _)) => {
                assert_eq!(*rgb, Rgb::from_hex("FF0000").unwrap());
            }
            other => panic!("expected Solid(Srgb), got {other:?}"),
        }
    }

    // ── side-wall fill ────────────────────────────────────────────────────────

    #[test]
    fn side_wall_solid_scheme_color() {
        let c = parse_xml(ALL_SURFACES_XML, "c.xml").unwrap();
        let fill = c.surface.as_ref().unwrap().side_wall_fill.as_ref().unwrap();
        match fill {
            Fill::Solid(ColorSpec::Scheme(slot, mods)) => {
                assert_eq!(*slot, ThemeColorSlot::Accent1);
                assert_eq!(mods, &[ColorMod::LumMod(75_000)]);
            }
            other => panic!("expected Solid(Scheme(Accent1)), got {other:?}"),
        }
    }

    #[test]
    fn side_wall_scheme_resolves_with_theme() {
        use crate::model::theme::Theme;
        let mut theme = Theme::default();
        theme.set(ThemeColorSlot::Accent1, Rgb::from_hex("4472C4").unwrap());

        let c = parse_xml(ALL_SURFACES_XML, "c.xml").unwrap();
        let fill = c.surface.as_ref().unwrap().side_wall_fill.as_ref().unwrap();
        let rgb = fill
            .solid_rgb(Some(&theme))
            .expect("should resolve with theme");
        // accent1 = #4472C4, lumMod 75% darkens it — just check it's not None
        // and not exactly the original (modifier applied)
        assert_ne!(
            rgb,
            Rgb::from_hex("4472C4").unwrap(),
            "lumMod(75000) should have darkened the color"
        );
    }

    #[test]
    fn side_wall_scheme_needs_theme() {
        let c = parse_xml(ALL_SURFACES_XML, "c.xml").unwrap();
        let fill = c.surface.as_ref().unwrap().side_wall_fill.as_ref().unwrap();
        // Without theme, scheme color cannot resolve
        assert!(fill.solid_rgb(None).is_none());
    }

    #[test]
    fn side_wall_sys_color() {
        let c = parse_xml(SYS_PRESET_SURFACES_XML, "c.xml").unwrap();
        let fill = c.surface.as_ref().unwrap().side_wall_fill.as_ref().unwrap();
        match fill {
            Fill::Solid(ColorSpec::Sys(rgb, _)) => {
                assert_eq!(*rgb, Rgb::from_hex("FFFFFF").unwrap());
            }
            other => panic!("expected Solid(Sys), got {other:?}"),
        }
    }

    // ── back-wall fill ────────────────────────────────────────────────────────

    #[test]
    fn back_wall_gradient_stops() {
        let c = parse_xml(ALL_SURFACES_XML, "c.xml").unwrap();
        let fill = c.surface.as_ref().unwrap().back_wall_fill.as_ref().unwrap();
        match fill {
            Fill::Gradient(grad) => {
                assert_eq!(grad.stops.len(), 2);
                assert_eq!(grad.stops[0].position, 0);
                assert_eq!(grad.stops[1].position, 100_000);
            }
            other => panic!("expected Gradient, got {other:?}"),
        }
    }

    #[test]
    fn back_wall_gradient_stop0_color() {
        let c = parse_xml(ALL_SURFACES_XML, "c.xml").unwrap();
        let fill = c.surface.as_ref().unwrap().back_wall_fill.as_ref().unwrap();
        if let Fill::Gradient(grad) = fill {
            match &grad.stops[0].color {
                ColorSpec::Srgb(rgb, _) => assert_eq!(*rgb, Rgb::from_hex("4472C4").unwrap()),
                other => panic!("expected Srgb, got {other:?}"),
            }
        }
    }

    #[test]
    fn back_wall_gradient_stop1_color() {
        let c = parse_xml(ALL_SURFACES_XML, "c.xml").unwrap();
        let fill = c.surface.as_ref().unwrap().back_wall_fill.as_ref().unwrap();
        if let Fill::Gradient(grad) = fill {
            match &grad.stops[1].color {
                ColorSpec::Srgb(rgb, _) => assert_eq!(*rgb, Rgb::from_hex("FFFFFF").unwrap()),
                other => panic!("expected Srgb, got {other:?}"),
            }
        }
    }

    #[test]
    fn back_wall_gradient_linear_direction() {
        let c = parse_xml(ALL_SURFACES_XML, "c.xml").unwrap();
        let fill = c.surface.as_ref().unwrap().back_wall_fill.as_ref().unwrap();
        if let Fill::Gradient(grad) = fill {
            match grad.direction.as_ref().unwrap() {
                GradientDirection::Linear { angle_deg, scaled } => {
                    // ang=5400000 → 5400000/60000 = 90°
                    assert!((angle_deg - 90.0).abs() < 0.01);
                    assert!(!scaled);
                }
                other => panic!("expected Linear, got {other:?}"),
            }
        }
    }

    #[test]
    fn back_wall_preset_color() {
        let c = parse_xml(SYS_PRESET_SURFACES_XML, "c.xml").unwrap();
        let fill = c.surface.as_ref().unwrap().back_wall_fill.as_ref().unwrap();
        match fill {
            Fill::Solid(ColorSpec::Preset(name, _)) => {
                assert_eq!(name, "blue");
            }
            other => panic!("expected Solid(Preset), got {other:?}"),
        }
    }

    // ── noFill surfaces ───────────────────────────────────────────────────────

    #[test]
    fn floor_no_fill() {
        let c = parse_xml(NO_FILL_SURFACES_XML, "c.xml").unwrap();
        let fill = c.surface.as_ref().unwrap().floor_fill.as_ref().unwrap();
        assert_eq!(fill, &Fill::None);
    }

    #[test]
    fn side_wall_no_fill() {
        let c = parse_xml(NO_FILL_SURFACES_XML, "c.xml").unwrap();
        let fill = c.surface.as_ref().unwrap().side_wall_fill.as_ref().unwrap();
        assert_eq!(fill, &Fill::None);
    }

    #[test]
    fn back_wall_no_fill() {
        let c = parse_xml(NO_FILL_SURFACES_XML, "c.xml").unwrap();
        let fill = c.surface.as_ref().unwrap().back_wall_fill.as_ref().unwrap();
        assert_eq!(fill, &Fill::None);
    }

    // ── isolation — surface fill doesn't bleed into series or chart fill ──────

    #[test]
    fn surface_fill_does_not_bleed_into_chart_fill() {
        // ALL_SURFACES_XML has no chart-level spPr → chart_fill should be None
        let c = parse_xml(ALL_SURFACES_XML, "c.xml").unwrap();
        assert!(
            c.chart_fill.is_none(),
            "surface fills must not bleed into chart_fill"
        );
    }

    #[test]
    fn surface_fill_does_not_bleed_into_series_fill() {
        let c = parse_xml(ALL_SURFACES_XML, "c.xml").unwrap();
        for s in &c.series {
            assert!(
                s.fill.is_none(),
                "surface fills must not bleed into series fill"
            );
        }
    }

    #[test]
    fn surface_fill_does_not_bleed_into_plot_area_fill() {
        let c = parse_xml(ALL_SURFACES_XML, "c.xml").unwrap();
        assert!(
            c.plot_area.fill.is_none(),
            "surface fills must not bleed into plot_area.fill"
        );
    }

    // ── view_3d unaffected by surface parsing ─────────────────────────────────

    #[test]
    fn surfaces_and_view3d_coexist() {
        let c = parse_xml(ALL_SURFACES_XML, "c.xml").unwrap();
        let v = c.view_3d.as_ref().expect("view_3d should be Some");
        assert_eq!(v.rotation_x, Some(15));
        assert_eq!(v.rotation_y, Some(20));
        assert!(c.surface.is_some());
    }

    // ── 2-D chart has no surface ──────────────────────────────────────────────

    #[test]
    fn two_d_bar_chart_surface_is_none() {
        let c = parse_xml(BAR_XML, "c.xml").unwrap();
        assert!(c.surface.is_none());
    }

    // ═════════════════════════════════════════════════════════════════════════
    // Phase 10 — Pivot chart detection tests
    // ═════════════════════════════════════════════════════════════════════════

    // ── XML fixtures ─────────────────────────────────────────────────────────

    /// Pivot bar chart with a fully-qualified pivot source name.
    const PIVOT_BAR_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:pivotSource>
    <c:name>Sheet1!PivotTable1</c:name>
    <c:fmtId val="0"/>
  </c:pivotSource>
  <c:chart>
    <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/>
      <a:p><a:r><a:t>Pivot Chart</a:t></a:r></a:p>
    </c:rich></c:tx><c:overlay val="0"/></c:title>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f></c:strRef></c:tx>
          <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$4</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet1!$B$2:$B$4</c:f>
            <c:numCache>
              <c:ptCount val="3"/>
              <c:pt idx="0"><c:v>100</c:v></c:pt>
              <c:pt idx="1"><c:v>200</c:v></c:pt>
              <c:pt idx="2"><c:v>150</c:v></c:pt>
            </c:numCache>
          </c:numRef></c:val>
        </c:ser>
        <c:axId val="1"/><c:axId val="2"/>
      </c:barChart>
      <c:catAx><c:axId val="1"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
      <c:valAx><c:axId val="2"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="b"/></c:legend>
  </c:chart>
</c:chartSpace>"#;

    /// Pivot chart where the pivot table name has no sheet prefix.
    const PIVOT_NO_SHEET_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:pivotSource>
    <c:name>SalesData</c:name>
    <c:fmtId val="1"/>
  </c:pivotSource>
  <c:chart>
    <c:plotArea>
      <c:pieChart>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
      </c:pieChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

    /// Pivot chart with a series that has an inline <c:name> text node —
    /// confirms the series name does NOT pollute pivot_table_name.
    const PIVOT_WITH_SERIES_NAME_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:pivotSource>
    <c:name>Sheet2!PivotTable2</c:name>
    <c:fmtId val="0"/>
  </c:pivotSource>
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:tx><c:strRef><c:f>Sheet2!$B$1</c:f>
            <c:strCache>
              <c:ptCount val="1"/>
              <c:pt idx="0"><c:v>Revenue</c:v></c:pt>
            </c:strCache>
          </c:strRef></c:tx>
          <c:val><c:numRef><c:f>Sheet2!$B$2:$B$3</c:f></c:numRef></c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

    /// Non-pivot chart — no <c:pivotSource> at all.
    const NON_PIVOT_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

    /// <c:pivotSource> present but <c:name> child is empty — edge case.
    const PIVOT_EMPTY_NAME_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:pivotSource>
    <c:name></c:name>
    <c:fmtId val="0"/>
  </c:pivotSource>
  <c:chart>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:ser><c:idx val="0"/><c:order val="0"/></c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#;

    // ── is_pivot_chart ────────────────────────────────────────────────────────

    #[test]
    fn pivot_bar_is_pivot_chart() {
        assert!(parse_xml(PIVOT_BAR_XML, "c.xml").unwrap().is_pivot_chart);
    }

    #[test]
    fn pivot_no_sheet_is_pivot_chart() {
        assert!(
            parse_xml(PIVOT_NO_SHEET_XML, "c.xml")
                .unwrap()
                .is_pivot_chart
        );
    }

    #[test]
    fn pivot_with_series_name_is_pivot_chart() {
        assert!(
            parse_xml(PIVOT_WITH_SERIES_NAME_XML, "c.xml")
                .unwrap()
                .is_pivot_chart
        );
    }

    #[test]
    fn non_pivot_is_not_pivot_chart() {
        assert!(!parse_xml(NON_PIVOT_XML, "c.xml").unwrap().is_pivot_chart);
    }

    #[test]
    fn regular_bar_not_pivot() {
        assert!(!parse_xml(BAR_XML, "c.xml").unwrap().is_pivot_chart);
    }

    #[test]
    fn regular_line_not_pivot() {
        assert!(!parse_xml(LINE_XML, "c.xml").unwrap().is_pivot_chart);
    }

    // ── pivot_table_name ──────────────────────────────────────────────────────

    #[test]
    fn pivot_bar_name_full() {
        let c = parse_xml(PIVOT_BAR_XML, "c.xml").unwrap();
        assert_eq!(c.pivot_table_name.as_deref(), Some("Sheet1!PivotTable1"));
    }

    #[test]
    fn pivot_no_sheet_name_bare() {
        let c = parse_xml(PIVOT_NO_SHEET_XML, "c.xml").unwrap();
        assert_eq!(c.pivot_table_name.as_deref(), Some("SalesData"));
    }

    #[test]
    fn pivot_with_series_name_correct() {
        // pivot_table_name must be "Sheet2!PivotTable2", not "Revenue"
        let c = parse_xml(PIVOT_WITH_SERIES_NAME_XML, "c.xml").unwrap();
        assert_eq!(c.pivot_table_name.as_deref(), Some("Sheet2!PivotTable2"));
    }

    #[test]
    fn non_pivot_name_is_none() {
        let c = parse_xml(NON_PIVOT_XML, "c.xml").unwrap();
        assert!(c.pivot_table_name.is_none());
    }

    #[test]
    fn regular_bar_name_is_none() {
        let c = parse_xml(BAR_XML, "c.xml").unwrap();
        assert!(c.pivot_table_name.is_none());
    }

    // ── empty <c:name> edge case ──────────────────────────────────────────────

    #[test]
    fn pivot_empty_name_is_not_pivot() {
        // Empty <c:name> → no name extracted → treated as non-pivot
        let c = parse_xml(PIVOT_EMPTY_NAME_XML, "c.xml").unwrap();
        assert!(
            !c.is_pivot_chart,
            "empty <c:name> should not mark chart as pivot"
        );
    }

    #[test]
    fn pivot_empty_name_name_is_none() {
        let c = parse_xml(PIVOT_EMPTY_NAME_XML, "c.xml").unwrap();
        assert!(c.pivot_table_name.is_none());
    }

    // ── isolation — series name does not leak into pivot_table_name ───────────

    #[test]
    fn series_name_does_not_pollute_pivot_name() {
        let c = parse_xml(PIVOT_WITH_SERIES_NAME_XML, "c.xml").unwrap();
        // The series cache has name "Revenue" — must not appear in pivot_table_name
        let name = c.pivot_table_name.as_deref().unwrap_or("");
        assert_ne!(
            name, "Revenue",
            "series name 'Revenue' must not bleed into pivot_table_name"
        );
    }

    // ── other chart fields unaffected ─────────────────────────────────────────

    #[test]
    fn pivot_chart_type_still_bar() {
        let c = parse_xml(PIVOT_BAR_XML, "c.xml").unwrap();
        assert_eq!(c.chart_type, ChartType::Bar);
    }

    #[test]
    fn pivot_chart_title_still_parsed() {
        let c = parse_xml(PIVOT_BAR_XML, "c.xml").unwrap();
        assert_eq!(c.title.as_deref(), Some("Pivot Chart"));
    }

    #[test]
    fn pivot_chart_legend_still_parsed() {
        let c = parse_xml(PIVOT_BAR_XML, "c.xml").unwrap();
        assert_eq!(c.legend_position, Some(LegendPosition::Bottom));
    }

    #[test]
    fn pivot_chart_series_cache_intact() {
        let c = parse_xml(PIVOT_BAR_XML, "c.xml").unwrap();
        let cache = c.series[0]
            .value_cache
            .as_ref()
            .expect("value cache should be present");
        assert_eq!(cache.values, vec![100.0, 200.0, 150.0]);
    }

    #[test]
    fn pivot_chart_surface_is_none() {
        let c = parse_xml(PIVOT_BAR_XML, "c.xml").unwrap();
        assert!(c.surface.is_none());
    }

    #[test]
    fn pivot_chart_view3d_is_none() {
        let c = parse_xml(PIVOT_BAR_XML, "c.xml").unwrap();
        assert!(c.view_3d.is_none());
    }
}
