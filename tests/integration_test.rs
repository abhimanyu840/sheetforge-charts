//! Integration tests — Phases 1–10.
//!
//! Fixtures:
//!   tests/fixtures/test_charts.xlsx      (Phases 1–7, 2-D charts)
//!     Sheet "Sales"    → chart1.xml  (bar,  2 series, theme fills)
//!     Sheet "Expenses" → chart2.xml  (line, 1 series, no fill)
//!     Sheet "Data"     → no chart
//!     Theme: Office Theme (accent1 = #4472C4, accent2 = #ED7D31, …)
//!     Anchors: both charts occupy rows 0–15, cols 0–8
//!
//!   tests/fixtures/test_3d_charts.xlsx   (Phase 8, 3-D charts)
//!     Sheet "Bar3D"     → chart1.xml  bar3DChart     rotX=30  rotY=20  rAngAx=1 persp=30
//!     Sheet "Line3D"    → chart2.xml  line3DChart    rotX=15  rotY=10  rAngAx=1 persp=0
//!     Sheet "Area3D"    → chart3.xml  area3DChart    rotX=10  rotY=5   rAngAx=0 persp=45
//!     Sheet "Pie3D"     → chart4.xml  pie3DChart     rotX=15  rotY=0   rAngAx=0 persp=45
//!     Sheet "Surface3D" → chart5.xml  surface3DChart rotX=-30 rotY=180 (no rAngAx/persp)
//!     Sheet "HBar3D"    → chart6.xml  bar3DChart(bar)rotX=20  rotY=15  rAngAx=1
//!
//!   tests/fixtures/test_surface_charts.xlsx  (Phase 9, surface geometry fills)
//!     Sheet "AllSurfaces"    → chart1.xml  floor solid #D9D9D9, sideWall solid #4472C4,
//!                                          backWall gradient #FF0000→#FFFFFF linear 90°
//!     Sheet "FloorOnly"      → chart2.xml  floor solid #FF0000; sideWall+backWall absent
//!     Sheet "NoFillSurfaces" → chart3.xml  all three surfaces carry explicit <a:noFill/>
//!
//!   tests/fixtures/test_pivot_charts.xlsx  (Phase 10, pivot chart detection)
//!     Sheet "PivotChart"  → chart1.xml  bar chart + <pivotSource> "Sheet1!PivotTable1"
//!     Sheet "NoPivot"     → chart2.xml  bar chart, no <pivotSource>
//!     Sheet "MultiPivot"  → chart3.xml  bar chart + <pivotSource> "Sales!RevenueByRegion"

use sheetforge_charts::{
    extract_charts,
    model::{
        axis::{AxisPosition, AxisType},
        chart::{Chart3DSurface, Chart3DView, ChartAnchor, ChartType, Grouping, LegendPosition},
        color::{ColorSpec, Fill, GradientDirection, Rgb, ThemeColorSlot},
        series::CacheState,
    },
};

fn fixture() -> String {
    std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/test_charts.xlsx")
        .to_string_lossy()
        .into_owned()
}

// ── Workbook ──────────────────────────────────────────────────────────────────

#[test]
fn workbook_has_three_sheets() {
    let wb = extract_charts(&fixture()).unwrap();
    assert_eq!(wb.sheets.len(), 3);
}

#[test]
fn total_chart_count_is_two() {
    let wb = extract_charts(&fixture()).unwrap();
    assert_eq!(wb.chart_count(), 2);
}

#[test]
fn data_sheet_has_no_charts() {
    let wb = extract_charts(&fixture()).unwrap();
    assert!(wb.sheets[2].charts.is_empty());
}

// ── Chart types ───────────────────────────────────────────────────────────────

#[test]
fn sales_chart_is_bar() {
    let wb = extract_charts(&fixture()).unwrap();
    assert_eq!(wb.sheets[0].charts[0].chart_type, ChartType::Bar);
}

#[test]
fn expenses_chart_is_line() {
    let wb = extract_charts(&fixture()).unwrap();
    assert_eq!(wb.sheets[1].charts[0].chart_type, ChartType::Line);
}

// ── Titles ────────────────────────────────────────────────────────────────────

#[test]
fn sales_chart_title() {
    let wb = extract_charts(&fixture()).unwrap();
    assert_eq!(
        wb.sheets[0].charts[0].title.as_deref(),
        Some("Sales Overview")
    );
}

#[test]
fn expenses_chart_title() {
    let wb = extract_charts(&fixture()).unwrap();
    assert_eq!(
        wb.sheets[1].charts[0].title.as_deref(),
        Some("Monthly Expenses")
    );
}

// ── Legend ────────────────────────────────────────────────────────────────────

#[test]
fn sales_chart_legend_field_accessible() {
    // The fixture does not declare a legend — field must be None without panicking.
    let wb = extract_charts(&fixture()).unwrap();
    assert!(wb.sheets[0].charts[0].legend_position.is_none());
}

// ── Series count ──────────────────────────────────────────────────────────────

#[test]
fn sales_chart_has_two_series() {
    let wb = extract_charts(&fixture()).unwrap();
    assert_eq!(wb.sheets[0].charts[0].series.len(), 2);
}

#[test]
fn expenses_chart_has_one_series() {
    let wb = extract_charts(&fixture()).unwrap();
    assert_eq!(wb.sheets[1].charts[0].series.len(), 1);
}

// ── Series name references ────────────────────────────────────────────────────

#[test]
fn sales_series0_name_ref() {
    let wb = extract_charts(&fixture()).unwrap();
    let s = &wb.sheets[0].charts[0].series[0];
    assert_eq!(
        s.name_ref.as_ref().map(|r| r.formula.as_str()),
        Some("Sales!$B$1")
    );
}

#[test]
fn sales_series1_name_ref() {
    let wb = extract_charts(&fixture()).unwrap();
    let s = &wb.sheets[0].charts[0].series[1];
    assert_eq!(
        s.name_ref.as_ref().map(|r| r.formula.as_str()),
        Some("Sales!$C$1")
    );
}

// ── Category references ───────────────────────────────────────────────────────

#[test]
fn sales_series0_category_ref() {
    let wb = extract_charts(&fixture()).unwrap();
    let s = &wb.sheets[0].charts[0].series[0];
    assert_eq!(
        s.category_ref.as_ref().map(|r| r.formula.as_str()),
        Some("Sales!$A$2:$A$4")
    );
}

// ── Value references ──────────────────────────────────────────────────────────

#[test]
fn sales_series0_value_ref() {
    let wb = extract_charts(&fixture()).unwrap();
    let s = &wb.sheets[0].charts[0].series[0];
    assert_eq!(
        s.value_ref.as_ref().map(|r| r.formula.as_str()),
        Some("Sales!$B$2:$B$4")
    );
}

#[test]
fn expenses_series0_value_ref() {
    let wb = extract_charts(&fixture()).unwrap();
    let s = &wb.sheets[1].charts[0].series[0];
    assert_eq!(
        s.value_ref.as_ref().map(|r| r.formula.as_str()),
        Some("Expenses!$B$2:$B$3")
    );
}

// ── Numeric cache ─────────────────────────────────────────────────────────────

#[test]
fn sales_series0_value_cache() {
    let wb = extract_charts(&fixture()).unwrap();
    let s = &wb.sheets[0].charts[0].series[0];
    let cache = s
        .value_cache
        .as_ref()
        .expect("value cache should be present");
    assert_eq!(cache.values, vec![1000.0, 1500.0, 1200.0]);
}

#[test]
fn sales_series0_cache_state_complete() {
    let wb = extract_charts(&fixture()).unwrap();
    let s = &wb.sheets[0].charts[0].series[0];
    assert_eq!(s.value_cache_state, CacheState::Complete);
}

#[test]
fn sales_series1_value_cache() {
    let wb = extract_charts(&fixture()).unwrap();
    let s = &wb.sheets[0].charts[0].series[1];
    let cache = s
        .value_cache
        .as_ref()
        .expect("value cache should be present");
    assert_eq!(cache.values, vec![50.0, 75.0, 60.0]);
}

#[test]
fn expenses_series0_value_cache() {
    let wb = extract_charts(&fixture()).unwrap();
    let s = &wb.sheets[1].charts[0].series[0];
    let cache = s
        .value_cache
        .as_ref()
        .expect("value cache should be present");
    assert_eq!(cache.values, vec![800.0, 950.0]);
}

#[test]
fn expenses_series0_pt_count() {
    let wb = extract_charts(&fixture()).unwrap();
    let s = &wb.sheets[1].charts[0].series[0];
    let cache = s.value_cache.as_ref().unwrap();
    assert_eq!(cache.pt_count, Some(2));
}

// ── Axes ──────────────────────────────────────────────────────────────────────

#[test]
fn sales_chart_has_two_axes() {
    let wb = extract_charts(&fixture()).unwrap();
    assert_eq!(wb.sheets[0].charts[0].axes.len(), 2);
}

#[test]
fn sales_chart_has_category_and_value_axes() {
    let wb = extract_charts(&fixture()).unwrap();
    let axes = &wb.sheets[0].charts[0].axes;
    assert!(axes.iter().any(|a| a.axis_type == AxisType::Category));
    assert!(axes.iter().any(|a| a.axis_type == AxisType::Value));
}

#[test]
fn category_axis_is_at_bottom() {
    let wb = extract_charts(&fixture()).unwrap();
    let axes = &wb.sheets[0].charts[0].axes;
    let cat = axes
        .iter()
        .find(|a| a.axis_type == AxisType::Category)
        .unwrap();
    assert_eq!(cat.position, Some(AxisPosition::Bottom));
}

#[test]
fn value_axis_is_at_left() {
    let wb = extract_charts(&fixture()).unwrap();
    let axes = &wb.sheets[0].charts[0].axes;
    let val = axes
        .iter()
        .find(|a| a.axis_type == AxisType::Value)
        .unwrap();
    assert_eq!(val.position, Some(AxisPosition::Left));
}

// ── Grouping ──────────────────────────────────────────────────────────────────

#[test]
fn sales_chart_grouping_clustered() {
    let wb = extract_charts(&fixture()).unwrap();
    assert_eq!(
        wb.sheets[0].charts[0].plot_area.grouping,
        Some(Grouping::Clustered)
    );
}

#[test]
fn expenses_chart_grouping_standard() {
    let wb = extract_charts(&fixture()).unwrap();
    assert_eq!(
        wb.sheets[1].charts[0].plot_area.grouping,
        Some(Grouping::Standard)
    );
}

// ── Phase 5: Theme ────────────────────────────────────────────────────────────

#[test]
fn theme_is_present() {
    let wb = extract_charts(&fixture()).unwrap();
    assert!(wb.theme.is_some(), "fixture should contain a theme part");
}

#[test]
fn theme_name_is_office() {
    let wb = extract_charts(&fixture()).unwrap();
    let theme = wb.theme.as_ref().unwrap();
    assert_eq!(theme.name.as_deref(), Some("Office Theme"));
}

#[test]
fn theme_accent1_is_4472c4() {
    let wb = extract_charts(&fixture()).unwrap();
    let theme = wb.theme.as_ref().unwrap();
    assert_eq!(theme.accent1(), Some(Rgb::from_hex("4472C4").unwrap()));
}

#[test]
fn theme_accent2_is_ed7d31() {
    let wb = extract_charts(&fixture()).unwrap();
    let theme = wb.theme.as_ref().unwrap();
    assert_eq!(theme.accent2(), Some(Rgb::from_hex("ED7D31").unwrap()));
}

#[test]
fn theme_has_all_12_slots() {
    let wb = extract_charts(&fixture()).unwrap();
    let theme = wb.theme.as_ref().unwrap();
    assert_eq!(theme.all_colors().len(), 12);
}

#[test]
fn theme_dk1_is_black() {
    let wb = extract_charts(&fixture()).unwrap();
    let theme = wb.theme.as_ref().unwrap();
    assert_eq!(theme.dk1(), Some(Rgb::BLACK));
}

#[test]
fn theme_lt1_is_white() {
    let wb = extract_charts(&fixture()).unwrap();
    let theme = wb.theme.as_ref().unwrap();
    assert_eq!(theme.lt1(), Some(Rgb::WHITE));
}

// ── Phase 5: Series solid fill (schemeClr) ────────────────────────────────────

#[test]
fn sales_series0_has_solid_fill() {
    let wb = extract_charts(&fixture()).unwrap();
    let s = &wb.sheets[0].charts[0].series[0];
    assert!(
        matches!(s.fill, Some(Fill::Solid(_))),
        "series 0 should have a solid fill"
    );
}

#[test]
fn sales_series0_fill_is_scheme_accent1() {
    let wb = extract_charts(&fixture()).unwrap();
    let s = &wb.sheets[0].charts[0].series[0];
    match s.fill.as_ref().unwrap() {
        Fill::Solid(ColorSpec::Scheme(slot, _)) => {
            assert_eq!(*slot, ThemeColorSlot::Accent1);
        }
        other => panic!("expected Solid(Scheme(accent1)), got {other:?}"),
    }
}

#[test]
fn sales_series0_fill_resolves_to_4472c4() {
    let wb = extract_charts(&fixture()).unwrap();
    let theme = wb.theme.as_ref().unwrap();
    let s = &wb.sheets[0].charts[0].series[0];
    // accent1 with lumMod 100000 (identity modifier) = base color unchanged
    let rgb = s.fill.as_ref().unwrap().solid_rgb(Some(theme)).unwrap();
    assert_eq!(rgb, Rgb::from_hex("4472C4").unwrap());
}

#[test]
fn sales_series0_fill_needs_theme_to_resolve() {
    let wb = extract_charts(&fixture()).unwrap();
    let s = &wb.sheets[0].charts[0].series[0];
    // schemeClr cannot resolve without a theme
    assert_eq!(s.fill.as_ref().unwrap().solid_rgb(None), None);
}

// ── Phase 5: Series gradient fill ────────────────────────────────────────────

#[test]
fn sales_series1_has_gradient_fill() {
    let wb = extract_charts(&fixture()).unwrap();
    let s = &wb.sheets[0].charts[0].series[1];
    assert!(
        matches!(s.fill, Some(Fill::Gradient(_))),
        "series 1 should have a gradient fill"
    );
}

#[test]
fn sales_series1_gradient_has_two_stops() {
    let wb = extract_charts(&fixture()).unwrap();
    let s = &wb.sheets[0].charts[0].series[1];
    let grad = match s.fill.as_ref().unwrap() {
        Fill::Gradient(g) => g,
        _ => panic!("expected gradient"),
    };
    assert_eq!(grad.stops.len(), 2);
}

#[test]
fn sales_series1_gradient_stop0_color() {
    let wb = extract_charts(&fixture()).unwrap();
    let s = &wb.sheets[0].charts[0].series[1];
    let grad = match s.fill.as_ref().unwrap() {
        Fill::Gradient(g) => g,
        _ => panic!(),
    };
    assert_eq!(
        grad.stops[0].color,
        ColorSpec::Srgb(Rgb::from_hex("4472C4").unwrap(), vec![])
    );
}

#[test]
fn sales_series1_gradient_stop1_is_white() {
    let wb = extract_charts(&fixture()).unwrap();
    let s = &wb.sheets[0].charts[0].series[1];
    let grad = match s.fill.as_ref().unwrap() {
        Fill::Gradient(g) => g,
        _ => panic!(),
    };
    assert_eq!(grad.stops[1].color, ColorSpec::Srgb(Rgb::WHITE, vec![]));
}

#[test]
fn sales_series1_gradient_stop_positions() {
    let wb = extract_charts(&fixture()).unwrap();
    let s = &wb.sheets[0].charts[0].series[1];
    let grad = match s.fill.as_ref().unwrap() {
        Fill::Gradient(g) => g,
        _ => panic!(),
    };
    assert_eq!(grad.stops[0].position, 0);
    assert_eq!(grad.stops[1].position, 100_000);
}

#[test]
fn sales_series1_gradient_direction_is_linear_90deg() {
    let wb = extract_charts(&fixture()).unwrap();
    let s = &wb.sheets[0].charts[0].series[1];
    let grad = match s.fill.as_ref().unwrap() {
        Fill::Gradient(g) => g,
        _ => panic!(),
    };
    match grad.direction.as_ref().unwrap() {
        GradientDirection::Linear { angle_deg, scaled } => {
            // 5400000 / 60000 = 90 degrees
            assert!(
                (angle_deg - 90.0).abs() < 0.01,
                "angle should be 90°, got {angle_deg}"
            );
            assert!(!scaled);
        }
        other => panic!("expected Linear, got {other:?}"),
    }
}

#[test]
fn sales_series1_gradient_resolves_stops() {
    let wb = extract_charts(&fixture()).unwrap();
    let s = &wb.sheets[0].charts[0].series[1];
    let grad = match s.fill.as_ref().unwrap() {
        Fill::Gradient(g) => g,
        _ => panic!(),
    };
    // srgbClr stops resolve without a theme
    let resolved = grad.resolve_stops(None);
    assert_eq!(resolved.len(), 2);
    assert!((resolved[0].0 - 0.0).abs() < 1e-9);
    assert!((resolved[1].0 - 1.0).abs() < 1e-9);
    assert_eq!(resolved[0].1, Rgb::from_hex("4472C4").unwrap());
    assert_eq!(resolved[1].1, Rgb::WHITE);
}

// ── Phase 5: Expenses series has no fill (no spPr) ───────────────────────────

#[test]
fn expenses_series0_has_no_fill() {
    let wb = extract_charts(&fixture()).unwrap();
    let s = &wb.sheets[1].charts[0].series[0];
    assert!(
        s.fill.is_none(),
        "expenses series has no spPr so fill should be None"
    );
}

// ── Phase 6: Drawing anchors ──────────────────────────────────────────────────

#[test]
fn sales_chart_has_anchor() {
    let wb = extract_charts(&fixture()).unwrap();
    assert!(
        wb.sheets[0].charts[0].anchor.is_some(),
        "Sales chart should have an anchor parsed from drawing1.xml"
    );
}

#[test]
fn expenses_chart_has_anchor() {
    let wb = extract_charts(&fixture()).unwrap();
    assert!(
        wb.sheets[1].charts[0].anchor.is_some(),
        "Expenses chart should have an anchor parsed from drawing2.xml"
    );
}

#[test]
fn sales_anchor_col_start() {
    let wb = extract_charts(&fixture()).unwrap();
    assert_eq!(wb.sheets[0].charts[0].anchor.as_ref().unwrap().col_start, 0);
}

#[test]
fn sales_anchor_row_start() {
    let wb = extract_charts(&fixture()).unwrap();
    assert_eq!(wb.sheets[0].charts[0].anchor.as_ref().unwrap().row_start, 0);
}

#[test]
fn sales_anchor_col_end() {
    let wb = extract_charts(&fixture()).unwrap();
    assert_eq!(wb.sheets[0].charts[0].anchor.as_ref().unwrap().col_end, 8);
}

#[test]
fn sales_anchor_row_end() {
    let wb = extract_charts(&fixture()).unwrap();
    assert_eq!(wb.sheets[0].charts[0].anchor.as_ref().unwrap().row_end, 15);
}

#[test]
fn sales_anchor_col_span() {
    let wb = extract_charts(&fixture()).unwrap();
    assert_eq!(
        wb.sheets[0].charts[0].anchor.as_ref().unwrap().col_span(),
        8
    );
}

#[test]
fn sales_anchor_row_span() {
    let wb = extract_charts(&fixture()).unwrap();
    assert_eq!(
        wb.sheets[0].charts[0].anchor.as_ref().unwrap().row_span(),
        15
    );
}

#[test]
fn sales_anchor_offsets_zero() {
    let wb = extract_charts(&fixture()).unwrap();
    let anch = wb.sheets[0].charts[0].anchor.as_ref().unwrap();
    assert_eq!(anch.col_off, 0);
    assert_eq!(anch.row_off, 0);
    assert_eq!(anch.col_end_off, 0);
    assert_eq!(anch.row_end_off, 0);
}

#[test]
fn expenses_anchor_matches_sales() {
    // Both charts in our fixture are given the same 0,0 → 8,15 anchor
    let wb = extract_charts(&fixture()).unwrap();
    let sales = wb.sheets[0].charts[0].anchor.as_ref().unwrap();
    let expens = wb.sheets[1].charts[0].anchor.as_ref().unwrap();
    assert_eq!(sales.col_start, expens.col_start);
    assert_eq!(sales.row_start, expens.row_start);
    assert_eq!(sales.col_end, expens.col_end);
    assert_eq!(sales.row_end, expens.row_end);
}

// ── Error handling ────────────────────────────────────────────────────────────

#[test]
fn missing_file_returns_error() {
    assert!(extract_charts("no_such_file.xlsx").is_err());
}

#[test]
fn non_zip_file_returns_error() {
    let tmp = std::env::temp_dir().join("not_an_xlsx.xlsx");
    std::fs::write(&tmp, b"not a zip").unwrap();
    assert!(extract_charts(tmp.to_str().unwrap()).is_err());
}

// ═════════════════════════════════════════════════════════════════════════════
// Phase 8 — 3-D chart integration tests
// Fixture: tests/fixtures/test_3d_charts.xlsx
//
//  idx  sheet       XML tag             barDir  rotX  rotY  rAngAx  persp
//   0   Bar3D       c:bar3DChart        col      30    20     1      30
//   1   Line3D      c:line3DChart       —        15    10     1       0
//   2   Area3D      c:area3DChart       —        10     5     0      45
//   3   Pie3D       c:pie3DChart        —        15     0     0      45
//   4   Surface3D   c:surface3DChart    —       -30   180    (none) (none)
//   5   HBar3D      c:bar3DChart        bar      20    15     1     (none)
// ═════════════════════════════════════════════════════════════════════════════

// use sheetforge_charts::model::chart::Chart3DView;

fn fixture_3d() -> String {
    std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/test_3d_charts.xlsx")
        .to_string_lossy()
        .into_owned()
}

// ── Workbook structure ────────────────────────────────────────────────────────

#[test]
fn fixture_3d_has_six_sheets() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    assert_eq!(wb.sheets.len(), 6);
}

#[test]
fn fixture_3d_has_six_charts() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    assert_eq!(wb.chart_count(), 6);
}

// ── ChartType detection ───────────────────────────────────────────────────────

#[test]
fn sheet0_bar3d_chart_type() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    assert_eq!(wb.sheets[0].charts[0].chart_type, ChartType::Bar3D);
}

#[test]
fn sheet1_line3d_chart_type() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    assert_eq!(wb.sheets[1].charts[0].chart_type, ChartType::Line3D);
}

#[test]
fn sheet2_area3d_chart_type() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    assert_eq!(wb.sheets[2].charts[0].chart_type, ChartType::Area3D);
}

#[test]
fn sheet3_pie3d_chart_type() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    assert_eq!(wb.sheets[3].charts[0].chart_type, ChartType::Pie3D);
}

#[test]
fn sheet4_surface3d_chart_type() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    assert_eq!(wb.sheets[4].charts[0].chart_type, ChartType::Surface3D);
}

#[test]
fn sheet5_horizontal_bar3d_chart_type() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    assert_eq!(
        wb.sheets[5].charts[0].chart_type,
        ChartType::HorizontalBar3D
    );
}

// ── is_3d() helper ────────────────────────────────────────────────────────────

#[test]
fn all_3d_charts_report_is_3d_true() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    for sheet in &wb.sheets {
        for chart in &sheet.charts {
            assert!(
                chart.chart_type.is_3d(),
                "Expected is_3d() = true for {:?} (sheet '{}')",
                chart.chart_type,
                sheet.name,
            );
        }
    }
}

// ── view_3d presence ──────────────────────────────────────────────────────────

#[test]
fn all_3d_charts_have_view_3d() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    for sheet in &wb.sheets {
        for chart in &sheet.charts {
            assert!(
                chart.view_3d.is_some(),
                "Expected view_3d to be Some for {:?} (sheet '{}')",
                chart.chart_type,
                sheet.name,
            );
        }
    }
}

// ── rotX ─────────────────────────────────────────────────────────────────────

#[test]
fn bar3d_rotx_is_30() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    let v = wb.sheets[0].charts[0].view_3d.as_ref().unwrap();
    assert_eq!(v.rotation_x, Some(30));
}

#[test]
fn line3d_rotx_is_15() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    let v = wb.sheets[1].charts[0].view_3d.as_ref().unwrap();
    assert_eq!(v.rotation_x, Some(15));
}

#[test]
fn area3d_rotx_is_10() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    let v = wb.sheets[2].charts[0].view_3d.as_ref().unwrap();
    assert_eq!(v.rotation_x, Some(10));
}

#[test]
fn pie3d_rotx_is_15() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    let v = wb.sheets[3].charts[0].view_3d.as_ref().unwrap();
    assert_eq!(v.rotation_x, Some(15));
}

#[test]
fn surface3d_rotx_is_negative_30() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    let v = wb.sheets[4].charts[0].view_3d.as_ref().unwrap();
    assert_eq!(v.rotation_x, Some(-30));
}

#[test]
fn hbar3d_rotx_is_20() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    let v = wb.sheets[5].charts[0].view_3d.as_ref().unwrap();
    assert_eq!(v.rotation_x, Some(20));
}

// ── rotY ─────────────────────────────────────────────────────────────────────

#[test]
fn bar3d_roty_is_20() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    let v = wb.sheets[0].charts[0].view_3d.as_ref().unwrap();
    assert_eq!(v.rotation_y, Some(20));
}

#[test]
fn line3d_roty_is_10() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    let v = wb.sheets[1].charts[0].view_3d.as_ref().unwrap();
    assert_eq!(v.rotation_y, Some(10));
}

#[test]
fn pie3d_roty_is_0() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    let v = wb.sheets[3].charts[0].view_3d.as_ref().unwrap();
    assert_eq!(v.rotation_y, Some(0));
}

#[test]
fn surface3d_roty_is_180() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    let v = wb.sheets[4].charts[0].view_3d.as_ref().unwrap();
    assert_eq!(v.rotation_y, Some(180));
}

// ── right_angle_axes ──────────────────────────────────────────────────────────

#[test]
fn bar3d_right_angle_axes_true() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    let v = wb.sheets[0].charts[0].view_3d.as_ref().unwrap();
    assert_eq!(v.right_angle_axes, Some(true));
}

#[test]
fn area3d_right_angle_axes_false() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    let v = wb.sheets[2].charts[0].view_3d.as_ref().unwrap();
    assert_eq!(v.right_angle_axes, Some(false));
}

#[test]
fn pie3d_right_angle_axes_false() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    let v = wb.sheets[3].charts[0].view_3d.as_ref().unwrap();
    assert_eq!(v.right_angle_axes, Some(false));
}

#[test]
fn surface3d_right_angle_axes_absent() {
    // chart5.xml has no <c:rAngAx> element
    let wb = extract_charts(&fixture_3d()).unwrap();
    let v = wb.sheets[4].charts[0].view_3d.as_ref().unwrap();
    assert_eq!(v.right_angle_axes, None);
}

#[test]
fn hbar3d_right_angle_axes_true() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    let v = wb.sheets[5].charts[0].view_3d.as_ref().unwrap();
    assert_eq!(v.right_angle_axes, Some(true));
}

// ── perspective ───────────────────────────────────────────────────────────────

#[test]
fn bar3d_perspective_is_30() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    let v = wb.sheets[0].charts[0].view_3d.as_ref().unwrap();
    assert_eq!(v.perspective, Some(30));
}

#[test]
fn line3d_perspective_is_0() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    let v = wb.sheets[1].charts[0].view_3d.as_ref().unwrap();
    assert_eq!(v.perspective, Some(0));
}

#[test]
fn area3d_perspective_is_45() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    let v = wb.sheets[2].charts[0].view_3d.as_ref().unwrap();
    assert_eq!(v.perspective, Some(45));
}

#[test]
fn pie3d_perspective_is_45() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    let v = wb.sheets[3].charts[0].view_3d.as_ref().unwrap();
    assert_eq!(v.perspective, Some(45));
}

#[test]
fn surface3d_perspective_absent() {
    // chart5.xml has no <c:perspective> element
    let wb = extract_charts(&fixture_3d()).unwrap();
    let v = wb.sheets[4].charts[0].view_3d.as_ref().unwrap();
    assert_eq!(v.perspective, None);
}

#[test]
fn hbar3d_perspective_absent() {
    // chart6.xml has no <c:perspective> element
    let wb = extract_charts(&fixture_3d()).unwrap();
    let v = wb.sheets[5].charts[0].view_3d.as_ref().unwrap();
    assert_eq!(v.perspective, None);
}

// ── is_empty() on populated view ─────────────────────────────────────────────

#[test]
fn bar3d_view3d_is_not_empty() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    assert!(!wb.sheets[0].charts[0].view_3d.as_ref().unwrap().is_empty());
}

// ── existing 2-D fixture is unaffected ───────────────────────────────────────

#[test]
fn existing_2d_sales_chart_has_no_view3d() {
    let wb = extract_charts(&fixture()).unwrap();
    assert!(wb.sheets[0].charts[0].view_3d.is_none());
}

#[test]
fn existing_2d_expenses_chart_has_no_view3d() {
    let wb = extract_charts(&fixture()).unwrap();
    assert!(wb.sheets[1].charts[0].view_3d.is_none());
}

// ── titles are preserved for 3D charts ───────────────────────────────────────

#[test]
fn bar3d_title_parsed() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    assert_eq!(
        wb.sheets[0].charts[0].title.as_deref(),
        Some("Sales 3D Bar")
    );
}

#[test]
fn surface3d_title_parsed() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    assert_eq!(wb.sheets[4].charts[0].title.as_deref(), Some("Surface 3D"));
}

// ── series caches work inside 3D chart XML ───────────────────────────────────

#[test]
fn bar3d_series_value_cache() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    let ser = &wb.sheets[0].charts[0].series[0];
    let cache = ser.value_cache.as_ref().expect("should have value cache");
    assert_eq!(cache.values, vec![100.0, 200.0, 150.0]);
}

#[test]
fn line3d_series_value_cache() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    let ser = &wb.sheets[1].charts[0].series[0];
    let cache = ser.value_cache.as_ref().expect("should have value cache");
    assert_eq!(cache.values, vec![100.0, 200.0, 150.0]);
}

// ── anchors preserved for 3D charts ──────────────────────────────────────────

#[test]
fn bar3d_has_anchor() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    assert!(wb.sheets[0].charts[0].anchor.is_some());
}

#[test]
fn bar3d_anchor_col_span() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    let a = wb.sheets[0].charts[0].anchor.as_ref().unwrap();
    assert_eq!(a.col_span(), 8);
}

#[test]
fn bar3d_anchor_row_span() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    let a = wb.sheets[0].charts[0].anchor.as_ref().unwrap();
    assert_eq!(a.row_span(), 15);
}

// ── legend preserved ─────────────────────────────────────────────────────────

#[test]
fn bar3d_legend_position_bottom() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    let chart = &wb.sheets[0].charts[0];
    assert_eq!(chart.legend_position, Some(LegendPosition::Bottom));
}

// ═════════════════════════════════════════════════════════════════════════════
// Phase 9 — 3-D geometry surface integration tests
// ═════════════════════════════════════════════════════════════════════════════

fn fixture_surface() -> String {
    std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/test_surface_charts.xlsx")
        .to_string_lossy()
        .into_owned()
}

// ── Workbook structure ────────────────────────────────────────────────────────

#[test]
fn surface_fixture_has_three_sheets() {
    let wb = extract_charts(&fixture_surface()).unwrap();
    assert_eq!(wb.sheets.len(), 3);
}

#[test]
fn surface_fixture_total_chart_count() {
    let wb = extract_charts(&fixture_surface()).unwrap();
    assert_eq!(wb.chart_count(), 3);
}

#[test]
fn surface_sheet_names() {
    let wb = extract_charts(&fixture_surface()).unwrap();
    assert_eq!(wb.sheets[0].name, "AllSurfaces");
    assert_eq!(wb.sheets[1].name, "FloorOnly");
    assert_eq!(wb.sheets[2].name, "NoFillSurfaces");
}

// ── All charts are 3-D bar charts ─────────────────────────────────────────────

#[test]
fn all_surface_charts_are_bar3d() {
    let wb = extract_charts(&fixture_surface()).unwrap();
    for sheet in &wb.sheets {
        assert_eq!(
            sheet.charts[0].chart_type,
            ChartType::Bar3D,
            "sheet {} should be Bar3D",
            sheet.name
        );
    }
}

#[test]
fn all_surface_charts_have_view3d() {
    let wb = extract_charts(&fixture_surface()).unwrap();
    for sheet in &wb.sheets {
        assert!(
            sheet.charts[0].view_3d.is_some(),
            "sheet {} should have view_3d",
            sheet.name
        );
    }
}

// ── AllSurfaces — surface present ────────────────────────────────────────────

#[test]
fn all_surfaces_surface_is_some() {
    let wb = extract_charts(&fixture_surface()).unwrap();
    assert!(wb.sheets[0].charts[0].surface.is_some());
}

// ── AllSurfaces — floor fill ─────────────────────────────────────────────────

#[test]
fn all_surfaces_floor_fill_is_some() {
    let wb = extract_charts(&fixture_surface()).unwrap();
    let surf = wb.sheets[0].charts[0].surface.as_ref().unwrap();
    assert!(surf.floor_fill.is_some());
}

#[test]
fn all_surfaces_floor_solid_srgb() {
    let wb = extract_charts(&fixture_surface()).unwrap();
    let surf = wb.sheets[0].charts[0].surface.as_ref().unwrap();
    match surf.floor_fill.as_ref().unwrap() {
        Fill::Solid(ColorSpec::Srgb(rgb, mods)) => {
            assert_eq!(*rgb, Rgb::from_hex("D9D9D9").unwrap());
            assert!(mods.is_empty(), "no mods expected");
        }
        other => panic!("expected Solid(Srgb(D9D9D9)), got {other:?}"),
    }
}

// ── AllSurfaces — side-wall fill ──────────────────────────────────────────────

#[test]
fn all_surfaces_side_wall_fill_is_some() {
    let wb = extract_charts(&fixture_surface()).unwrap();
    let surf = wb.sheets[0].charts[0].surface.as_ref().unwrap();
    assert!(surf.side_wall_fill.is_some());
}

#[test]
fn all_surfaces_side_wall_solid_srgb() {
    let wb = extract_charts(&fixture_surface()).unwrap();
    let surf = wb.sheets[0].charts[0].surface.as_ref().unwrap();
    match surf.side_wall_fill.as_ref().unwrap() {
        Fill::Solid(ColorSpec::Srgb(rgb, _)) => {
            assert_eq!(*rgb, Rgb::from_hex("4472C4").unwrap());
        }
        other => panic!("expected Solid(Srgb(4472C4)), got {other:?}"),
    }
}

// ── AllSurfaces — back-wall fill (gradient) ───────────────────────────────────

#[test]
fn all_surfaces_back_wall_fill_is_gradient() {
    let wb = extract_charts(&fixture_surface()).unwrap();
    let surf = wb.sheets[0].charts[0].surface.as_ref().unwrap();
    assert!(
        matches!(surf.back_wall_fill.as_ref().unwrap(), Fill::Gradient(_)),
        "back_wall should be a gradient"
    );
}

#[test]
fn all_surfaces_back_wall_gradient_two_stops() {
    let wb = extract_charts(&fixture_surface()).unwrap();
    let surf = wb.sheets[0].charts[0].surface.as_ref().unwrap();
    if let Fill::Gradient(g) = surf.back_wall_fill.as_ref().unwrap() {
        assert_eq!(g.stops.len(), 2);
    }
}

#[test]
fn all_surfaces_back_wall_gradient_stop0() {
    let wb = extract_charts(&fixture_surface()).unwrap();
    let surf = wb.sheets[0].charts[0].surface.as_ref().unwrap();
    if let Fill::Gradient(g) = surf.back_wall_fill.as_ref().unwrap() {
        assert_eq!(g.stops[0].position, 0);
        match &g.stops[0].color {
            ColorSpec::Srgb(rgb, _) => assert_eq!(*rgb, Rgb::from_hex("FF0000").unwrap()),
            other => panic!("expected Srgb, got {other:?}"),
        }
    }
}

#[test]
fn all_surfaces_back_wall_gradient_stop1() {
    let wb = extract_charts(&fixture_surface()).unwrap();
    let surf = wb.sheets[0].charts[0].surface.as_ref().unwrap();
    if let Fill::Gradient(g) = surf.back_wall_fill.as_ref().unwrap() {
        assert_eq!(g.stops[1].position, 100_000);
        match &g.stops[1].color {
            ColorSpec::Srgb(rgb, _) => assert_eq!(*rgb, Rgb::from_hex("FFFFFF").unwrap()),
            other => panic!("expected Srgb, got {other:?}"),
        }
    }
}

#[test]
fn all_surfaces_back_wall_gradient_linear_90deg() {
    let wb = extract_charts(&fixture_surface()).unwrap();
    let surf = wb.sheets[0].charts[0].surface.as_ref().unwrap();
    if let Fill::Gradient(g) = surf.back_wall_fill.as_ref().unwrap() {
        match g.direction.as_ref().unwrap() {
            GradientDirection::Linear { angle_deg, scaled } => {
                // ang=5400000 in XML → 5400000/60000 = 90.0°
                assert!(
                    (angle_deg - 90.0).abs() < 0.01,
                    "angle should be 90°, got {angle_deg}"
                );
                assert!(!scaled);
            }
            other => panic!("expected Linear direction, got {other:?}"),
        }
    }
}

// ── AllSurfaces — isolation checks ────────────────────────────────────────────

#[test]
fn all_surfaces_chart_fill_unaffected() {
    // The AllSurfaces chart has no chart-space <c:spPr> → chart_fill should be None
    let wb = extract_charts(&fixture_surface()).unwrap();
    assert!(
        wb.sheets[0].charts[0].chart_fill.is_none(),
        "surface fills must not bleed into chart_fill"
    );
}

#[test]
fn all_surfaces_plot_area_fill_unaffected() {
    let wb = extract_charts(&fixture_surface()).unwrap();
    assert!(
        wb.sheets[0].charts[0].plot_area.fill.is_none(),
        "surface fills must not bleed into plot_area.fill"
    );
}

#[test]
fn all_surfaces_series_fill_unaffected() {
    let wb = extract_charts(&fixture_surface()).unwrap();
    for s in &wb.sheets[0].charts[0].series {
        assert!(
            s.fill.is_none(),
            "surface fills must not bleed into series fill"
        );
    }
}

// ── FloorOnly ─────────────────────────────────────────────────────────────────

#[test]
fn floor_only_surface_is_some() {
    let wb = extract_charts(&fixture_surface()).unwrap();
    assert!(wb.sheets[1].charts[0].surface.is_some());
}

#[test]
fn floor_only_floor_fill_present() {
    let wb = extract_charts(&fixture_surface()).unwrap();
    let surf = wb.sheets[1].charts[0].surface.as_ref().unwrap();
    assert!(surf.floor_fill.is_some());
}

#[test]
fn floor_only_floor_solid_red() {
    let wb = extract_charts(&fixture_surface()).unwrap();
    let surf = wb.sheets[1].charts[0].surface.as_ref().unwrap();
    match surf.floor_fill.as_ref().unwrap() {
        Fill::Solid(ColorSpec::Srgb(rgb, _)) => {
            assert_eq!(*rgb, Rgb::from_hex("FF0000").unwrap());
        }
        other => panic!("expected Solid(Srgb(FF0000)), got {other:?}"),
    }
}

#[test]
fn floor_only_side_wall_absent() {
    let wb = extract_charts(&fixture_surface()).unwrap();
    let surf = wb.sheets[1].charts[0].surface.as_ref().unwrap();
    assert!(
        surf.side_wall_fill.is_none(),
        "sideWall should be None — element absent"
    );
}

#[test]
fn floor_only_back_wall_absent() {
    let wb = extract_charts(&fixture_surface()).unwrap();
    let surf = wb.sheets[1].charts[0].surface.as_ref().unwrap();
    assert!(
        surf.back_wall_fill.is_none(),
        "backWall should be None — element absent"
    );
}

// ── NoFillSurfaces ────────────────────────────────────────────────────────────

#[test]
fn no_fill_surfaces_surface_is_some() {
    // Even though fills are noFill, the field is Fill::None — so surface is Some
    let wb = extract_charts(&fixture_surface()).unwrap();
    assert!(wb.sheets[2].charts[0].surface.is_some());
}

#[test]
fn no_fill_surfaces_floor_is_no_fill() {
    let wb = extract_charts(&fixture_surface()).unwrap();
    let surf = wb.sheets[2].charts[0].surface.as_ref().unwrap();
    assert_eq!(surf.floor_fill.as_ref().unwrap(), &Fill::None);
}

#[test]
fn no_fill_surfaces_side_wall_is_no_fill() {
    let wb = extract_charts(&fixture_surface()).unwrap();
    let surf = wb.sheets[2].charts[0].surface.as_ref().unwrap();
    assert_eq!(surf.side_wall_fill.as_ref().unwrap(), &Fill::None);
}

#[test]
fn no_fill_surfaces_back_wall_is_no_fill() {
    let wb = extract_charts(&fixture_surface()).unwrap();
    let surf = wb.sheets[2].charts[0].surface.as_ref().unwrap();
    assert_eq!(surf.back_wall_fill.as_ref().unwrap(), &Fill::None);
}

// ── 2-D regression — existing charts have surface = None ─────────────────────

#[test]
fn two_d_sales_chart_no_surface() {
    let wb = extract_charts(&fixture()).unwrap();
    assert!(wb.sheets[0].charts[0].surface.is_none());
}

#[test]
fn two_d_expenses_chart_no_surface() {
    let wb = extract_charts(&fixture()).unwrap();
    assert!(wb.sheets[1].charts[0].surface.is_none());
}

// ── 3-D charts without surface elements have surface = None ──────────────────

#[test]
fn bar3d_fixture_no_surface_elements() {
    // The Phase 8 3D fixture has no <c:floor>/<c:sideWall>/<c:backWall>
    let wb = extract_charts(&fixture_3d()).unwrap();
    for sheet in &wb.sheets {
        assert!(
            sheet.charts[0].surface.is_none(),
            "sheet {} in 3d fixture has no surface elements",
            sheet.name
        );
    }
}

// ── view_3d unaffected by surface parsing ─────────────────────────────────────

#[test]
fn all_surfaces_view3d_still_correct() {
    let wb = extract_charts(&fixture_surface()).unwrap();
    let v = wb.sheets[0].charts[0].view_3d.as_ref().unwrap();
    assert_eq!(v.rotation_x, Some(30));
    assert_eq!(v.rotation_y, Some(20));
    assert_eq!(v.right_angle_axes, Some(true));
    assert_eq!(v.perspective, Some(30));
}

// ── series cache still works alongside surface fills ──────────────────────────

#[test]
fn all_surfaces_series_value_cache_intact() {
    let wb = extract_charts(&fixture_surface()).unwrap();
    let cache = wb.sheets[0].charts[0].series[0]
        .value_cache
        .as_ref()
        .expect("AllSurfaces chart should have value cache");
    assert_eq!(cache.values, vec![100.0, 200.0, 150.0]);
}

// ── anchors present on all surface charts ─────────────────────────────────────

#[test]
fn all_surface_charts_have_anchor() {
    let wb = extract_charts(&fixture_surface()).unwrap();
    for sheet in &wb.sheets {
        assert!(
            sheet.charts[0].anchor.is_some(),
            "sheet {} should have an anchor",
            sheet.name
        );
    }
}

#[test]
fn surface_chart_anchor_col_span() {
    let wb = extract_charts(&fixture_surface()).unwrap();
    let a = wb.sheets[0].charts[0].anchor.as_ref().unwrap();
    assert_eq!(a.col_span(), 8);
}

#[test]
fn surface_chart_anchor_row_span() {
    let wb = extract_charts(&fixture_surface()).unwrap();
    let a = wb.sheets[0].charts[0].anchor.as_ref().unwrap();
    assert_eq!(a.row_span(), 15);
}

// ── Chart3DSurface::is_empty() via fixture data ───────────────────────────────

#[test]
fn all_surfaces_surface_is_not_empty() {
    let wb = extract_charts(&fixture_surface()).unwrap();
    let surf = wb.sheets[0].charts[0].surface.as_ref().unwrap();
    assert!(!surf.is_empty());
}

#[test]
fn floor_only_surface_is_not_empty() {
    let wb = extract_charts(&fixture_surface()).unwrap();
    let surf = wb.sheets[1].charts[0].surface.as_ref().unwrap();
    assert!(!surf.is_empty());
}

// ═════════════════════════════════════════════════════════════════════════════
// Phase 10 — Pivot chart detection
// ═════════════════════════════════════════════════════════════════════════════

fn fixture_pivot() -> String {
    std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/test_pivot_charts.xlsx")
        .to_string_lossy()
        .into_owned()
}

// ── Workbook structure ────────────────────────────────────────────────────────

#[test]
fn pivot_fixture_has_three_sheets() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    assert_eq!(wb.sheets.len(), 3);
}

#[test]
fn pivot_fixture_sheet_names() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    let names: Vec<&str> = wb.sheets.iter().map(|s| s.name.as_str()).collect();
    assert_eq!(names, vec!["PivotChart", "NoPivot", "MultiPivot"]);
}

#[test]
fn pivot_fixture_each_sheet_has_one_chart() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    for sheet in &wb.sheets {
        assert_eq!(
            sheet.charts.len(),
            1,
            "sheet '{}' should have exactly 1 chart",
            sheet.name
        );
    }
}

// ── is_pivot_chart ────────────────────────────────────────────────────────────

#[test]
fn pivot_chart_sheet_is_pivot() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    assert!(
        wb.sheets[0].charts[0].is_pivot_chart,
        "PivotChart sheet: is_pivot_chart should be true"
    );
}

#[test]
fn no_pivot_sheet_is_not_pivot() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    assert!(
        !wb.sheets[1].charts[0].is_pivot_chart,
        "NoPivot sheet: is_pivot_chart should be false"
    );
}

#[test]
fn multi_pivot_sheet_is_pivot() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    assert!(
        wb.sheets[2].charts[0].is_pivot_chart,
        "MultiPivot sheet: is_pivot_chart should be true"
    );
}

// ── pivot_table_name ──────────────────────────────────────────────────────────

#[test]
fn pivot_chart_name_correct() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    assert_eq!(
        wb.sheets[0].charts[0].pivot_table_name.as_deref(),
        Some("Sheet1!PivotTable1")
    );
}

#[test]
fn no_pivot_name_is_none() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    assert!(wb.sheets[1].charts[0].pivot_table_name.is_none());
}

#[test]
fn multi_pivot_name_correct() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    assert_eq!(
        wb.sheets[2].charts[0].pivot_table_name.as_deref(),
        Some("Sales!RevenueByRegion")
    );
}

// ── other fields unaffected by pivot detection ────────────────────────────────

#[test]
fn pivot_chart_type_is_bar() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    assert_eq!(wb.sheets[0].charts[0].chart_type, ChartType::Bar);
}

#[test]
fn no_pivot_chart_type_is_bar() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    assert_eq!(wb.sheets[1].charts[0].chart_type, ChartType::Bar);
}

#[test]
fn multi_pivot_chart_type_is_bar() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    assert_eq!(wb.sheets[2].charts[0].chart_type, ChartType::Bar);
}

#[test]
fn pivot_chart_has_three_series() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    assert_eq!(wb.sheets[0].charts[0].series.len(), 3);
}

#[test]
fn no_pivot_chart_has_three_series() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    assert_eq!(wb.sheets[1].charts[0].series.len(), 3);
}

#[test]
fn pivot_chart_series_have_caches() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    let chart = &wb.sheets[0].charts[0];
    // Every series should have a value cache with 4 data points
    for s in &chart.series {
        let cache = s
            .value_cache
            .as_ref()
            .unwrap_or_else(|| panic!("series {} missing value cache", s.index));
        assert_eq!(
            cache.values.len(),
            4,
            "series {} cache should have 4 values",
            s.index
        );
    }
}

#[test]
fn pivot_chart_anchor_present() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    assert!(
        wb.sheets[0].charts[0].anchor.is_some(),
        "pivot chart should have an anchor from twoCellAnchor"
    );
}

#[test]
fn no_pivot_chart_anchor_present() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    assert!(wb.sheets[1].charts[0].anchor.is_some());
}

#[test]
fn pivot_chart_surface_is_none() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    assert!(
        wb.sheets[0].charts[0].surface.is_none(),
        "2-D pivot bar chart should have no surface"
    );
}

#[test]
fn pivot_chart_view3d_is_none() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    assert!(
        wb.sheets[0].charts[0].view_3d.is_none(),
        "2-D pivot bar chart should have no view_3d"
    );
}

// ── regression: existing fixtures unaffected ─────────────────────────────────

#[test]
fn sales_chart_not_pivot() {
    // The main 2-D fixture has no pivotSource elements
    let wb = extract_charts(&fixture()).unwrap();
    assert!(!wb.sheets[0].charts[0].is_pivot_chart);
}

#[test]
fn sales_chart_pivot_name_none() {
    let wb = extract_charts(&fixture()).unwrap();
    assert!(wb.sheets[0].charts[0].pivot_table_name.is_none());
}

#[test]
fn bar3d_fixture_not_pivot() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    for sheet in &wb.sheets {
        for chart in &sheet.charts {
            assert!(
                !chart.is_pivot_chart,
                "3-D fixture chart {} should not be pivot",
                chart.chart_path
            );
        }
    }
}

#[test]
fn surface_fixture_not_pivot() {
    let wb = extract_charts(&fixture_surface()).unwrap();
    for sheet in &wb.sheets {
        for chart in &sheet.charts {
            assert!(
                !chart.is_pivot_chart,
                "surface fixture chart {} should not be pivot",
                chart.chart_path
            );
        }
    }
}
