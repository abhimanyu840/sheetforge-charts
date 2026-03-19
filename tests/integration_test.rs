//! Integration tests — Phases 1–11.
//!
//! Fixtures:
//!   tests/fixtures/test_charts.xlsx      (Phases 1–7, 2-D charts)
//!   tests/fixtures/test_3d_charts.xlsx   (Phase 8, 3-D charts)
//!   tests/fixtures/test_surface_charts.xlsx  (Phase 9, surface geometry fills)
//!   tests/fixtures/test_pivot_charts.xlsx  (Phases 10–11, pivot detection + metadata)
//!     Sheet "PivotChart"  → chart1.xml  bar chart + <pivotSource> "Sheet1!PivotTable1"
//!                           chart1.xml.rels → pivotTable1.xml (PivotTable1, 4 fields)
//!                           → pivotCacheDefinition1.xml (SourceData A1:D101)
//!     Sheet "NoPivot"     → chart2.xml  bar chart, no <pivotSource>
//!     Sheet "MultiPivot"  → chart3.xml  bar chart + <pivotSource> "Sales!RevenueByRegion"
//!                           chart3.xml.rels → pivotTable2.xml (RevenueByRegion, 3 fields)
//!                           → pivotCacheDefinition2.xml (SalesData B1:D51)

use sheetforge_charts::{
    extract_charts,
    model::{
        axis::{AxisPosition, AxisType},
        chart::{Chart3DSurface, Chart3DView, ChartAnchor, ChartType, Grouping, LegendPosition},
        color::{ColorSpec, Fill, GradientDirection, Rgb, ThemeColorSlot},
        pivot::{PivotField, PivotTableMeta},
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

use sheetforge_charts::model::chart::Chart3DView;

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
// ═════════════════════════════════════════════════════════════════════════════
// Phase 11 — Pivot Table Metadata
// ═════════════════════════════════════════════════════════════════════════════

// ── pivot_meta presence ───────────────────────────────────────────────────────

#[test]
fn pivot_chart_has_meta() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    assert!(
        wb.sheets[0].charts[0].pivot_meta.is_some(),
        "PivotChart sheet: pivot_meta should be Some"
    );
}

#[test]
fn no_pivot_chart_meta_is_none() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    assert!(
        wb.sheets[1].charts[0].pivot_meta.is_none(),
        "NoPivot sheet: pivot_meta should be None"
    );
}

#[test]
fn multi_pivot_chart_has_meta() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    assert!(
        wb.sheets[2].charts[0].pivot_meta.is_some(),
        "MultiPivot sheet: pivot_meta should be Some"
    );
}

// ── pivot_table_name (from pivotTableDefinition, not from pivotSource) ────────

#[test]
fn pivot_meta_table_name_first_chart() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    let meta = wb.sheets[0].charts[0].pivot_meta.as_ref().unwrap();
    assert_eq!(meta.pivot_table_name, "PivotTable1");
}

#[test]
fn pivot_meta_table_name_third_chart() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    let meta = wb.sheets[2].charts[0].pivot_meta.as_ref().unwrap();
    assert_eq!(meta.pivot_table_name, "RevenueByRegion");
}

// ── source_sheet ──────────────────────────────────────────────────────────────

#[test]
fn pivot_meta_source_sheet_first_chart() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    let meta = wb.sheets[0].charts[0].pivot_meta.as_ref().unwrap();
    assert_eq!(meta.source_sheet.as_deref(), Some("SourceData"));
}

#[test]
fn pivot_meta_source_sheet_third_chart() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    let meta = wb.sheets[2].charts[0].pivot_meta.as_ref().unwrap();
    assert_eq!(meta.source_sheet.as_deref(), Some("SalesData"));
}

// ── source_range ──────────────────────────────────────────────────────────────

#[test]
fn pivot_meta_source_range_first_chart() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    let meta = wb.sheets[0].charts[0].pivot_meta.as_ref().unwrap();
    assert_eq!(meta.source_range.as_deref(), Some("A1:D101"));
}

#[test]
fn pivot_meta_source_range_third_chart() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    let meta = wb.sheets[2].charts[0].pivot_meta.as_ref().unwrap();
    assert_eq!(meta.source_range.as_deref(), Some("B1:D51"));
}

// ── pivot_fields count ────────────────────────────────────────────────────────

#[test]
fn pivot_meta_field_count_first_chart() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    let meta = wb.sheets[0].charts[0].pivot_meta.as_ref().unwrap();
    assert_eq!(
        meta.pivot_fields.len(),
        4,
        "PivotTable1 should have 4 fields"
    );
}

#[test]
fn pivot_meta_field_count_third_chart() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    let meta = wb.sheets[2].charts[0].pivot_meta.as_ref().unwrap();
    assert_eq!(
        meta.pivot_fields.len(),
        3,
        "RevenueByRegion should have 3 fields"
    );
}

// ── pivot_fields names in order ───────────────────────────────────────────────

#[test]
fn pivot_meta_fields_first_chart_in_order() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    let meta = wb.sheets[0].charts[0].pivot_meta.as_ref().unwrap();
    let names: Vec<&str> = meta.pivot_fields.iter().map(|f| f.name.as_str()).collect();
    assert_eq!(names, vec!["Region", "Product", "Category", "Sales"]);
}

#[test]
fn pivot_meta_fields_third_chart_in_order() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    let meta = wb.sheets[2].charts[0].pivot_meta.as_ref().unwrap();
    let names: Vec<&str> = meta.pivot_fields.iter().map(|f| f.name.as_str()).collect();
    assert_eq!(names, vec!["Region", "Quarter", "Revenue"]);
}

#[test]
fn pivot_meta_first_field_name_first_chart() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    let meta = wb.sheets[0].charts[0].pivot_meta.as_ref().unwrap();
    assert_eq!(meta.pivot_fields[0].name, "Region");
}

#[test]
fn pivot_meta_last_field_name_first_chart() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    let meta = wb.sheets[0].charts[0].pivot_meta.as_ref().unwrap();
    assert_eq!(meta.pivot_fields[3].name, "Sales");
}

// ── other chart fields still correct alongside pivot_meta ─────────────────────

#[test]
fn pivot_meta_chart_type_still_bar() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    assert_eq!(wb.sheets[0].charts[0].chart_type, ChartType::Bar);
}

#[test]
fn pivot_meta_is_pivot_chart_still_true() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    assert!(wb.sheets[0].charts[0].is_pivot_chart);
}

#[test]
fn pivot_meta_pivot_table_name_field_matches_source_name() {
    // pivot_table_name (from <pivotSource>) is "Sheet1!PivotTable1"
    // pivot_meta.pivot_table_name (from pivotTableDefinition) is "PivotTable1"
    // These are two different representations — both should be present.
    let wb = extract_charts(&fixture_pivot()).unwrap();
    let chart = &wb.sheets[0].charts[0];
    assert_eq!(
        chart.pivot_table_name.as_deref(),
        Some("Sheet1!PivotTable1")
    );
    assert_eq!(
        chart.pivot_meta.as_ref().unwrap().pivot_table_name,
        "PivotTable1"
    );
}

// ── regression: non-pivot fixture charts have no pivot_meta ──────────────────

#[test]
fn sales_chart_pivot_meta_is_none() {
    let wb = extract_charts(&fixture()).unwrap();
    assert!(wb.sheets[0].charts[0].pivot_meta.is_none());
}

#[test]
fn bar3d_charts_pivot_meta_is_none() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    for sheet in &wb.sheets {
        for chart in &sheet.charts {
            assert!(
                chart.pivot_meta.is_none(),
                "3-D chart {} should have no pivot_meta",
                chart.chart_path
            );
        }
    }
}

#[test]
fn surface_charts_pivot_meta_is_none() {
    let wb = extract_charts(&fixture_surface()).unwrap();
    for sheet in &wb.sheets {
        for chart in &sheet.charts {
            assert!(
                chart.pivot_meta.is_none(),
                "surface chart {} should have no pivot_meta",
                chart.chart_path
            );
        }
    }
}

// ═════════════════════════════════════════════════════════════════════════════
// Phase 12 — Pivot Chart Data Extraction (cached records aggregation)
// ═════════════════════════════════════════════════════════════════════════════
//
// Fixture: test_pivot_charts.xlsx
//
// Pivot 1 (PivotTable1) — single series, no col field:
//   Fields: Region(row=0), Product(idx=1), Category(idx=2), Sales(data=3)
//   sharedItems: Region=[North,South,East], Product=[Widget,Gadget], Category=[Electronics,Hardware]
//   Records (6 rows) → Sum of Sales by Region:
//     North = 1500+2300 = 3800
//     South = 800+1200  = 2000
//     East  = 3100+900  = 4000
//
// Pivot 2 (RevenueByRegion) — multi-series, col field=Quarter:
//   Fields: Region(row=0), Quarter(col=1), Revenue(data=2)
//   sharedItems: Region=[East,West], Quarter=[Q1,Q2]
//   Records (4 rows) → Sum of Revenue:
//     Q1: East=5000, West=3000
//     Q2: East=6000, West=4500

// ── pivot_series presence ─────────────────────────────────────────────────────

#[test]
fn pivot1_has_cached_series() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    let meta = wb.sheets[0].charts[0].pivot_meta.as_ref().unwrap();
    assert!(
        !meta.pivot_series.is_empty(),
        "PivotTable1 should have cached series from records"
    );
}

#[test]
fn pivot2_has_cached_series() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    let meta = wb.sheets[2].charts[0].pivot_meta.as_ref().unwrap();
    assert!(
        !meta.pivot_series.is_empty(),
        "RevenueByRegion should have cached series from records"
    );
}

#[test]
fn no_pivot_chart_has_no_cached_series() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    // NoPivot sheet has no pivot_meta at all
    assert!(wb.sheets[1].charts[0].pivot_meta.is_none());
}

// ── Pivot 1: single-series aggregation ───────────────────────────────────────

#[test]
fn pivot1_single_series_count() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    let meta = wb.sheets[0].charts[0].pivot_meta.as_ref().unwrap();
    assert_eq!(
        meta.pivot_series.len(),
        1,
        "PivotTable1 has one data field and no col field → 1 series"
    );
}

#[test]
fn pivot1_series_name_is_data_field_name() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    let meta = wb.sheets[0].charts[0].pivot_meta.as_ref().unwrap();
    assert_eq!(meta.pivot_series[0].name.as_deref(), Some("Sum of Sales"));
}

#[test]
fn pivot1_categories_in_order() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    let meta = wb.sheets[0].charts[0].pivot_meta.as_ref().unwrap();
    let cats = meta.pivot_series[0]
        .category_values
        .as_ref()
        .unwrap()
        .values
        .as_slice();
    assert_eq!(cats, &["North", "South", "East"]);
}

#[test]
fn pivot1_values_summed_correctly() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    let meta = wb.sheets[0].charts[0].pivot_meta.as_ref().unwrap();
    let vals = &meta.pivot_series[0].value_cache.as_ref().unwrap().values;
    // North=1500+2300=3800, South=800+1200=2000, East=3100+900=4000
    assert_eq!(vals, &[3800.0, 2000.0, 4000.0]);
}

#[test]
fn pivot1_value_cache_state_complete() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    let meta = wb.sheets[0].charts[0].pivot_meta.as_ref().unwrap();
    use sheetforge_charts::model::series::CacheState;
    assert_eq!(meta.pivot_series[0].value_cache_state, CacheState::Complete);
}

#[test]
fn pivot1_category_cache_state_complete() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    let meta = wb.sheets[0].charts[0].pivot_meta.as_ref().unwrap();
    use sheetforge_charts::model::series::CacheState;
    assert_eq!(
        meta.pivot_series[0].category_cache_state,
        CacheState::Complete
    );
}

#[test]
fn pivot1_category_count_matches_value_count() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    let meta = wb.sheets[0].charts[0].pivot_meta.as_ref().unwrap();
    let s = &meta.pivot_series[0];
    let n_cats = s.category_values.as_ref().unwrap().values.len();
    let n_vals = s.value_cache.as_ref().unwrap().values.len();
    assert_eq!(n_cats, n_vals);
}

// ── Pivot 2: multi-series (col field) aggregation ────────────────────────────

#[test]
fn pivot2_two_series() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    let meta = wb.sheets[2].charts[0].pivot_meta.as_ref().unwrap();
    assert_eq!(
        meta.pivot_series.len(),
        2,
        "RevenueByRegion has 1 data field × 2 col values = 2 series"
    );
}

#[test]
fn pivot2_series_names_include_quarter() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    let meta = wb.sheets[2].charts[0].pivot_meta.as_ref().unwrap();
    let names: Vec<&str> = meta
        .pivot_series
        .iter()
        .map(|s| s.name.as_deref().unwrap_or(""))
        .collect();
    // Series names include the data field name and the col key
    assert!(
        names[0].contains("Q1"),
        "first series should reference Q1, got {:?}",
        names[0]
    );
    assert!(
        names[1].contains("Q2"),
        "second series should reference Q2, got {:?}",
        names[1]
    );
}

#[test]
fn pivot2_categories_are_regions() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    let meta = wb.sheets[2].charts[0].pivot_meta.as_ref().unwrap();
    let cats = meta.pivot_series[0]
        .category_values
        .as_ref()
        .unwrap()
        .values
        .as_slice();
    assert_eq!(cats, &["East", "West"]);
}

#[test]
fn pivot2_categories_same_across_series() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    let meta = wb.sheets[2].charts[0].pivot_meta.as_ref().unwrap();
    let cats0 = &meta.pivot_series[0]
        .category_values
        .as_ref()
        .unwrap()
        .values;
    let cats1 = &meta.pivot_series[1]
        .category_values
        .as_ref()
        .unwrap()
        .values;
    assert_eq!(cats0, cats1, "all series must share the same category axis");
}

#[test]
fn pivot2_q1_values_correct() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    let meta = wb.sheets[2].charts[0].pivot_meta.as_ref().unwrap();
    // Q1 series: East=5000, West=3000
    let q1 = meta
        .pivot_series
        .iter()
        .find(|s| s.name.as_deref().unwrap_or("").contains("Q1"))
        .expect("Q1 series must exist");
    let vals = &q1.value_cache.as_ref().unwrap().values;
    assert_eq!(vals, &[5000.0, 3000.0]);
}

#[test]
fn pivot2_q2_values_correct() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    let meta = wb.sheets[2].charts[0].pivot_meta.as_ref().unwrap();
    // Q2 series: East=6000, West=4500
    let q2 = meta
        .pivot_series
        .iter()
        .find(|s| s.name.as_deref().unwrap_or("").contains("Q2"))
        .expect("Q2 series must exist");
    let vals = &q2.value_cache.as_ref().unwrap().values;
    assert_eq!(vals, &[6000.0, 4500.0]);
}

// ── other chart fields unaffected ─────────────────────────────────────────────

#[test]
fn pivot1_chart_type_still_bar_with_series() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    assert_eq!(wb.sheets[0].charts[0].chart_type, ChartType::Bar);
}

#[test]
fn pivot1_is_pivot_chart_still_true_with_series() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    assert!(wb.sheets[0].charts[0].is_pivot_chart);
}

// ── regression: non-pivot charts unaffected ───────────────────────────────────

#[test]
fn sales_chart_pivot_series_absent() {
    let wb = extract_charts(&fixture()).unwrap();
    // No pivot_meta at all on regular charts
    assert!(wb.sheets[0].charts[0].pivot_meta.is_none());
}

#[test]
fn no_pivot_sheet_pivot_series_absent() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    // NoPivot sheet: pivot_meta is None → no series
    assert!(wb.sheets[1].charts[0].pivot_meta.is_none());
}

// ═════════════════════════════════════════════════════════════════════════════
// Phase 13 — Combo Chart Support
// ═════════════════════════════════════════════════════════════════════════════
//
// Fixture: test_combo_charts.xlsx
//   Sheet "BarLine"   → chart1.xml  barChart (2 series) + lineChart (1 series)
//   Sheet "BarArea"   → chart2.xml  barChart (1 series) + areaChart (1 series)
//   Sheet "SingleBar" → chart3.xml  barChart (2 series, regression)

use sheetforge_charts::model::chart::ChartLayer;

fn fixture_combo() -> String {
    std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/test_combo_charts.xlsx")
        .to_string_lossy()
        .into_owned()
}

// ── Workbook structure ────────────────────────────────────────────────────────

#[test]
fn combo_fixture_has_three_sheets() {
    let wb = extract_charts(&fixture_combo()).unwrap();
    assert_eq!(wb.sheets.len(), 3);
}

#[test]
fn combo_fixture_sheet_names() {
    let wb = extract_charts(&fixture_combo()).unwrap();
    let names: Vec<&str> = wb.sheets.iter().map(|s| s.name.as_str()).collect();
    assert_eq!(names, vec!["BarLine", "BarArea", "SingleBar"]);
}

// ── chart_type == Combo ───────────────────────────────────────────────────────

#[test]
fn bar_line_chart_type_is_combo() {
    let wb = extract_charts(&fixture_combo()).unwrap();
    assert_eq!(wb.sheets[0].charts[0].chart_type, ChartType::Combo);
}

#[test]
fn bar_area_chart_type_is_combo() {
    let wb = extract_charts(&fixture_combo()).unwrap();
    assert_eq!(wb.sheets[1].charts[0].chart_type, ChartType::Combo);
}

#[test]
fn single_bar_not_combo() {
    let wb = extract_charts(&fixture_combo()).unwrap();
    assert_eq!(wb.sheets[2].charts[0].chart_type, ChartType::Bar);
}

// ── layers count ─────────────────────────────────────────────────────────────

#[test]
fn bar_line_has_two_layers() {
    let wb = extract_charts(&fixture_combo()).unwrap();
    assert_eq!(wb.sheets[0].charts[0].layers.len(), 2);
}

#[test]
fn bar_area_has_two_layers() {
    let wb = extract_charts(&fixture_combo()).unwrap();
    assert_eq!(wb.sheets[1].charts[0].layers.len(), 2);
}

#[test]
fn single_bar_has_one_layer() {
    let wb = extract_charts(&fixture_combo()).unwrap();
    assert_eq!(wb.sheets[2].charts[0].layers.len(), 1);
}

// ── layer types in order ──────────────────────────────────────────────────────

#[test]
fn bar_line_layer0_is_bar() {
    let wb = extract_charts(&fixture_combo()).unwrap();
    assert_eq!(wb.sheets[0].charts[0].layers[0].chart_type, ChartType::Bar);
}

#[test]
fn bar_line_layer1_is_line() {
    let wb = extract_charts(&fixture_combo()).unwrap();
    assert_eq!(wb.sheets[0].charts[0].layers[1].chart_type, ChartType::Line);
}

#[test]
fn bar_area_layer0_is_bar() {
    let wb = extract_charts(&fixture_combo()).unwrap();
    assert_eq!(wb.sheets[1].charts[0].layers[0].chart_type, ChartType::Bar);
}

#[test]
fn bar_area_layer1_is_area() {
    let wb = extract_charts(&fixture_combo()).unwrap();
    assert_eq!(wb.sheets[1].charts[0].layers[1].chart_type, ChartType::Area);
}

#[test]
fn single_bar_layer0_is_bar() {
    let wb = extract_charts(&fixture_combo()).unwrap();
    assert_eq!(wb.sheets[2].charts[0].layers[0].chart_type, ChartType::Bar);
}

// ── per-layer series counts ───────────────────────────────────────────────────

#[test]
fn bar_line_bar_layer_has_two_series() {
    let wb = extract_charts(&fixture_combo()).unwrap();
    assert_eq!(wb.sheets[0].charts[0].layers[0].series.len(), 2);
}

#[test]
fn bar_line_line_layer_has_one_series() {
    let wb = extract_charts(&fixture_combo()).unwrap();
    assert_eq!(wb.sheets[1].charts[0].layers[1].series.len(), 1);
}

#[test]
fn single_bar_layer_series_count() {
    let wb = extract_charts(&fixture_combo()).unwrap();
    assert_eq!(wb.sheets[2].charts[0].layers[0].series.len(), 2);
}

// ── flat series mirrors all layers ───────────────────────────────────────────

#[test]
fn bar_line_flat_series_count() {
    let wb = extract_charts(&fixture_combo()).unwrap();
    // bar layer: 2 + line layer: 1 = 3 total
    assert_eq!(wb.sheets[0].charts[0].series.len(), 3);
}

#[test]
fn bar_area_flat_series_count() {
    let wb = extract_charts(&fixture_combo()).unwrap();
    assert_eq!(wb.sheets[1].charts[0].series.len(), 2);
}

#[test]
fn plot_area_series_mirrors_flat() {
    let wb = extract_charts(&fixture_combo()).unwrap();
    let chart = &wb.sheets[0].charts[0];
    assert_eq!(chart.plot_area.series.len(), chart.series.len());
}

// ── layer grouping ────────────────────────────────────────────────────────────

#[test]
fn bar_line_bar_layer_grouping_clustered() {
    let wb = extract_charts(&fixture_combo()).unwrap();
    assert_eq!(
        wb.sheets[0].charts[0].layers[0].grouping,
        Some(Grouping::Clustered)
    );
}

// ── regression: non-combo charts unaffected ──────────────────────────────────

#[test]
fn sales_chart_layers_populated() {
    // The original 2-D fixture charts must now also have layers
    let wb = extract_charts(&fixture()).unwrap();
    assert!(
        !wb.sheets[0].charts[0].layers.is_empty(),
        "even single-type charts must have at least one layer"
    );
}

#[test]
fn sales_chart_one_layer() {
    let wb = extract_charts(&fixture()).unwrap();
    assert_eq!(wb.sheets[0].charts[0].layers.len(), 1);
}

#[test]
fn sales_chart_layer_type_matches_chart_type() {
    let wb = extract_charts(&fixture()).unwrap();
    let chart = &wb.sheets[0].charts[0];
    assert_eq!(chart.layers[0].chart_type, chart.chart_type);
}

#[test]
fn bar3d_chart_one_layer() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    for sheet in &wb.sheets {
        for chart in &sheet.charts {
            assert_eq!(
                chart.layers.len(),
                1,
                "3-D chart {} should have exactly 1 layer",
                chart.chart_path
            );
        }
    }
}

#[test]
fn pivot_chart_one_layer() {
    let wb = extract_charts(&fixture_pivot()).unwrap();
    // All pivot fixture charts are single-type bar charts
    for sheet in &wb.sheets {
        for chart in &sheet.charts {
            assert_eq!(
                chart.layers.len(),
                1,
                "pivot chart {} should have 1 layer",
                chart.chart_path
            );
        }
    }
}

// ═════════════════════════════════════════════════════════════════════════════
// Phase 14 — Secondary Axis Support
// ═════════════════════════════════════════════════════════════════════════════
//
// Fixture: test_secondary_axis.xlsx
//   Sheet "SecondaryAxis" → chart1.xml  barChart (2 ser, axId=2/left primary)
//                                     + lineChart (1 ser, axId=3/right secondary)
//   Sheet "PrimaryOnly"  → chart2.xml  barChart (2 ser, primary only)
//   Sheet "TwinValue"    → chart3.xml  barChart (2 ser, primary only)

fn fixture_secondary() -> String {
    std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/fixtures/test_secondary_axis.xlsx")
        .to_string_lossy()
        .into_owned()
}

// ── Workbook structure ────────────────────────────────────────────────────────

#[test]
fn secondary_fixture_has_three_sheets() {
    let wb = extract_charts(&fixture_secondary()).unwrap();
    assert_eq!(wb.sheets.len(), 3);
}

#[test]
fn secondary_fixture_sheet_names() {
    let wb = extract_charts(&fixture_secondary()).unwrap();
    let names: Vec<&str> = wb.sheets.iter().map(|s| s.name.as_str()).collect();
    assert_eq!(names, vec!["SecondaryAxis", "PrimaryOnly", "TwinValue"]);
}

// ── chart_type == Combo ───────────────────────────────────────────────────────

#[test]
fn secondary_chart_is_combo() {
    let wb = extract_charts(&fixture_secondary()).unwrap();
    assert_eq!(wb.sheets[0].charts[0].chart_type, ChartType::Combo);
}

#[test]
fn primary_only_chart_is_bar() {
    let wb = extract_charts(&fixture_secondary()).unwrap();
    assert_eq!(wb.sheets[1].charts[0].chart_type, ChartType::Bar);
}

// ── axis counts ───────────────────────────────────────────────────────────────

#[test]
fn secondary_chart_has_three_axes() {
    let wb = extract_charts(&fixture_secondary()).unwrap();
    // catAx(id=1) + valAx(id=2, left) + valAx(id=3, right)
    assert_eq!(wb.sheets[0].charts[0].axes.len(), 3);
}

#[test]
fn primary_only_has_two_axes() {
    let wb = extract_charts(&fixture_secondary()).unwrap();
    assert_eq!(wb.sheets[1].charts[0].axes.len(), 2);
}

// ── secondary axis position ───────────────────────────────────────────────────

#[test]
fn secondary_val_axis_position_right() {
    use sheetforge_charts::model::axis::AxisPosition;
    let wb = extract_charts(&fixture_secondary()).unwrap();
    let axes = &wb.sheets[0].charts[0].axes;
    let sec_ax = axes
        .iter()
        .find(|ax| ax.id == 3)
        .expect("axis id=3 must exist");
    assert_eq!(sec_ax.position, Some(AxisPosition::Right));
}

#[test]
fn primary_val_axis_position_left() {
    use sheetforge_charts::model::axis::AxisPosition;
    let wb = extract_charts(&fixture_secondary()).unwrap();
    let axes = &wb.sheets[0].charts[0].axes;
    let pri_ax = axes
        .iter()
        .find(|ax| ax.id == 2)
        .expect("axis id=2 must exist");
    assert_eq!(pri_ax.position, Some(AxisPosition::Left));
}

// ── series flat list ──────────────────────────────────────────────────────────

#[test]
fn secondary_chart_has_three_series() {
    let wb = extract_charts(&fixture_secondary()).unwrap();
    assert_eq!(wb.sheets[0].charts[0].series.len(), 3);
}

// ── axis_id on series ─────────────────────────────────────────────────────────

#[test]
fn bar_series0_axis_id_primary() {
    let wb = extract_charts(&fixture_secondary()).unwrap();
    let s = &wb.sheets[0].charts[0].series[0];
    assert_eq!(
        s.axis_id,
        Some(2),
        "bar series 0 must reference primary val axis id=2"
    );
}

#[test]
fn bar_series1_axis_id_primary() {
    let wb = extract_charts(&fixture_secondary()).unwrap();
    let s = &wb.sheets[0].charts[0].series[1];
    assert_eq!(s.axis_id, Some(2));
}

#[test]
fn line_series_axis_id_secondary() {
    let wb = extract_charts(&fixture_secondary()).unwrap();
    let s = &wb.sheets[0].charts[0].series[2];
    assert_eq!(
        s.axis_id,
        Some(3),
        "line series must reference secondary val axis id=3"
    );
}

// ── is_secondary_axis on series ───────────────────────────────────────────────

#[test]
fn bar_series0_not_secondary() {
    let wb = extract_charts(&fixture_secondary()).unwrap();
    assert!(!wb.sheets[0].charts[0].series[0].is_secondary_axis);
}

#[test]
fn bar_series1_not_secondary() {
    let wb = extract_charts(&fixture_secondary()).unwrap();
    assert!(!wb.sheets[0].charts[0].series[1].is_secondary_axis);
}

#[test]
fn line_series_is_secondary() {
    let wb = extract_charts(&fixture_secondary()).unwrap();
    assert!(
        wb.sheets[0].charts[0].series[2].is_secondary_axis,
        "line series on right-position axis must be secondary"
    );
}

// ── is_on_secondary_axis() convenience method ─────────────────────────────────

#[test]
fn is_on_secondary_axis_method_true() {
    let wb = extract_charts(&fixture_secondary()).unwrap();
    assert!(wb.sheets[0].charts[0].series[2].is_on_secondary_axis());
}

#[test]
fn is_on_secondary_axis_method_false() {
    let wb = extract_charts(&fixture_secondary()).unwrap();
    assert!(!wb.sheets[0].charts[0].series[0].is_on_secondary_axis());
}

// ── layer axis_ids ────────────────────────────────────────────────────────────

#[test]
fn bar_layer_axis_ids() {
    let wb = extract_charts(&fixture_secondary()).unwrap();
    let ids = &wb.sheets[0].charts[0].layers[0].axis_ids;
    assert!(ids.contains(&1), "barChart must ref catAx id=1");
    assert!(ids.contains(&2), "barChart must ref primary valAx id=2");
}

#[test]
fn line_layer_axis_ids() {
    let wb = extract_charts(&fixture_secondary()).unwrap();
    let ids = &wb.sheets[0].charts[0].layers[1].axis_ids;
    assert!(ids.contains(&1), "lineChart must ref catAx id=1");
    assert!(ids.contains(&3), "lineChart must ref secondary valAx id=3");
}

// ── layer series mirrors axis fields ─────────────────────────────────────────

#[test]
fn bar_layer_series_match_flat() {
    let wb = extract_charts(&fixture_secondary()).unwrap();
    let chart = &wb.sheets[0].charts[0];
    for ls in &chart.layers[0].series {
        let fs = chart.series.iter().find(|s| s.index == ls.index).unwrap();
        assert_eq!(ls.axis_id, fs.axis_id);
        assert_eq!(ls.is_secondary_axis, fs.is_secondary_axis);
    }
}

#[test]
fn line_layer_series_is_secondary() {
    let wb = extract_charts(&fixture_secondary()).unwrap();
    assert!(wb.sheets[0].charts[0].layers[1].series[0].is_secondary_axis);
}

// ── regression: primary-only charts unaffected ───────────────────────────────

#[test]
fn primary_only_series_not_secondary() {
    let wb = extract_charts(&fixture_secondary()).unwrap();
    for s in &wb.sheets[1].charts[0].series {
        assert!(
            !s.is_secondary_axis,
            "primary-only chart must have no secondary series"
        );
    }
}

#[test]
fn original_sales_chart_series_not_secondary() {
    let wb = extract_charts(&fixture()).unwrap();
    for s in &wb.sheets[0].charts[0].series {
        assert!(!s.is_secondary_axis);
    }
}

#[test]
fn combo_fixture_primary_series_not_secondary() {
    let wb = extract_charts(&fixture_combo()).unwrap();
    // combo fixture bar layer: primary (left axis)
    for s in &wb.sheets[0].charts[0].layers[0].series {
        assert!(
            !s.is_secondary_axis,
            "bar layer in combo fixture must be primary"
        );
    }
}

// ═════════════════════════════════════════════════════════════════════════════
// Phase 15 — Chart Placement (ChartPosition)
// ═════════════════════════════════════════════════════════════════════════════

use sheetforge_charts::model::chart::ChartPosition;

// ── ChartPosition present ─────────────────────────────────────────────────────

#[test]
fn sales_chart_has_position() {
    let wb = extract_charts(&fixture()).unwrap();
    assert!(
        wb.sheets[0].charts[0].position.is_some(),
        "Sales chart must have a ChartPosition"
    );
}

#[test]
fn expenses_chart_has_position() {
    let wb = extract_charts(&fixture()).unwrap();
    assert!(wb.sheets[1].charts[0].position.is_some());
}

// ── sheet name in position ────────────────────────────────────────────────────

#[test]
fn sales_chart_position_sheet_name() {
    let wb = extract_charts(&fixture()).unwrap();
    let pos = wb.sheets[0].charts[0].position.as_ref().unwrap();
    assert_eq!(pos.sheet, "Sales");
}

#[test]
fn expenses_chart_position_sheet_name() {
    let wb = extract_charts(&fixture()).unwrap();
    let pos = wb.sheets[1].charts[0].position.as_ref().unwrap();
    assert_eq!(pos.sheet, "Expenses");
}

// ── top_left / bottom_right A1 notation ──────────────────────────────────────

#[test]
fn sales_chart_position_top_left_a1() {
    let wb = extract_charts(&fixture()).unwrap();
    let pos = wb.sheets[0].charts[0].position.as_ref().unwrap();
    // anchor.col_start=0, row_start=0 → A1
    assert_eq!(pos.top_left, "A1");
}

#[test]
fn sales_chart_position_bottom_right_format() {
    let wb = extract_charts(&fixture()).unwrap();
    let pos = wb.sheets[0].charts[0].position.as_ref().unwrap();
    // bottom_right must be a non-empty A1-notation string
    assert!(!pos.bottom_right.is_empty());
    // must start with a letter
    assert!(pos
        .bottom_right
        .chars()
        .next()
        .unwrap()
        .is_ascii_uppercase());
    // must end with a digit
    assert!(pos.bottom_right.chars().last().unwrap().is_ascii_digit());
}

#[test]
fn position_top_left_and_bottom_right_differ_for_two_cell() {
    let wb = extract_charts(&fixture()).unwrap();
    let pos = wb.sheets[0].charts[0].position.as_ref().unwrap();
    // twoCellAnchor: bottom_right must be further than top_left
    assert_ne!(
        pos.top_left, pos.bottom_right,
        "twoCellAnchor chart must have distinct top_left and bottom_right"
    );
}

// ── consistency with ChartAnchor ─────────────────────────────────────────────

#[test]
fn position_consistent_with_anchor_col_start() {
    let wb = extract_charts(&fixture()).unwrap();
    let chart = &wb.sheets[0].charts[0];
    let anchor = chart.anchor.as_ref().unwrap();
    let pos = chart.position.as_ref().unwrap();
    let expected_tl = ChartPosition::cell_address(anchor.col_start, anchor.row_start);
    assert_eq!(pos.top_left, expected_tl);
}

#[test]
fn position_consistent_with_anchor_col_end() {
    let wb = extract_charts(&fixture()).unwrap();
    let chart = &wb.sheets[0].charts[0];
    let anchor = chart.anchor.as_ref().unwrap();
    let pos = chart.position.as_ref().unwrap();
    let expected_br = ChartPosition::cell_address(anchor.col_end, anchor.row_end);
    assert_eq!(pos.bottom_right, expected_br);
}

// ── width_emu / height_emu (None for twoCellAnchor) ──────────────────────────

#[test]
fn two_cell_position_width_emu_none() {
    let wb = extract_charts(&fixture()).unwrap();
    let pos = wb.sheets[0].charts[0].position.as_ref().unwrap();
    assert!(
        pos.width_emu.is_none(),
        "twoCellAnchor chart should not have width_emu"
    );
}

#[test]
fn two_cell_position_height_emu_none() {
    let wb = extract_charts(&fixture()).unwrap();
    let pos = wb.sheets[0].charts[0].position.as_ref().unwrap();
    assert!(pos.height_emu.is_none());
}

// ── col_to_letter correctness ─────────────────────────────────────────────────

#[test]
fn col_letter_a() {
    assert_eq!(ChartPosition::col_to_letter(0), "A");
}
#[test]
fn col_letter_z() {
    assert_eq!(ChartPosition::col_to_letter(25), "Z");
}
#[test]
fn col_letter_aa() {
    assert_eq!(ChartPosition::col_to_letter(26), "AA");
}
#[test]
fn col_letter_ab() {
    assert_eq!(ChartPosition::col_to_letter(27), "AB");
}

// ── cell_address correctness ──────────────────────────────────────────────────

#[test]
fn cell_address_a1() {
    assert_eq!(ChartPosition::cell_address(0, 0), "A1");
}
#[test]
fn cell_address_b2() {
    assert_eq!(ChartPosition::cell_address(1, 1), "B2");
}
#[test]
fn cell_address_z10() {
    assert_eq!(ChartPosition::cell_address(25, 9), "Z10");
}
#[test]
fn cell_address_aa1() {
    assert_eq!(ChartPosition::cell_address(26, 0), "AA1");
}

// ── regression: all existing fixture charts have position ─────────────────────

#[test]
fn all_3d_charts_have_position() {
    let wb = extract_charts(&fixture_3d()).unwrap();
    for sheet in &wb.sheets {
        for chart in &sheet.charts {
            assert!(
                chart.position.is_some(),
                "3-D chart {} must have position",
                chart.chart_path
            );
        }
    }
}

#[test]
fn all_surface_charts_have_position() {
    let wb = extract_charts(&fixture_surface()).unwrap();
    for sheet in &wb.sheets {
        for chart in &sheet.charts {
            assert!(
                chart.position.is_some(),
                "surface chart {} must have position",
                chart.chart_path
            );
        }
    }
}

#[test]
fn position_sheet_matches_parent_sheet() {
    // For every chart in every fixture, position.sheet must equal the
    // SheetCharts name that owns it.
    let wb = extract_charts(&fixture()).unwrap();
    for sheet in &wb.sheets {
        for chart in &sheet.charts {
            if let Some(pos) = &chart.position {
                assert_eq!(
                    pos.sheet, sheet.name,
                    "chart {} position.sheet must match parent sheet name",
                    chart.chart_path
                );
            }
        }
    }
}
