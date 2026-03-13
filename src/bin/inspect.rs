//! Manual inspection binary — shows the full Phase 1–9 output.
//!
//! Usage:
//!   cargo run --bin inspect -- tests/fixtures/test_charts.xlsx
//!   cargo run --bin inspect -- path/to/any.xlsx

use sheetforge_charts::{
    extract_charts,
    model::color::{ColorSpec, Fill, GradientDirection},
};

fn main() {
    let path = std::env::args().nth(1).unwrap_or_else(|| {
        eprintln!("Usage: inspect <path-to-xlsx>");
        eprintln!("Example: cargo run --bin inspect -- tests/fixtures/test_charts.xlsx");
        std::process::exit(1);
    });

    println!("Inspecting: {path}\n");

    match extract_charts(&path) {
        Err(e) => {
            eprintln!("ERROR: {e:#}");
            std::process::exit(1);
        }
        Ok(wb) => {
            println!("Sheets : {}", wb.sheets.len());
            println!("Charts : {}", wb.chart_count());

            // ── Theme ─────────────────────────────────────────────────────────
            println!("{}", "─".repeat(70));
            match &wb.theme {
                None => println!("Theme  : (none)"),
                Some(t) => {
                    println!(
                        "Theme  : {}",
                        t.name.as_deref().unwrap_or("(unnamed)")
                    );
                    for (name, rgb) in t.all_colors() {
                        println!("  {:<10} #{}", name, rgb.to_hex());
                    }
                }
            }

            println!("{}", "─".repeat(70));

            for sheet in &wb.sheets {
                println!("\nSheet [{:>2}] \"{}\"", sheet.index, sheet.name);
                println!(
                    "  part   : {}",
                    sheet.part_path.as_deref().unwrap_or("(unresolved)")
                );

                if sheet.charts.is_empty() {
                    println!("  charts : none");
                    continue;
                }

                for (ci, chart) in sheet.charts.iter().enumerate() {
                    println!(
                        "\n  ┌─ Chart [{}] ─────────────────────────────────────",
                        ci
                    );
                    println!("  │  path    : {}", chart.chart_path);
                    println!("  │  type    : {:?}", chart.chart_type);
                    println!(
                        "  │  title   : {}",
                        chart.title.as_deref().unwrap_or("(none)")
                    );
                    println!(
                        "  │  style   : {}",
                        chart.style.map(|s| s.to_string()).unwrap_or_else(|| "(none)".into())
                    );
                    println!("  │  legend  : {:?}", chart.legend_position);
                    println!("  │  grouping: {:?}", chart.plot_area.grouping);

                    // Anchor / position
                    match &chart.anchor {
                        None => println!("  │  anchor  : (none)"),
                        Some(a) => println!(
                            "  │  anchor  : rows {}–{}  cols {}–{}  \
                             (span {} rows × {} cols)",
                            a.row_start, a.row_end,
                            a.col_start, a.col_end,
                            a.row_span(), a.col_span(),
                        ),
                    }

                    // 3-D view (only printed for 3-D chart types)
                    if chart.chart_type.is_3d() {
                        match &chart.view_3d {
                            None => println!("  │  3D view : (absent — <c:view3D> not in XML)"),
                            Some(v) => {
                                println!("  │  3D view :");
                                match v.rotation_x {
                                    Some(x) => println!("  │    rotX         : {x}°"),
                                    None    => println!("  │    rotX         : (default)"),
                                }
                                match v.rotation_y {
                                    Some(y) => println!("  │    rotY         : {y}°"),
                                    None    => println!("  │    rotY         : (default)"),
                                }
                                match v.right_angle_axes {
                                    Some(true)  => println!("  │    rAngAx       : true  (orthographic)"),
                                    Some(false) => println!("  │    rAngAx       : false (perspective)"),
                                    None        => println!("  │    rAngAx       : (default)"),
                                }
                                match v.perspective {
                                    Some(p) => println!("  │    perspective  : {p}"),
                                    None    => println!("  │    perspective  : (default)"),
                                }
                            }
                        }

                        // 3-D geometry surfaces (floor, side-wall, back-wall)
                        match &chart.surface {
                            None => println!("  │  3D surf : (none — no <c:floor>/<c:sideWall>/<c:backWall>)"),
                            Some(surf) => {
                                println!("  │  3D surf :");
                                println!(
                                    "  │    floor     : {}",
                                    fmt_fill(surf.floor_fill.as_ref(), wb.theme.as_ref())
                                );
                                println!(
                                    "  │    sideWall  : {}",
                                    fmt_fill(surf.side_wall_fill.as_ref(), wb.theme.as_ref())
                                );
                                println!(
                                    "  │    backWall  : {}",
                                    fmt_fill(surf.back_wall_fill.as_ref(), wb.theme.as_ref())
                                );
                            }
                        }
                    }

                    // Chart-space fill
                    println!(
                        "  │  chart fill : {}",
                        fmt_fill(chart.chart_fill.as_ref(), wb.theme.as_ref())
                    );
                    // Plot-area fill
                    println!(
                        "  │  area fill  : {}",
                        fmt_fill(chart.plot_area.fill.as_ref(), wb.theme.as_ref())
                    );

                    // ── Series ────────────────────────────────────────────────
                    println!("  │  series  : {}", chart.series.len());
                    for s in &chart.series {
                        println!(
                            "  │    [{:>2}] name_ref : {}",
                            s.index,
                            s.name_ref.as_ref().map(|r| r.formula.as_str()).unwrap_or("—")
                        );
                        if let Some(n) = &s.name {
                            println!("  │         name     : {n}");
                        }
                        println!(
                            "  │         cat_ref  : {}",
                            s.category_ref.as_ref().map(|r| r.formula.as_str()).unwrap_or("—")
                        );
                        println!(
                            "  │         val_ref  : {}",
                            s.value_ref.as_ref().map(|r| r.formula.as_str()).unwrap_or("—")
                        );
                        if let Some(cache) = &s.value_cache {
                            println!(
                                "  │         cache    : {:?} (fmt={}, state={:?})",
                                cache.values,
                                cache.format_code.as_deref().unwrap_or("—"),
                                s.value_cache_state,
                            );
                        }
                        if let Some(cats) = &s.category_values {
                            println!("  │         cat vals : {:?}", cats.values);
                        }
                        // Fill
                        println!(
                            "  │         fill     : {}",
                            fmt_fill(s.fill.as_ref(), wb.theme.as_ref())
                        );
                    }

                    // ── Axes ──────────────────────────────────────────────────
                    println!("  │  axes    : {}", chart.axes.len());
                    for ax in &chart.axes {
                        println!(
                            "  │    id={} type={:?} pos={:?} crossAx={:?} fmt={}",
                            ax.id,
                            ax.axis_type,
                            ax.position,
                            ax.cross_axis_id,
                            ax.number_format.as_deref().unwrap_or("—")
                        );
                    }
                    println!("  └─────────────────────────────────────────────────────");
                }
            }

            println!("\n{}", "─".repeat(70));
            println!("Done (Phases 1–9).");
        }
    }
}

// ── Fill formatting helpers ───────────────────────────────────────────────────

fn fmt_fill(
    fill: Option<&Fill>,
    theme: Option<&sheetforge_charts::model::theme::Theme>,
) -> String {
    match fill {
        None => "(none — no spPr)".into(),
        Some(Fill::None) => "noFill (explicit transparent; color from style/theme)".into(),
        Some(Fill::Pattern) => "pattern".into(),
        Some(Fill::Solid(spec)) => {
            let resolved = spec.resolve(theme)
                .map(|rgb| format!(" → #{}", rgb.to_hex()))
                .unwrap_or_else(|| " → (needs theme)".into());
            format!("solid  {}{}", fmt_color_spec(spec), resolved)
        }
        Some(Fill::Gradient(grad)) => {
            let dir = match grad.direction.as_ref() {
                None => "no-dir".into(),
                Some(GradientDirection::Linear { angle_deg, .. }) => {
                    format!("linear {:.0}°", angle_deg)
                }
                Some(GradientDirection::Path(p)) => format!("path({p})"),
            };
            let stops: Vec<String> = grad
                .resolve_stops(theme)
                .iter()
                .map(|(pos, rgb)| format!("{:.0}%=#{}", pos * 100.0, rgb.to_hex()))
                .collect();
            format!("gradient {} [{}]", dir, stops.join(", "))
        }
    }
}

fn fmt_color_spec(spec: &ColorSpec) -> String {
    match spec {
        ColorSpec::Srgb(rgb, mods) => {
            format!("srgb(#{}{}) ", rgb.to_hex(), fmt_mods(mods))
        }
        ColorSpec::Sys(rgb, mods) => {
            format!("sys(#{}{}) ", rgb.to_hex(), fmt_mods(mods))
        }
        ColorSpec::Scheme(slot, mods) => {
            format!("scheme({}{}) ", slot.as_str(), fmt_mods(mods))
        }
        ColorSpec::Preset(name, mods) => {
            format!("preset({}{}) ", name, fmt_mods(mods))
        }
    }
}

fn fmt_mods(mods: &[sheetforge_charts::model::color::ColorMod]) -> String {
    if mods.is_empty() {
        return String::new();
    }
    let parts: Vec<String> = mods.iter().map(|m| format!("{m:?}")).collect();
    format!(" [{}]", parts.join(", "))
}
