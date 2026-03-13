//! # sheetforge-charts
//!
//! High-performance Rust library for extracting chart metadata from Excel
//! (`.xlsx`) files.  Designed as the native core for Python (PyO3) and
//! TypeScript/JavaScript (napi-rs) bindings.
//!
//! ## Quick start
//!
//! ```rust,no_run
//! use sheetforge_charts::extract_charts;
//!
//! let workbook = extract_charts("workbook.xlsx").unwrap();
//!
//! if let Some(theme) = &workbook.theme {
//!     println!("accent1 = {:?}", theme.accent1());
//! }
//!
//! for sheet in &workbook.sheets {
//!     for chart in &sheet.charts {
//!         println!("{}: {:?}", chart.chart_path, chart.chart_type);
//!     }
//! }
//! ```
//!
//! ## Performance design
//!
//! The pipeline has two phases, deliberately separated:
//!
//! ### Phase A — serial I/O
//! All ZIP reads happen on a single thread.  `ZipArchive<BufReader<File>>`
//! is not `Send` (it holds a file cursor), so archive access cannot be
//! shared across threads.  We read every chart's raw bytes into memory
//! and collect the complete set of skeleton charts before leaving this phase.
//!
//! ### Phase B — parallel parse
//! Once every `(chart_path, raw_bytes, anchor)` tuple is in memory, we hand
//! the Vec to Rayon's `par_iter`.  Each worker calls `chart_parser::parse_bytes`
//! on its own slice — pure CPU, no I/O, no shared state — then returns a fully
//! parsed `Chart`.  The results are collected back in original order.
//!
//! On a 50-chart workbook, Phase B completes in ≈ 1/N_CORES the serial time.

pub mod archive;
pub mod model;
pub mod openxml;
pub mod parser;

pub use model::workbook::WorkbookCharts;

use anyhow::Result;
use rayon::prelude::*;

use crate::{
    archive::zip_reader::{open_xlsx, read_entry_bytes},
    model::chart::{Chart, ChartAnchor},
    openxml::relationships::{parse_for_part, rel_type, resolve_relative},
    parser::chart_parser,
};

// ── Public API ────────────────────────────────────────────────────────────────

/// Extract all chart metadata from an Excel (`.xlsx`) file at `path`.
///
/// Charts are parsed **concurrently** using Rayon.  The function is safe to
/// call from multiple threads simultaneously (each call opens its own file
/// handle and archive).
pub fn extract_charts(path: &str) -> Result<WorkbookCharts> {
    extract_charts_impl(path)
}

// ── Implementation ────────────────────────────────────────────────────────────

fn extract_charts_impl(path: &str) -> Result<WorkbookCharts> {
    // ═══════════════════════════════════════════════════════════════════════
    // PHASE A — Serial I/O
    // All ZIP reads happen here.  Nothing in this phase is parallel.
    // ═══════════════════════════════════════════════════════════════════════

    let mut archive = open_xlsx(path)?;

    // ── A1: content types → workbook path ─────────────────────────────────
    let content_types = openxml::content_types::parse(&mut archive)?;
    let workbook_path = content_types.workbook_path()?;

    // ── A2: relationship chain → chart skeletons ──────────────────────────
    let mut resolver = openxml::relationships::RelationshipResolver::new();
    let mut workbook = parser::workbook_parser::parse(&mut archive, &workbook_path)?;

    parser::sheet_parser::resolve_charts(
        &mut archive,
        &mut resolver,
        &workbook_path,
        &mut workbook.sheets,
    )?;

    // ── A3: theme ─────────────────────────────────────────────────────────
    workbook.theme = load_theme(&mut archive, &workbook_path);

    // ── A4: read all chart bytes into memory ──────────────────────────────
    // Collect (chart_path, raw_bytes, anchor) for every chart across every
    // sheet.  We preserve insertion order via a flat index so we can write
    // parsed results back into the correct Sheet/Chart slots in Phase B.
    //
    // Layout of `chart_jobs`:
    //   index → (sheet_idx, chart_idx_in_sheet, chart_path, bytes, anchor)
    let chart_jobs: Vec<(usize, usize, String, Vec<u8>, Option<ChartAnchor>)> = {
        let mut jobs = Vec::new();
        for (si, sheet) in workbook.sheets.iter().enumerate() {
            for (ci, chart) in sheet.charts.iter().enumerate() {
                match read_entry_bytes(&mut archive, &chart.chart_path) {
                    Ok(bytes) => {
                        jobs.push((
                            si,
                            ci,
                            chart.chart_path.clone(),
                            bytes,
                            chart.anchor.clone(),
                        ));
                    }
                    Err(e) => {
                        eprintln!(
                            "Warning: cannot read chart bytes '{}': {e:#}",
                            chart.chart_path
                        );
                    }
                }
            }
        }
        jobs
    };

    // Drop the archive — we no longer need the file handle.
    drop(archive);

    // ═══════════════════════════════════════════════════════════════════════
    // PHASE B — Parallel XML parse
    // Each chart is parsed independently.  No shared mutable state.
    // Rayon distributes work across all available CPU cores.
    // ═══════════════════════════════════════════════════════════════════════

    let parsed_charts: Vec<(usize, usize, Chart)> = chart_jobs
        .into_par_iter()
        .filter_map(|(si, ci, chart_path, bytes, anchor)| {
            match chart_parser::parse_bytes(&bytes, &chart_path) {
                Ok(mut chart) => {
                    // Restore the anchor that was resolved in Phase A.
                    // parse_bytes has no access to drawing XML so it always
                    // sets anchor = None; we put the real value back here.
                    chart.anchor = anchor;
                    Some((si, ci, chart))
                }
                Err(e) => {
                    eprintln!("Warning: could not parse chart '{chart_path}': {e:#}");
                    None  // keep the skeleton that was already in workbook.sheets
                }
            }
        })
        .collect();

    // ── Reassemble: write parsed charts back into the workbook ────────────
    for (si, ci, chart) in parsed_charts {
        workbook.sheets[si].charts[ci] = chart;
    }

    Ok(workbook)
}

// ── Theme helper ──────────────────────────────────────────────────────────────

/// Attempt to load the theme.  Non-fatal on any error.
fn load_theme(
    archive: &mut archive::zip_reader::XlsxArchive,
    workbook_path: &str,
) -> Option<model::theme::Theme> {
    let rels = match parse_for_part(archive, workbook_path) {
        Ok(r) => r,
        Err(e) => {
            eprintln!("Warning: could not read workbook rels for theme: {e:#}");
            return None;
        }
    };

    let theme_path = rels
        .by_type(rel_type::THEME)
        .map(|r| resolve_relative(workbook_path, &r.target))
        .next()?;

    match openxml::theme_parser::parse(archive, &theme_path) {
        Ok(t) => Some(t),
        Err(e) => {
            eprintln!("Warning: could not parse theme '{theme_path}': {e:#}");
            None
        }
    }
}

#[cfg(feature = "python")]
mod python_bindings;

#[cfg(feature = "nodejs")]
mod nodejs_bindings;
