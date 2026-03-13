//! Sheet-level relationship walker.
//!
//! For each sheet discovered by [`crate::parser::workbook_parser`] this module:
//!
//! 1. Resolves the sheet's ZIP path via the workbook `.rels`.
//! 2. Reads the sheet's own `.rels` to find attached drawing parts.
//! 3. Reads each drawing's XML to collect `<c:chart r:id="…"/>` anchors.
//! 4. Resolves each anchor through the drawing's `.rels` to get the
//!    chart part path.
//! 5. Records a [`Chart`] skeleton (path only) on the owning [`SheetCharts`].
//!
//! At the end of this phase every [`SheetCharts`] carries:
//! * `part_path`     — the resolved worksheet ZIP path
//! * `drawing_paths` — every drawing attached to the sheet
//! * `charts`        — skeleton [`Chart`] values ready for Phase 3 parsing
//!
//! No chart XML is parsed here — that happens in Phase 3.

use anyhow::{Context, Result};

use crate::{
    archive::zip_reader::XlsxArchive,
    model::{chart::Chart, workbook::SheetCharts},
    openxml::{
        drawing,
        relationships::{rel_type, RelationshipResolver},
    },
};

// ── Public entry point ────────────────────────────────────────────────────────

/// Populate `sheets` with part paths, drawing paths, and chart skeletons.
///
/// `workbook_path` is the resolved ZIP path of `workbook.xml`
/// (e.g. `"xl/workbook.xml"`).
///
/// Mutates `sheets` in place; returns the same vec for ergonomics.
pub fn resolve_charts(
    archive: &mut XlsxArchive,
    resolver: &mut RelationshipResolver,
    workbook_path: &str,
    sheets: &mut Vec<SheetCharts>,
) -> Result<()> {
    for sheet in sheets.iter_mut() {
        resolve_sheet(archive, resolver, workbook_path, sheet)
            .with_context(|| format!("Failed resolving charts for sheet '{}'", sheet.name))?;
    }
    Ok(())
}

// ── Per-sheet logic ───────────────────────────────────────────────────────────

fn resolve_sheet(
    archive: &mut XlsxArchive,
    resolver: &mut RelationshipResolver,
    workbook_path: &str,
    sheet: &mut SheetCharts,
) -> Result<()> {
    // ── Step 1: workbook → sheet part path ───────────────────────────────────
    let sheet_path = resolver
        .resolve_target(archive, workbook_path, &sheet.relationship_id)
        .with_context(|| {
            format!(
                "Cannot resolve sheet '{}' (r:id={})",
                sheet.name, sheet.relationship_id
            )
        })?;

    sheet.set_part_path(&sheet_path);

    // ── Step 2: sheet → drawing parts ────────────────────────────────────────
    let drawing_paths = resolver.targets_of_type(archive, &sheet_path, rel_type::DRAWING)?;

    for drawing_path in drawing_paths {
        sheet.add_drawing_path(&drawing_path);

        // ── Step 3: drawing XML → chart r:id anchors ─────────────────────────
        let chart_refs = drawing::parse(archive, &drawing_path)
            .with_context(|| format!("Cannot read drawing: {drawing_path}"))?;

        // ── Step 4: drawing → chart part paths ───────────────────────────────
        for chart_ref in &chart_refs.refs {
            let chart_path = resolver
                .resolve_target(archive, &drawing_path, &chart_ref.rel_id)
                .with_context(|| {
                    format!(
                        "Cannot resolve chart r:id='{}' in drawing '{}'",
                        chart_ref.rel_id, drawing_path
                    )
                })?;

            // ── Step 5: record skeleton chart with anchor ─────────────────────
            let mut chart = Chart::new_skeleton(chart_path);
            chart.anchor = chart_ref.anchor.clone();
            sheet.charts.push(chart);
        }
    }

    Ok(())
}

// ── Tests ─────────────────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    // These tests validate the resolve_sheet logic using pure relationship
    // data (no real archive).  They exercise the same code paths that run
    // against a real XLSX in integration tests.
    //
    // Because RelationshipResolver caches by archive lookup we test the
    // underlying chain logic through the relationship and drawing modules
    // directly.  Full end-to-end integration testing requires a real .xlsx
    // fixture (see tests/integration_test.rs, added in a later phase).

    use crate::openxml::{
        drawing,
        relationships::{parse_xml as parse_rels, rel_type, resolve_relative},
    };

    #[test]
    fn sheet_to_drawing_resolution() {
        let sheet_rels_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
    Target="../drawings/drawing1.xml"/>
</Relationships>"#;

        let rels = parse_rels(sheet_rels_xml).unwrap();
        let drawings: Vec<String> = rels
            .by_type(rel_type::DRAWING)
            .map(|r| rels.resolve_target(r, "xl/worksheets/sheet1.xml"))
            .collect();

        assert_eq!(drawings, vec!["xl/drawings/drawing1.xml"]);
    }

    #[test]
    fn drawing_to_chart_resolution() {
        let drawing_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr
  xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <xdr:twoCellAnchor>
    <xdr:graphicFrame>
      <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:graphicData>
          <c:chart r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

        let drawing_rels_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
    Target="../charts/chart1.xml"/>
</Relationships>"#;

        let refs = drawing::parse_xml(drawing_xml).unwrap();
        assert_eq!(refs.len(), 1);
        assert_eq!(refs.refs[0].rel_id, "rId1");

        let dr_rels = parse_rels(drawing_rels_xml).unwrap();
        let chart_path = dr_rels
            .resolve_id("rId1", "xl/drawings/drawing1.xml")
            .unwrap();
        assert_eq!(chart_path, "xl/charts/chart1.xml");
    }

    #[test]
    fn sheet_with_no_drawings_produces_no_charts() {
        // A sheet .rels with only a styles relationship — no drawing
        let sheet_rels_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    Target="../styles.xml"/>
</Relationships>"#;

        let rels = parse_rels(sheet_rels_xml).unwrap();
        let drawings: Vec<_> = rels.by_type(rel_type::DRAWING).collect();
        assert!(drawings.is_empty(), "no drawing rels should mean no charts");
    }

    #[test]
    fn multi_chart_drawing_produces_multiple_skeletons() {
        let drawing_rels_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
    Target="../charts/chart1.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
    Target="../charts/chart2.xml"/>
  <Relationship Id="rId3"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
    Target="../charts/chart3.xml"/>
</Relationships>"#;

        let dr_rels = parse_rels(drawing_rels_xml).unwrap();
        let chart_paths: Vec<String> = dr_rels
            .by_type(rel_type::CHART)
            .map(|r| dr_rels.resolve_target(r, "xl/drawings/drawing1.xml"))
            .collect();

        assert_eq!(chart_paths.len(), 3);
        assert_eq!(chart_paths[0], "xl/charts/chart1.xml");
        assert_eq!(chart_paths[2], "xl/charts/chart3.xml");
    }
}
