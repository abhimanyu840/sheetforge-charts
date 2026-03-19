//! Parser for `xl/drawings/drawingN.xml`.
//!
//! Extracts every `<xdr:twoCellAnchor>` **and** `<xdr:oneCellAnchor>` and
//! pairs each with the `<c:chart r:id="…"/>` reference nested inside it.
//!
//! ## XML structures
//!
//! ### twoCellAnchor (most common)
//! ```xml
//! <xdr:twoCellAnchor>
//!   <xdr:from>
//!     <xdr:col>0</xdr:col>  <xdr:colOff>0</xdr:colOff>
//!     <xdr:row>0</xdr:row>  <xdr:rowOff>0</xdr:rowOff>
//!   </xdr:from>
//!   <xdr:to>
//!     <xdr:col>8</xdr:col>  <xdr:colOff>0</xdr:colOff>
//!     <xdr:row>15</xdr:row> <xdr:rowOff>0</xdr:rowOff>
//!   </xdr:to>
//!   <xdr:graphicFrame>…<c:chart r:id="rId1"/>…</xdr:graphicFrame>
//! </xdr:twoCellAnchor>
//! ```
//!
//! ### oneCellAnchor (less common)
//! ```xml
//! <xdr:oneCellAnchor>
//!   <xdr:from>
//!     <xdr:col>5</xdr:col>  <xdr:colOff>0</xdr:colOff>
//!     <xdr:row>1</xdr:row>  <xdr:rowOff>0</xdr:rowOff>
//!   </xdr:from>
//!   <xdr:ext cx="3000000" cy="2000000"/>   ← EMU width/height
//!   <xdr:graphicFrame>…<c:chart r:id="rId1"/>…</xdr:graphicFrame>
//! </xdr:oneCellAnchor>
//! ```

use anyhow::{Context, Result};
use quick_xml::{events::Event, Reader};

use crate::{
    archive::zip_reader::{read_entry_to_string, XlsxArchive},
    model::chart::ChartAnchor,
};

// ── Public types ──────────────────────────────────────────────────────────────

/// A `<c:chart r:id="…"/>` element paired with the anchor that positions it.
#[derive(Debug, Clone, PartialEq)]
pub struct ChartRef {
    /// The `r:id` value from `<c:chart r:id="…"/>`.
    pub rel_id: String,
    /// Anchor from the surrounding `<xdr:twoCellAnchor>`.
    /// `None` only for charts embedded in a non-twoCellAnchor container
    /// (e.g. `<xdr:absoluteAnchor>`), which is very rare in practice.
    pub anchor: Option<ChartAnchor>,
}

/// All chart refs found in a single drawing part.
#[derive(Debug, Default)]
pub struct DrawingChartRefs {
    pub refs: Vec<ChartRef>,
}

impl DrawingChartRefs {
    pub fn is_empty(&self) -> bool {
        self.refs.is_empty()
    }
    pub fn len(&self) -> usize {
        self.refs.len()
    }
}

// ── Entry points ──────────────────────────────────────────────────────────────

pub fn parse(archive: &mut XlsxArchive, drawing_path: &str) -> Result<DrawingChartRefs> {
    let xml = read_entry_to_string(archive, drawing_path)
        .with_context(|| format!("Cannot read drawing part: {drawing_path}"))?;
    parse_xml(&xml).with_context(|| format!("Failed to parse drawing XML: {drawing_path}"))
}

pub(crate) fn parse_xml(xml: &str) -> Result<DrawingChartRefs> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut result = DrawingChartRefs::default();
    let mut st = State::default();

    loop {
        match reader.read_event()? {
            Event::Start(ref e) | Event::Empty(ref e) => {
                let ln = e.local_name();
                let tag = std::str::from_utf8(ln.as_ref()).unwrap_or("");

                match tag {
                    "twoCellAnchor" => st.begin_anchor(AnchorKind::TwoCell),
                    "oneCellAnchor" => st.begin_anchor(AnchorKind::OneCell),
                    "from" => st.corner = Corner::From,
                    "to" => st.corner = Corner::To,
                    "col" => st.pending = Some(Field::Col),
                    "colOff" => st.pending = Some(Field::ColOff),
                    "row" => st.pending = Some(Field::Row),
                    "rowOff" => st.pending = Some(Field::RowOff),
                    // <xdr:ext cx="…" cy="…"/> — present in oneCellAnchor
                    "ext" if st.anchor_kind == AnchorKind::OneCell => {
                        let dec = reader.decoder();
                        for attr in e.attributes() {
                            let attr = attr.context("Malformed attr in <xdr:ext>")?;
                            match attr.key.local_name().as_ref() {
                                b"cx" => {
                                    st.ext_cx =
                                        attr.decode_and_unescape_value(dec)?.parse().unwrap_or(0);
                                }
                                b"cy" => {
                                    st.ext_cy =
                                        attr.decode_and_unescape_value(dec)?.parse().unwrap_or(0);
                                }
                                _ => {}
                            }
                        }
                    }
                    "chart" => {
                        if let Some(rel_id) = extract_r_id(e, reader.decoder())? {
                            st.pending_rel_id = Some(rel_id);
                        }
                    }
                    _ => {}
                }
            }

            Event::Text(ref e) => {
                if let Some(field) = st.pending.take() {
                    let raw = e.unescape().unwrap_or_default();
                    st.apply(field, raw.trim());
                }
            }

            Event::End(ref e) => {
                let ln = e.local_name();
                let tag = std::str::from_utf8(ln.as_ref()).unwrap_or("");
                match tag {
                    "from" | "to" => {
                        st.corner = Corner::None;
                    }
                    "twoCellAnchor" => {
                        if let Some(rel_id) = st.pending_rel_id.take() {
                            result.refs.push(ChartRef {
                                rel_id,
                                anchor: Some(st.build_two_cell()),
                            });
                        }
                        st.reset();
                    }
                    "oneCellAnchor" => {
                        if let Some(rel_id) = st.pending_rel_id.take() {
                            result.refs.push(ChartRef {
                                rel_id,
                                anchor: Some(st.build_one_cell()),
                            });
                        }
                        st.reset();
                    }
                    _ => {}
                }
            }

            Event::Eof => break,
            _ => {}
        }
    }

    Ok(result)
}

// ── Internal state machine ────────────────────────────────────────────────────

#[derive(Debug, Default, Clone, Copy, PartialEq)]
enum Corner {
    #[default]
    None,
    From,
    To,
}

/// Which anchor element we are currently inside.
#[derive(Debug, Default, Clone, Copy, PartialEq)]
enum AnchorKind {
    #[default]
    None,
    TwoCell,
    OneCell,
}

#[derive(Debug, Clone, Copy)]
enum Field {
    Col,
    ColOff,
    Row,
    RowOff,
}

#[derive(Debug, Default)]
struct State {
    anchor_kind: AnchorKind,
    corner: Corner,
    pending: Option<Field>,
    pending_rel_id: Option<String>,
    from_col: u32,
    from_col_off: i64,
    from_row: u32,
    from_row_off: i64,
    to_col: u32,
    to_col_off: i64,
    to_row: u32,
    to_row_off: i64,
    /// EMU width from `<xdr:ext cx="…"/>` (oneCellAnchor only).
    ext_cx: i64,
    /// EMU height from `<xdr:ext cy="…"/>` (oneCellAnchor only).
    ext_cy: i64,
}

impl State {
    fn begin_anchor(&mut self, kind: AnchorKind) {
        self.reset();
        self.anchor_kind = kind;
    }

    fn reset(&mut self) {
        self.anchor_kind = AnchorKind::None;
        self.corner = Corner::None;
        self.pending = None;
        self.from_col = 0;
        self.from_col_off = 0;
        self.from_row = 0;
        self.from_row_off = 0;
        self.to_col = 0;
        self.to_col_off = 0;
        self.to_row = 0;
        self.to_row_off = 0;
        self.ext_cx = 0;
        self.ext_cy = 0;
        // pending_rel_id is consumed at anchor-close, not here
    }

    fn apply(&mut self, field: Field, text: &str) {
        macro_rules! parse {
            ($f:expr) => {
                text.parse().unwrap_or(0)
            };
        }
        match self.corner {
            Corner::From => match field {
                Field::Col => {
                    self.from_col = parse!(field);
                }
                Field::ColOff => {
                    self.from_col_off = parse!(field);
                }
                Field::Row => {
                    self.from_row = parse!(field);
                }
                Field::RowOff => {
                    self.from_row_off = parse!(field);
                }
            },
            Corner::To => match field {
                Field::Col => {
                    self.to_col = parse!(field);
                }
                Field::ColOff => {
                    self.to_col_off = parse!(field);
                }
                Field::Row => {
                    self.to_row = parse!(field);
                }
                Field::RowOff => {
                    self.to_row_off = parse!(field);
                }
            },
            Corner::None => {}
        }
    }

    fn build_two_cell(&self) -> ChartAnchor {
        ChartAnchor {
            col_start: self.from_col,
            col_off: self.from_col_off,
            row_start: self.from_row,
            row_off: self.from_row_off,
            col_end: self.to_col,
            col_end_off: self.to_col_off,
            row_end: self.to_row,
            row_end_off: self.to_row_off,
        }
    }

    /// Build a `ChartAnchor` from a oneCellAnchor: col/row end = col/row start
    /// (we don't know the exact end cell without column/row pixel data).
    fn build_one_cell(&self) -> ChartAnchor {
        ChartAnchor {
            col_start: self.from_col,
            col_off: self.from_col_off,
            row_start: self.from_row,
            row_off: self.from_row_off,
            col_end: self.from_col,
            col_end_off: self.ext_cx, // store EMU width in col_end_off for callers
            row_end: self.from_row,
            row_end_off: self.ext_cy, // store EMU height in row_end_off
        }
    }
}

// ── Attribute helper ──────────────────────────────────────────────────────────

fn extract_r_id(
    e: &quick_xml::events::BytesStart<'_>,
    decoder: quick_xml::Decoder,
) -> Result<Option<String>> {
    for attr in e.attributes() {
        let attr = attr.context("Malformed attribute in <c:chart>")?;
        if attr.key.local_name().as_ref() == b"id" {
            let val = attr.decode_and_unescape_value(decoder)?.into_owned();
            if !val.is_empty() {
                return Ok(Some(val));
            }
        }
    }
    Ok(None)
}

// ── Tests ─────────────────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;

    // ── fixtures ──────────────────────────────────────────────────────────────

    /// Matches the exact XML our Python fixture generator writes
    const ZERO_OFFSETS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr
  xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>15</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame><xdr:nvGraphicFramePr>
      <xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/>
    </xdr:nvGraphicFramePr>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart r:id="rId1"/>
      </a:graphicData></a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

    /// Non-zero offsets and a non-zero from corner
    const NON_ZERO: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr
  xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <xdr:twoCellAnchor>
    <xdr:from>
      <xdr:col>1</xdr:col><xdr:colOff>12345</xdr:colOff>
      <xdr:row>2</xdr:row><xdr:rowOff>67890</xdr:rowOff>
    </xdr:from>
    <xdr:to>
      <xdr:col>9</xdr:col><xdr:colOff>11111</xdr:colOff>
      <xdr:row>18</xdr:row><xdr:rowOff>22222</xdr:rowOff>
    </xdr:to>
    <xdr:graphicFrame>
      <a:graphic><a:graphicData>
        <c:chart r:id="rId3"/>
      </a:graphicData></a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

    /// Two charts side-by-side on one sheet
    const TWO_CHARTS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr
  xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame>
      <a:graphic><a:graphicData><c:chart r:id="rId1"/></a:graphicData></a:graphic>
    </xdr:graphicFrame><xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>12</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame>
      <a:graphic><a:graphicData><c:chart r:id="rId2"/></a:graphicData></a:graphic>
    </xdr:graphicFrame><xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#;

    // ── rel_id ────────────────────────────────────────────────────────────────

    #[test]
    fn finds_single_chart_ref() {
        let refs = parse_xml(ZERO_OFFSETS).unwrap();
        assert_eq!(refs.len(), 1);
        assert_eq!(refs.refs[0].rel_id, "rId1");
    }

    #[test]
    fn finds_two_chart_refs() {
        let refs = parse_xml(TWO_CHARTS).unwrap();
        assert_eq!(refs.len(), 2);
        assert_eq!(refs.refs[0].rel_id, "rId1");
        assert_eq!(refs.refs[1].rel_id, "rId2");
    }

    #[test]
    fn drawing_with_no_charts_is_empty() {
        let refs = parse_xml("<xdr:wsDr/>").unwrap();
        assert!(refs.is_empty());
    }

    // ── anchor: zero-offset fixture ───────────────────────────────────────────

    #[test]
    fn anchor_is_present() {
        let refs = parse_xml(ZERO_OFFSETS).unwrap();
        assert!(refs.refs[0].anchor.is_some());
    }

    #[test]
    fn anchor_col_start() {
        let a = parse_xml(ZERO_OFFSETS).unwrap();
        assert_eq!(a.refs[0].anchor.as_ref().unwrap().col_start, 0);
    }

    #[test]
    fn anchor_row_start() {
        let a = parse_xml(ZERO_OFFSETS).unwrap();
        assert_eq!(a.refs[0].anchor.as_ref().unwrap().row_start, 0);
    }

    #[test]
    fn anchor_col_end() {
        let a = parse_xml(ZERO_OFFSETS).unwrap();
        assert_eq!(a.refs[0].anchor.as_ref().unwrap().col_end, 8);
    }

    #[test]
    fn anchor_row_end() {
        let a = parse_xml(ZERO_OFFSETS).unwrap();
        assert_eq!(a.refs[0].anchor.as_ref().unwrap().row_end, 15);
    }

    #[test]
    fn anchor_offsets_are_zero() {
        let a = parse_xml(ZERO_OFFSETS).unwrap();
        let anch = a.refs[0].anchor.as_ref().unwrap();
        assert_eq!(anch.col_off, 0);
        assert_eq!(anch.row_off, 0);
        assert_eq!(anch.col_end_off, 0);
        assert_eq!(anch.row_end_off, 0);
    }

    #[test]
    fn anchor_col_span() {
        let a = parse_xml(ZERO_OFFSETS).unwrap();
        assert_eq!(a.refs[0].anchor.as_ref().unwrap().col_span(), 8);
    }

    #[test]
    fn anchor_row_span() {
        let a = parse_xml(ZERO_OFFSETS).unwrap();
        assert_eq!(a.refs[0].anchor.as_ref().unwrap().row_span(), 15);
    }

    // ── anchor: non-zero values ───────────────────────────────────────────────

    #[test]
    fn non_zero_col_start() {
        let a = parse_xml(NON_ZERO).unwrap();
        assert_eq!(a.refs[0].anchor.as_ref().unwrap().col_start, 1);
    }

    #[test]
    fn non_zero_col_off() {
        let a = parse_xml(NON_ZERO).unwrap();
        assert_eq!(a.refs[0].anchor.as_ref().unwrap().col_off, 12345);
    }

    #[test]
    fn non_zero_row_start() {
        let a = parse_xml(NON_ZERO).unwrap();
        assert_eq!(a.refs[0].anchor.as_ref().unwrap().row_start, 2);
    }

    #[test]
    fn non_zero_row_off() {
        let a = parse_xml(NON_ZERO).unwrap();
        assert_eq!(a.refs[0].anchor.as_ref().unwrap().row_off, 67890);
    }

    #[test]
    fn non_zero_col_end() {
        let a = parse_xml(NON_ZERO).unwrap();
        assert_eq!(a.refs[0].anchor.as_ref().unwrap().col_end, 9);
    }

    #[test]
    fn non_zero_col_end_off() {
        let a = parse_xml(NON_ZERO).unwrap();
        assert_eq!(a.refs[0].anchor.as_ref().unwrap().col_end_off, 11111);
    }

    #[test]
    fn non_zero_row_end() {
        let a = parse_xml(NON_ZERO).unwrap();
        assert_eq!(a.refs[0].anchor.as_ref().unwrap().row_end, 18);
    }

    #[test]
    fn non_zero_row_end_off() {
        let a = parse_xml(NON_ZERO).unwrap();
        assert_eq!(a.refs[0].anchor.as_ref().unwrap().row_end_off, 22222);
    }

    #[test]
    fn non_zero_rel_id() {
        let a = parse_xml(NON_ZERO).unwrap();
        assert_eq!(a.refs[0].rel_id, "rId3");
    }

    // ── two anchors stay independent ──────────────────────────────────────────

    #[test]
    fn two_anchors_col_starts_differ() {
        let refs = parse_xml(TWO_CHARTS).unwrap();
        assert_eq!(refs.refs[0].anchor.as_ref().unwrap().col_start, 0);
        assert_eq!(refs.refs[1].anchor.as_ref().unwrap().col_start, 6);
    }

    #[test]
    fn two_anchors_col_spans() {
        let refs = parse_xml(TWO_CHARTS).unwrap();
        assert_eq!(refs.refs[0].anchor.as_ref().unwrap().col_span(), 5);
        assert_eq!(refs.refs[1].anchor.as_ref().unwrap().col_span(), 6);
    }

    #[test]
    fn two_anchors_are_not_equal() {
        let refs = parse_xml(TWO_CHARTS).unwrap();
        assert_ne!(refs.refs[0].anchor, refs.refs[1].anchor);
    }

    // ── oneCellAnchor ─────────────────────────────────────────────────────────

    const ONE_CELL_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr
  xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <xdr:oneCellAnchor>
    <xdr:from>
      <xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff>
      <xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:ext cx="3000000" cy="2000000"/>
    <xdr:graphicFrame>
      <a:graphic><a:graphicData>
        <c:chart r:id="rId2"/>
      </a:graphicData></a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:oneCellAnchor>
</xdr:wsDr>"#;

    #[test]
    fn one_cell_anchor_found() {
        let refs = parse_xml(ONE_CELL_XML).unwrap();
        assert_eq!(refs.len(), 1);
    }

    #[test]
    fn one_cell_rel_id() {
        let refs = parse_xml(ONE_CELL_XML).unwrap();
        assert_eq!(refs.refs[0].rel_id, "rId2");
    }

    #[test]
    fn one_cell_anchor_present() {
        let refs = parse_xml(ONE_CELL_XML).unwrap();
        assert!(refs.refs[0].anchor.is_some());
    }

    #[test]
    fn one_cell_col_start() {
        let refs = parse_xml(ONE_CELL_XML).unwrap();
        assert_eq!(refs.refs[0].anchor.as_ref().unwrap().col_start, 5);
    }

    #[test]
    fn one_cell_row_start() {
        let refs = parse_xml(ONE_CELL_XML).unwrap();
        assert_eq!(refs.refs[0].anchor.as_ref().unwrap().row_start, 1);
    }

    #[test]
    fn one_cell_ext_cx_stored_in_col_end_off() {
        // For oneCellAnchor, ext cx is stored in col_end_off for downstream use
        let refs = parse_xml(ONE_CELL_XML).unwrap();
        assert_eq!(refs.refs[0].anchor.as_ref().unwrap().col_end_off, 3000000);
    }

    #[test]
    fn one_cell_ext_cy_stored_in_row_end_off() {
        let refs = parse_xml(ONE_CELL_XML).unwrap();
        assert_eq!(refs.refs[0].anchor.as_ref().unwrap().row_end_off, 2000000);
    }

    /// Mixed: one twoCellAnchor + one oneCellAnchor
    const MIXED_ANCHOR_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr
  xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff>
              <xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to>  <xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff>
              <xdr:row>15</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame><a:graphic><a:graphicData>
      <c:chart r:id="rId1"/>
    </a:graphicData></a:graphic></xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:oneCellAnchor>
    <xdr:from><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff>
              <xdr:row>17</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:ext cx="4572000" cy="2743200"/>
    <xdr:graphicFrame><a:graphic><a:graphicData>
      <c:chart r:id="rId2"/>
    </a:graphicData></a:graphic></xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:oneCellAnchor>
</xdr:wsDr>"#;

    #[test]
    fn mixed_two_refs_found() {
        let refs = parse_xml(MIXED_ANCHOR_XML).unwrap();
        assert_eq!(refs.len(), 2);
    }

    #[test]
    fn mixed_first_is_two_cell() {
        let refs = parse_xml(MIXED_ANCHOR_XML).unwrap();
        let a = refs.refs[0].anchor.as_ref().unwrap();
        // twoCellAnchor: col_end=8 (different from col_start=0)
        assert_eq!(a.col_start, 0);
        assert_eq!(a.col_end, 8);
    }

    #[test]
    fn mixed_second_is_one_cell() {
        let refs = parse_xml(MIXED_ANCHOR_XML).unwrap();
        let a = refs.refs[1].anchor.as_ref().unwrap();
        // oneCellAnchor: col_start=2, col_end=col_start=2
        assert_eq!(a.col_start, 2);
        assert_eq!(a.row_start, 17);
        // ext stored in offsets
        assert_eq!(a.col_end_off, 4572000);
        assert_eq!(a.row_end_off, 2743200);
    }
}
