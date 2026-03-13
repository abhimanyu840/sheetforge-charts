//! Parser for `xl/drawings/drawingN.xml`.
//!
//! Extracts every `<xdr:twoCellAnchor>` and pairs it with the
//! `<c:chart r:id="…"/>` reference nested inside it so that callers
//! can attach worksheet-position data to each [`Chart`] skeleton.
//!
//! ## XML structure
//!
//! ```xml
//! <xdr:wsDr …>
//!   <xdr:twoCellAnchor>
//!     <xdr:from>
//!       <xdr:col>0</xdr:col>   <xdr:colOff>0</xdr:colOff>
//!       <xdr:row>0</xdr:row>   <xdr:rowOff>0</xdr:rowOff>
//!     </xdr:from>
//!     <xdr:to>
//!       <xdr:col>8</xdr:col>   <xdr:colOff>0</xdr:colOff>
//!       <xdr:row>15</xdr:row>  <xdr:rowOff>0</xdr:rowOff>
//!     </xdr:to>
//!     <xdr:graphicFrame>
//!       …
//!       <c:chart r:id="rId1"/>
//!       …
//!     </xdr:graphicFrame>
//!   </xdr:twoCellAnchor>
//! </xdr:wsDr>
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
                    "twoCellAnchor" => st.begin_anchor(),
                    "from" => st.corner = Corner::From,
                    "to" => st.corner = Corner::To,
                    "col" => st.pending = Some(Field::Col),
                    "colOff" => st.pending = Some(Field::ColOff),
                    "row" => st.pending = Some(Field::Row),
                    "rowOff" => st.pending = Some(Field::RowOff),
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
                                anchor: Some(st.build()),
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

#[derive(Debug, Clone, Copy)]
enum Field {
    Col,
    ColOff,
    Row,
    RowOff,
}

#[derive(Debug, Default)]
struct State {
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
}

impl State {
    fn begin_anchor(&mut self) {
        self.reset();
    }

    fn reset(&mut self) {
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
        // pending_rel_id is consumed at </twoCellAnchor> close, not here
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

    fn build(&self) -> ChartAnchor {
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
}
