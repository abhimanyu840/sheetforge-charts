//! Parser for `xl/pivotTables/pivotTableN.xml`.
//!
//! ## Phase 11 (original)
//! Extracts `name` and `field_count` from `<pivotTableDefinition>`.
//!
//! ## Phase 12 (extended)
//! Also extracts the structural layout needed for pivot data aggregation:
//!
//! * `row_field_idxs` — which cache-field columns form the category (row) axis.
//! * `col_field_idxs` — which cache-field columns form the series-name (column) axis.
//! * `data_fields`    — which cache-field columns are the aggregated values and
//!   what their display names and subtotal functions are.
//!
//! ```xml
//! <pivotTableDefinition name="PivotTable1" cacheId="1">
//!   <pivotFields count="4">
//!     <pivotField axis="axisRow" showAll="0"/>   <!-- field 0 → row axis -->
//!     <pivotField showAll="0"/>                   <!-- field 1 → neither axis -->
//!     <pivotField axis="axisCol" showAll="0"/>   <!-- field 2 → col axis -->
//!     <pivotField dataField="1" showAll="0"/>    <!-- field 3 → data field -->
//!   </pivotFields>
//!   <rowFields count="1"><field x="0"/></rowFields>
//!   <colFields count="1"><field x="2"/></colFields>
//!   <dataFields count="1">
//!     <dataField name="Sum of Sales" fld="3" subtotal="sum"/>
//!   </dataFields>
//! </pivotTableDefinition>
//! ```
//!
//! The `rowFields` / `colFields` `<field x="N"/>` elements are the authoritative
//! source for which cache-field indices are on each axis (they may be a subset of
//! all fields in the cache, and their order in the XML determines axis order).

use anyhow::{Context, Result};
use quick_xml::{events::Event, Reader};

use crate::archive::zip_reader::{read_entry_bytes, XlsxArchive};

// ── Public types ──────────────────────────────────────────────────────────────

/// One entry from `<dataFields><dataField …/></dataFields>`.
#[derive(Debug, Clone, Default)]
pub struct DataFieldDef {
    /// Display name (e.g. `"Sum of Sales"`), from `name` attribute.
    pub name: String,
    /// Zero-based index of the backing `<cacheField>` (`fld` attribute).
    pub field_idx: usize,
    /// Aggregation function name (`subtotal` attribute), e.g. `"sum"`, `"count"`,
    /// `"average"`.  Defaults to `"sum"` when the attribute is absent.
    pub subtotal: String,
}

/// Full structural layout extracted from `pivotTableDefinition.xml`.
#[derive(Debug, Default)]
pub struct PivotTableRaw {
    /// Value of the `name` attribute on `<pivotTableDefinition>`.
    pub name: String,
    /// Total number of `<pivotField>` children (= number of cache fields used).
    pub field_count: usize,
    /// Cache-field indices that form the row (category) axis, in axis order.
    /// Sourced from `<rowFields><field x="N"/></rowFields>`.
    pub row_field_idxs: Vec<usize>,
    /// Cache-field indices that form the column (series-name) axis, in axis order.
    /// Sourced from `<colFields><field x="N"/></colFields>`.
    /// Empty when the pivot has no column field (single-series layout).
    pub col_field_idxs: Vec<usize>,
    /// Data field definitions from `<dataFields>`.
    pub data_fields: Vec<DataFieldDef>,
}

// ── Entry points ──────────────────────────────────────────────────────────────

pub fn parse(archive: &mut XlsxArchive, path: &str) -> Result<PivotTableRaw> {
    let bytes = read_entry_bytes(archive, path)
        .with_context(|| format!("Cannot read pivot table: {path}"))?;
    parse_bytes(&bytes).with_context(|| format!("Failed to parse pivot table XML: {path}"))
}

pub fn parse_bytes(bytes: &[u8]) -> Result<PivotTableRaw> {
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().trim_text(true);

    let mut out = PivotTableRaw::default();
    let mut in_pivot_fields = false;
    let mut in_row_fields = false;
    let mut in_col_fields = false;
    let mut in_data_fields = false;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) | Event::Empty(ref e) => {
                let local = e.local_name();
                let dec = reader.decoder();
                match local.as_ref() {
                    b"pivotTableDefinition" => {
                        for attr in e.attributes() {
                            let attr = attr.context("Malformed attr in pivotTableDefinition")?;
                            if attr.key.local_name().as_ref() == b"name" {
                                out.name = attr.decode_and_unescape_value(dec)?.into_owned();
                            }
                        }
                    }

                    // ── pivotFields section ───────────────────────────────────
                    b"pivotFields" => {
                        in_pivot_fields = true;
                    }
                    b"pivotField" if in_pivot_fields => {
                        out.field_count += 1;
                    }

                    // ── rowFields section ─────────────────────────────────────
                    b"rowFields" => {
                        in_row_fields = true;
                    }
                    b"field" if in_row_fields => {
                        if let Some(x) = read_x_attr(e, dec)? {
                            // x == -2 is the special "values" pseudo-field; skip it
                            if x != usize::MAX {
                                out.row_field_idxs.push(x);
                            }
                        }
                    }

                    // ── colFields section ─────────────────────────────────────
                    b"colFields" => {
                        in_col_fields = true;
                    }
                    b"field" if in_col_fields => {
                        if let Some(x) = read_x_attr(e, dec)? {
                            if x != usize::MAX {
                                out.col_field_idxs.push(x);
                            }
                        }
                    }

                    // ── dataFields section ────────────────────────────────────
                    b"dataFields" => {
                        in_data_fields = true;
                    }
                    b"dataField" if in_data_fields => {
                        let mut def = DataFieldDef {
                            subtotal: "sum".to_owned(),
                            ..Default::default()
                        };
                        for attr in e.attributes() {
                            let attr = attr.context("Malformed attr in dataField")?;
                            match attr.key.local_name().as_ref() {
                                b"name" => {
                                    def.name = attr.decode_and_unescape_value(dec)?.into_owned()
                                }
                                b"fld" => {
                                    def.field_idx =
                                        attr.decode_and_unescape_value(dec)?.parse().unwrap_or(0)
                                }
                                b"subtotal" => {
                                    def.subtotal = attr.decode_and_unescape_value(dec)?.into_owned()
                                }
                                _ => {}
                            }
                        }
                        out.data_fields.push(def);
                    }

                    _ => {}
                }
            }

            Event::End(ref e) => match e.local_name().as_ref() {
                b"pivotFields" => {
                    in_pivot_fields = false;
                }
                b"rowFields" => {
                    in_row_fields = false;
                }
                b"colFields" => {
                    in_col_fields = false;
                }
                b"dataFields" => {
                    in_data_fields = false;
                }
                _ => {}
            },

            Event::Eof => break,
            _ => {}
        }
    }

    Ok(out)
}

/// Read the `x` attribute of a `<field>` element.
/// Returns `None` if absent, `Some(usize::MAX)` for the special `-2` values field.
fn read_x_attr(
    e: &quick_xml::events::BytesStart<'_>,
    dec: quick_xml::Decoder,
) -> Result<Option<usize>> {
    for attr in e.attributes() {
        let attr = attr.context("Malformed attr in field")?;
        if attr.key.local_name().as_ref() == b"x" {
            let raw = attr.decode_and_unescape_value(dec)?;
            let val: i64 = raw.parse().unwrap_or(0);
            // -2 = special "values" pseudo-field; map to sentinel
            return Ok(Some(if val < 0 { usize::MAX } else { val as usize }));
        }
    }
    Ok(None)
}

// ── Tests ─────────────────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;

    // ── XML fixtures ──────────────────────────────────────────────────────────

    const PIVOT_TABLE_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition
    xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    name="PivotTable1"
    cacheId="1"
    dataCaption="Values"
    updatedVersion="4"
    createdVersion="4"
    indent="0"
    compact="0"
    compactData="0">
  <location ref="A1:E6" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
  <pivotFields count="4">
    <pivotField axis="axisRow" showAll="0">
      <items count="5"><item t="default"/></items>
    </pivotField>
    <pivotField showAll="0">
      <items count="4"><item t="default"/></items>
    </pivotField>
    <pivotField axis="axisCol" showAll="0">
      <items count="3"><item t="default"/></items>
    </pivotField>
    <pivotField dataField="1" showAll="0"/>
  </pivotFields>
  <rowFields count="1">
    <field x="0"/>
  </rowFields>
  <colFields count="1">
    <field x="2"/>
  </colFields>
  <dataFields count="1">
    <dataField name="Sum of Sales" fld="3" subtotal="sum"/>
  </dataFields>
</pivotTableDefinition>"#;

    /// Pivot with only row fields and data fields — no column axis.
    const NO_COL_FIELDS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition
    xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    name="SalesPivot"
    cacheId="2">
  <location ref="A1:B5" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
  <pivotFields count="2">
    <pivotField axis="axisRow"/>
    <pivotField dataField="1"/>
  </pivotFields>
  <rowFields count="1">
    <field x="0"/>
  </rowFields>
  <dataFields count="1">
    <dataField name="Sum of Revenue" fld="1" subtotal="sum"/>
  </dataFields>
</pivotTableDefinition>"#;

    /// Pivot with multiple row fields.
    const MULTI_ROW_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition
    xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    name="Multi"
    cacheId="3">
  <pivotFields count="3">
    <pivotField axis="axisRow"/>
    <pivotField axis="axisRow"/>
    <pivotField dataField="1"/>
  </pivotFields>
  <rowFields count="2">
    <field x="0"/>
    <field x="1"/>
  </rowFields>
  <dataFields count="1">
    <dataField name="Sum of Value" fld="2" subtotal="sum"/>
  </dataFields>
</pivotTableDefinition>"#;

    /// colFields containing the special -2 "values" pseudo-field.
    const VALUES_PSEUDO_FIELD_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition
    xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    name="MultiData"
    cacheId="4">
  <pivotFields count="3">
    <pivotField axis="axisRow"/>
    <pivotField dataField="1"/>
    <pivotField dataField="1"/>
  </pivotFields>
  <rowFields count="1"><field x="0"/></rowFields>
  <colFields count="1"><field x="-2"/></colFields>
  <dataFields count="2">
    <dataField name="Sum of A" fld="1" subtotal="sum"/>
    <dataField name="Sum of B" fld="2" subtotal="sum"/>
  </dataFields>
</pivotTableDefinition>"#;

    const NO_NAME_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" cacheId="4">
  <pivotFields count="1"><pivotField/></pivotFields>
</pivotTableDefinition>"#;

    // ── name ─────────────────────────────────────────────────────────────────

    #[test]
    fn extracts_name() {
        assert_eq!(
            parse_bytes(PIVOT_TABLE_XML.as_bytes()).unwrap().name,
            "PivotTable1"
        );
    }

    #[test]
    fn extracts_sales_pivot_name() {
        assert_eq!(
            parse_bytes(NO_COL_FIELDS_XML.as_bytes()).unwrap().name,
            "SalesPivot"
        );
    }

    #[test]
    fn missing_name_is_empty() {
        assert_eq!(parse_bytes(NO_NAME_XML.as_bytes()).unwrap().name, "");
    }

    // ── field_count ───────────────────────────────────────────────────────────

    #[test]
    fn counts_four_pivot_fields() {
        assert_eq!(
            parse_bytes(PIVOT_TABLE_XML.as_bytes()).unwrap().field_count,
            4
        );
    }

    #[test]
    fn counts_two_pivot_fields() {
        assert_eq!(
            parse_bytes(NO_COL_FIELDS_XML.as_bytes())
                .unwrap()
                .field_count,
            2
        );
    }

    #[test]
    fn items_not_counted_as_fields() {
        // <items><item/></items> inside pivotField must not inflate field_count
        assert_eq!(
            parse_bytes(PIVOT_TABLE_XML.as_bytes()).unwrap().field_count,
            4
        );
    }

    // ── row_field_idxs ────────────────────────────────────────────────────────

    #[test]
    fn row_field_idx_single() {
        let r = parse_bytes(PIVOT_TABLE_XML.as_bytes()).unwrap();
        assert_eq!(r.row_field_idxs, vec![0]);
    }

    #[test]
    fn row_field_idx_no_col() {
        let r = parse_bytes(NO_COL_FIELDS_XML.as_bytes()).unwrap();
        assert_eq!(r.row_field_idxs, vec![0]);
    }

    #[test]
    fn row_field_idxs_multiple() {
        let r = parse_bytes(MULTI_ROW_XML.as_bytes()).unwrap();
        assert_eq!(r.row_field_idxs, vec![0, 1]);
    }

    // ── col_field_idxs ────────────────────────────────────────────────────────

    #[test]
    fn col_field_idx_single() {
        let r = parse_bytes(PIVOT_TABLE_XML.as_bytes()).unwrap();
        assert_eq!(r.col_field_idxs, vec![2]);
    }

    #[test]
    fn col_field_idxs_empty_when_no_col_fields() {
        let r = parse_bytes(NO_COL_FIELDS_XML.as_bytes()).unwrap();
        assert!(r.col_field_idxs.is_empty());
    }

    #[test]
    fn col_field_pseudo_values_field_excluded() {
        // -2 is the "values" pseudo-field, must be filtered out
        let r = parse_bytes(VALUES_PSEUDO_FIELD_XML.as_bytes()).unwrap();
        assert!(
            r.col_field_idxs.is_empty(),
            "the -2 pseudo-field must not appear in col_field_idxs"
        );
    }

    // ── data_fields ───────────────────────────────────────────────────────────

    #[test]
    fn data_field_count() {
        let r = parse_bytes(PIVOT_TABLE_XML.as_bytes()).unwrap();
        assert_eq!(r.data_fields.len(), 1);
    }

    #[test]
    fn data_field_name() {
        let r = parse_bytes(PIVOT_TABLE_XML.as_bytes()).unwrap();
        assert_eq!(r.data_fields[0].name, "Sum of Sales");
    }

    #[test]
    fn data_field_fld_index() {
        let r = parse_bytes(PIVOT_TABLE_XML.as_bytes()).unwrap();
        assert_eq!(r.data_fields[0].field_idx, 3);
    }

    #[test]
    fn data_field_subtotal_sum() {
        let r = parse_bytes(PIVOT_TABLE_XML.as_bytes()).unwrap();
        assert_eq!(r.data_fields[0].subtotal, "sum");
    }

    #[test]
    fn data_field_subtotal_default_is_sum() {
        // When subtotal attribute is absent, default is "sum"
        let xml = r#"<?xml version="1.0"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="T" cacheId="1">
  <dataFields count="1">
    <dataField name="Values" fld="1"/>
  </dataFields>
</pivotTableDefinition>"#;
        let r = parse_bytes(xml.as_bytes()).unwrap();
        assert_eq!(r.data_fields[0].subtotal, "sum");
    }

    #[test]
    fn multi_data_fields() {
        let r = parse_bytes(VALUES_PSEUDO_FIELD_XML.as_bytes()).unwrap();
        assert_eq!(r.data_fields.len(), 2);
        assert_eq!(r.data_fields[0].name, "Sum of A");
        assert_eq!(r.data_fields[1].name, "Sum of B");
        assert_eq!(r.data_fields[0].field_idx, 1);
        assert_eq!(r.data_fields[1].field_idx, 2);
    }

    #[test]
    fn no_data_fields_returns_empty() {
        let r = parse_bytes(NO_NAME_XML.as_bytes()).unwrap();
        assert!(r.data_fields.is_empty());
    }
}
