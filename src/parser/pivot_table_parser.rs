//! Parser for `xl/pivotTables/pivotTableN.xml`.
//!
//! Extracts the pivot table name and the ordered list of pivot field indices
//! from `<pivotTableDefinition>`.  Field *names* live in the cache definition
//! (parsed separately); this module only captures the count / ordering so the
//! caller can join by position.
//!
//! ## Elements parsed
//!
//! ```xml
//! <pivotTableDefinition
//!     xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
//!     name="PivotTable1"
//!     cacheId="1"
//!     …>
//!   <location ref="A1:E6" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
//!   <pivotFields count="4">
//!     <pivotField axis="axisRow" showAll="0"/>
//!     <pivotField axis="axisCol" showAll="0"/>
//!     <pivotField dataField="1"/>
//!     <pivotField showAll="0"/>
//!   </pivotFields>
//!   …
//! </pivotTableDefinition>
//! ```
//!
//! We capture `name` and the count of `<pivotField>` children (which equals the
//! number of columns in the source range and maps 1:1 to `<cacheField>` entries).

use anyhow::{Context, Result};
use quick_xml::{events::Event, Reader};

use crate::archive::zip_reader::{read_entry_bytes, XlsxArchive};

// ── Public types ──────────────────────────────────────────────────────────────

/// Partial result from parsing `pivotTableDefinition.xml`.
/// Field names are filled in later by the caller from the cache definition.
#[derive(Debug, Default)]
pub struct PivotTableRaw {
    /// Value of the `name` attribute on `<pivotTableDefinition>`.
    pub name: String,
    /// Number of `<pivotField>` children inside `<pivotFields>`.
    /// Equals the number of columns / cache fields in the source range.
    pub field_count: usize,
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

    loop {
        match reader.read_event()? {
            Event::Start(ref e) | Event::Empty(ref e) => {
                let local = e.local_name();
                match local.as_ref() {
                    b"pivotTableDefinition" => {
                        let dec = reader.decoder();
                        for attr in e.attributes() {
                            let attr = attr.context("Malformed attr in pivotTableDefinition")?;
                            if attr.key.local_name().as_ref() == b"name" {
                                out.name = attr.decode_and_unescape_value(dec)?.into_owned();
                            }
                        }
                    }
                    b"pivotFields" => {
                        in_pivot_fields = true;
                        // `count` attribute is informational; we count children instead.
                    }
                    b"pivotField" if in_pivot_fields => {
                        out.field_count += 1;
                    }
                    _ => {}
                }
            }
            Event::End(ref e) => {
                if e.local_name().as_ref() == b"pivotFields" {
                    in_pivot_fields = false;
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(out)
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
    applyNumberFormats="0"
    applyBorderFormats="0"
    applyFontFormats="0"
    applyPatternFormats="0"
    applyAlignmentFormats="0"
    applyWidthHeightFormats="1"
    dataCaption="Values"
    updatedVersion="4"
    showMemberPropertyTips="0"
    useAutoFormatting="1"
    itemPrintTitles="1"
    createdVersion="4"
    indent="0"
    compact="0"
    compactData="0">
  <location ref="A1:E6" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
  <pivotFields count="4">
    <pivotField axis="axisRow" showAll="0">
      <items count="5">
        <item t="default"/>
      </items>
    </pivotField>
    <pivotField showAll="0">
      <items count="4">
        <item t="default"/>
      </items>
    </pivotField>
    <pivotField axis="axisCol" showAll="0">
      <items count="3">
        <item t="default"/>
      </items>
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
    <dataField name="Sum of Sales" fld="3" baseField="0" baseItem="0"/>
  </dataFields>
</pivotTableDefinition>"#;

    const MINIMAL_PIVOT_TABLE_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition
    xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    name="SalesPivot"
    cacheId="2">
  <location ref="A1:B4" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
  <pivotFields count="2">
    <pivotField axis="axisRow"/>
    <pivotField dataField="1"/>
  </pivotFields>
</pivotTableDefinition>"#;

    const NO_FIELDS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition
    xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    name="EmptyPivot"
    cacheId="3">
  <location ref="A1:A1" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
  <pivotFields count="0"/>
</pivotTableDefinition>"#;

    const NO_NAME_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition
    xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    cacheId="4">
  <pivotFields count="1">
    <pivotField/>
  </pivotFields>
</pivotTableDefinition>"#;

    // ── name extraction ───────────────────────────────────────────────────────

    #[test]
    fn extracts_name() {
        let r = parse_bytes(PIVOT_TABLE_XML.as_bytes()).unwrap();
        assert_eq!(r.name, "PivotTable1");
    }

    #[test]
    fn extracts_minimal_name() {
        let r = parse_bytes(MINIMAL_PIVOT_TABLE_XML.as_bytes()).unwrap();
        assert_eq!(r.name, "SalesPivot");
    }

    #[test]
    fn missing_name_attr_is_empty_string() {
        let r = parse_bytes(NO_NAME_XML.as_bytes()).unwrap();
        assert_eq!(r.name, "");
    }

    // ── field count ───────────────────────────────────────────────────────────

    #[test]
    fn counts_four_pivot_fields() {
        let r = parse_bytes(PIVOT_TABLE_XML.as_bytes()).unwrap();
        assert_eq!(r.field_count, 4);
    }

    #[test]
    fn counts_two_pivot_fields() {
        let r = parse_bytes(MINIMAL_PIVOT_TABLE_XML.as_bytes()).unwrap();
        assert_eq!(r.field_count, 2);
    }

    #[test]
    fn empty_pivot_fields_count_zero() {
        let r = parse_bytes(NO_FIELDS_XML.as_bytes()).unwrap();
        assert_eq!(r.field_count, 0);
    }

    #[test]
    fn field_in_no_name_xml_counted() {
        let r = parse_bytes(NO_NAME_XML.as_bytes()).unwrap();
        assert_eq!(r.field_count, 1);
    }

    // ── does not count items/fields outside pivotFields ───────────────────────

    #[test]
    fn items_inside_pivot_field_not_counted_as_fields() {
        // The <items><item/></items> inside each <pivotField> must not
        // be mistaken for additional pivot fields.
        let r = parse_bytes(PIVOT_TABLE_XML.as_bytes()).unwrap();
        assert_eq!(
            r.field_count, 4,
            "inner <item> elements must not inflate field_count"
        );
    }

    // ── empty/malformed inputs ────────────────────────────────────────────────

    #[test]
    fn empty_xml_returns_defaults() {
        let xml = b"<?xml version=\"1.0\"?><root/>";
        let r = parse_bytes(xml).unwrap();
        assert_eq!(r.name, "");
        assert_eq!(r.field_count, 0);
    }
}
