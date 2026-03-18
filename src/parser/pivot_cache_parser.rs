//! Parser for `xl/pivotCache/pivotCacheDefinitionN.xml`.
//!
//! ## Phase 11 (original)
//! Extracts source location and field names.
//!
//! ## Phase 12 (extended)
//! Also extracts the `sharedItems` lookup tables per field.
//! These are needed to decode `<x v="N"/>` index references in
//! `pivotCacheRecords.xml`.
//!
//! ```xml
//! <cacheField name="Region" numFmtId="0">
//!   <sharedItems count="3">
//!     <s v="North"/>   <!-- index 0 -->
//!     <s v="South"/>   <!-- index 1 -->
//!     <s v="East"/>    <!-- index 2 -->
//!   </sharedItems>
//! </cacheField>
//! <cacheField name="Sales" numFmtId="0">
//!   <!-- no sharedItems: numeric field — values appear inline in records -->
//! </cacheField>
//! ```
//!
//! A field with no `<sharedItems>` children (or with `containsNumber="1"` and
//! no `<s>` / `<b>` / `<e>` children) is a direct-value field.  Its column in
//! a record row will contain `<n v="…"/>` or `<s v="…"/>` directly.

use anyhow::{Context, Result};
use quick_xml::{events::Event, Reader};

use crate::archive::zip_reader::{read_entry_bytes, XlsxArchive};

// ── Public types ──────────────────────────────────────────────────────────────

/// Result of parsing a single `pivotCacheDefinition.xml`.
#[derive(Debug, Default)]
pub struct PivotCacheRaw {
    /// Worksheet name from `<worksheetSource sheet="…"/>`.
    pub source_sheet: Option<String>,
    /// Cell range from `<worksheetSource ref="…"/>` in A1 notation.
    pub source_range: Option<String>,
    /// Field names in order from `<cacheField name="…"/>`.
    pub field_names: Vec<String>,
    /// Shared-items lookup tables, one `Vec<String>` per field.
    /// `shared_items[i][j]` is the string value for index `j` in field `i`.
    /// Empty inner `Vec` means the field has no shared items (direct-value field).
    pub shared_items: Vec<Vec<String>>,
}

// ── Entry points ──────────────────────────────────────────────────────────────

pub fn parse(archive: &mut XlsxArchive, path: &str) -> Result<PivotCacheRaw> {
    let bytes = read_entry_bytes(archive, path)
        .with_context(|| format!("Cannot read pivot cache: {path}"))?;
    parse_bytes(&bytes).with_context(|| format!("Failed to parse pivot cache XML: {path}"))
}

pub fn parse_bytes(bytes: &[u8]) -> Result<PivotCacheRaw> {
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().trim_text(true);

    let mut out = PivotCacheRaw::default();
    let mut in_cache_source = false;
    let mut in_shared_items = false;
    // Index of the cacheField currently being parsed (-1 = none)
    let mut cur_field: isize = -1;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) | Event::Empty(ref e) => {
                let local = e.local_name();
                let dec = reader.decoder();
                match local.as_ref() {
                    b"cacheSource" => {
                        in_cache_source = true;
                    }

                    b"worksheetSource" if in_cache_source => {
                        for attr in e.attributes() {
                            let attr = attr.context("Malformed attr in worksheetSource")?;
                            match attr.key.local_name().as_ref() {
                                b"ref" => {
                                    out.source_range =
                                        Some(attr.decode_and_unescape_value(dec)?.into_owned())
                                }
                                b"sheet" => {
                                    out.source_sheet =
                                        Some(attr.decode_and_unescape_value(dec)?.into_owned())
                                }
                                _ => {}
                            }
                        }
                    }

                    b"cacheField" => {
                        let mut field_name = String::new();
                        for attr in e.attributes() {
                            let attr = attr.context("Malformed attr in cacheField")?;
                            if attr.key.local_name().as_ref() == b"name" {
                                field_name = attr.decode_and_unescape_value(dec)?.into_owned();
                            }
                        }
                        out.field_names.push(field_name);
                        out.shared_items.push(Vec::new());
                        cur_field = (out.field_names.len() as isize) - 1;
                    }

                    b"sharedItems" if cur_field >= 0 => {
                        in_shared_items = true;
                    }

                    // String shared item  <s v="North"/>
                    b"s" if in_shared_items && cur_field >= 0 => {
                        let mut val = String::new();
                        for attr in e.attributes() {
                            let attr = attr.context("Malformed attr in s")?;
                            if attr.key.local_name().as_ref() == b"v" {
                                val = attr.decode_and_unescape_value(dec)?.into_owned();
                            }
                        }
                        out.shared_items[cur_field as usize].push(val);
                    }

                    // Numeric shared item  <n v="42"/>
                    b"n" if in_shared_items && cur_field >= 0 => {
                        let mut val = String::new();
                        for attr in e.attributes() {
                            let attr = attr.context("Malformed attr in n (shared)")?;
                            if attr.key.local_name().as_ref() == b"v" {
                                val = attr.decode_and_unescape_value(dec)?.into_owned();
                            }
                        }
                        out.shared_items[cur_field as usize].push(val);
                    }

                    // Boolean shared item  <b v="1"/>
                    b"b" if in_shared_items && cur_field >= 0 => {
                        let mut val = String::new();
                        for attr in e.attributes() {
                            let attr = attr.context("Malformed attr in b (shared)")?;
                            if attr.key.local_name().as_ref() == b"v" {
                                val = attr.decode_and_unescape_value(dec)?.into_owned();
                            }
                        }
                        out.shared_items[cur_field as usize].push(val);
                    }

                    // Error shared item  <e v="#N/A"/>
                    b"e" if in_shared_items && cur_field >= 0 => {
                        let mut val = String::new();
                        for attr in e.attributes() {
                            let attr = attr.context("Malformed attr in e (shared)")?;
                            if attr.key.local_name().as_ref() == b"v" {
                                val = attr.decode_and_unescape_value(dec)?.into_owned();
                            }
                        }
                        out.shared_items[cur_field as usize].push(val);
                    }

                    // Missing shared item  <m/>  — push empty string placeholder
                    b"m" if in_shared_items && cur_field >= 0 => {
                        out.shared_items[cur_field as usize].push(String::new());
                    }

                    _ => {}
                }
            }

            Event::End(ref e) => match e.local_name().as_ref() {
                b"cacheSource" => {
                    in_cache_source = false;
                }
                b"sharedItems" => {
                    in_shared_items = false;
                }
                b"cacheField" => {
                    cur_field = -1;
                }
                _ => {}
            },

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

    const FULL_CACHE_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition
    xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    r:id="rId1" refreshedBy="Excel" refreshedDate="44927.5"
    createdVersion="4" refreshedVersion="4" minRefreshableVersion="3" recordCount="100">
  <cacheSource type="worksheet">
    <worksheetSource ref="A1:D101" sheet="SourceData"/>
  </cacheSource>
  <cacheFields count="4">
    <cacheField name="Region" numFmtId="0">
      <sharedItems count="3">
        <s v="North"/>
        <s v="South"/>
        <s v="East"/>
      </sharedItems>
    </cacheField>
    <cacheField name="Product" numFmtId="0">
      <sharedItems count="2">
        <s v="Widget"/>
        <s v="Gadget"/>
      </sharedItems>
    </cacheField>
    <cacheField name="Category" numFmtId="0">
      <sharedItems count="2">
        <s v="Electronics"/>
        <s v="Hardware"/>
      </sharedItems>
    </cacheField>
    <cacheField name="Sales" numFmtId="0">
      <sharedItems containsNumber="1" minValue="0" maxValue="9999"/>
    </cacheField>
  </cacheFields>
</pivotCacheDefinition>"#;

    const NO_SHARED_ITEMS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cacheSource type="worksheet">
    <worksheetSource ref="A1:B10" sheet="Data"/>
  </cacheSource>
  <cacheFields count="2">
    <cacheField name="Label" numFmtId="0"/>
    <cacheField name="Value" numFmtId="0"/>
  </cacheFields>
</pivotCacheDefinition>"#;

    // ── existing Phase 11 tests (unchanged) ──────────────────────────────────

    #[test]
    fn extracts_source_sheet() {
        assert_eq!(
            parse_bytes(FULL_CACHE_XML.as_bytes())
                .unwrap()
                .source_sheet
                .as_deref(),
            Some("SourceData")
        );
    }
    #[test]
    fn extracts_source_range() {
        assert_eq!(
            parse_bytes(FULL_CACHE_XML.as_bytes())
                .unwrap()
                .source_range
                .as_deref(),
            Some("A1:D101")
        );
    }
    #[test]
    fn extracts_four_field_names() {
        assert_eq!(
            parse_bytes(FULL_CACHE_XML.as_bytes())
                .unwrap()
                .field_names
                .len(),
            4
        );
    }
    #[test]
    fn field_names_in_order() {
        let r = parse_bytes(FULL_CACHE_XML.as_bytes()).unwrap();
        assert_eq!(
            r.field_names,
            vec!["Region", "Product", "Category", "Sales"]
        );
    }
    #[test]
    fn missing_sheet_attr_is_none() {
        let xml = r#"<?xml version="1.0"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cacheSource type="worksheet"><worksheetSource ref="B2:F50"/></cacheSource>
  <cacheFields count="0"/>
</pivotCacheDefinition>"#;
        assert!(parse_bytes(xml.as_bytes()).unwrap().source_sheet.is_none());
    }
    #[test]
    fn worksheet_source_outside_cache_source_ignored() {
        let xml = r#"<?xml version="1.0"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <worksheetSource ref="STRAY" sheet="Stray"/>
  <cacheSource type="worksheet"><worksheetSource ref="A1:B2" sheet="Real"/></cacheSource>
  <cacheFields count="0"/>
</pivotCacheDefinition>"#;
        let r = parse_bytes(xml.as_bytes()).unwrap();
        assert_eq!(r.source_sheet.as_deref(), Some("Real"));
    }

    // ── Phase 12: shared_items ────────────────────────────────────────────────

    #[test]
    fn shared_items_vec_count_matches_fields() {
        let r = parse_bytes(FULL_CACHE_XML.as_bytes()).unwrap();
        assert_eq!(
            r.shared_items.len(),
            4,
            "one shared_items Vec per cacheField"
        );
    }

    #[test]
    fn region_shared_items_correct() {
        let r = parse_bytes(FULL_CACHE_XML.as_bytes()).unwrap();
        assert_eq!(r.shared_items[0], vec!["North", "South", "East"]);
    }

    #[test]
    fn product_shared_items_correct() {
        let r = parse_bytes(FULL_CACHE_XML.as_bytes()).unwrap();
        assert_eq!(r.shared_items[1], vec!["Widget", "Gadget"]);
    }

    #[test]
    fn numeric_field_shared_items_empty() {
        // Sales field has <sharedItems containsNumber="1"/> but no <s> children
        let r = parse_bytes(FULL_CACHE_XML.as_bytes()).unwrap();
        assert!(
            r.shared_items[3].is_empty(),
            "numeric-only field should have empty shared_items"
        );
    }

    #[test]
    fn no_shared_items_element_means_empty_vec() {
        let r = parse_bytes(NO_SHARED_ITEMS_XML.as_bytes()).unwrap();
        assert_eq!(r.shared_items.len(), 2);
        assert!(r.shared_items[0].is_empty());
        assert!(r.shared_items[1].is_empty());
    }

    #[test]
    fn shared_items_index_zero_is_first_item() {
        let r = parse_bytes(FULL_CACHE_XML.as_bytes()).unwrap();
        assert_eq!(r.shared_items[0][0], "North");
    }

    #[test]
    fn shared_items_index_two_is_third_item() {
        let r = parse_bytes(FULL_CACHE_XML.as_bytes()).unwrap();
        assert_eq!(r.shared_items[0][2], "East");
    }
}
