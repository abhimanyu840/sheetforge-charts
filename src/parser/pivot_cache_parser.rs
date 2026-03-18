//! Parser for `xl/pivotCache/pivotCacheDefinitionN.xml`.
//!
//! Extracts the data-source location (worksheet + cell range) and the ordered
//! list of field names from `<cacheFields>`.
//!
//! ## Elements parsed
//!
//! ```xml
//! <pivotCacheDefinition
//!     xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
//!     xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
//!     r:id="rId1" refreshedBy="…" refreshedDate="…">
//!   <cacheSource type="worksheet">
//!     <worksheetSource ref="A1:D101" sheet="SourceData"/>
//!   </cacheSource>
//!   <cacheFields count="4">
//!     <cacheField name="Region"   numFmtId="0"/>
//!     <cacheField name="Product"  numFmtId="0"/>
//!     <cacheField name="Category" numFmtId="0"/>
//!     <cacheField name="Sales"    numFmtId="0"/>
//!   </cacheFields>
//! </pivotCacheDefinition>
//! ```
//!
//! We capture:
//! * `worksheetSource ref` → [`PivotCacheRaw::source_range`]
//! * `worksheetSource sheet` → [`PivotCacheRaw::source_sheet`]
//! * each `<cacheField name="…"/>` in order → [`PivotCacheRaw::field_names`]

use anyhow::{Context, Result};
use quick_xml::{events::Event, Reader};

use crate::archive::zip_reader::{read_entry_bytes, XlsxArchive};

// ── Public types ──────────────────────────────────────────────────────────────

/// Result of parsing a single `pivotCacheDefinition.xml`.
#[derive(Debug, Default)]
pub struct PivotCacheRaw {
    /// Worksheet name from `<worksheetSource sheet="…"/>`.
    /// `None` if absent (non-worksheet sources: OLAP, scenario, etc.).
    pub source_sheet: Option<String>,

    /// Cell range from `<worksheetSource ref="…"/>` in A1 notation.
    /// `None` if absent (some sources use a named table reference instead).
    pub source_range: Option<String>,

    /// Field names in order from `<cacheField name="…"/>`.
    /// Maps positionally 1:1 to `<pivotField>` entries in the pivot table.
    pub field_names: Vec<String>,
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
                                        Some(attr.decode_and_unescape_value(dec)?.into_owned());
                                }
                                b"sheet" => {
                                    out.source_sheet =
                                        Some(attr.decode_and_unescape_value(dec)?.into_owned());
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
                    }
                    _ => {}
                }
            }
            Event::End(ref e) => {
                if e.local_name().as_ref() == b"cacheSource" {
                    in_cache_source = false;
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

    const FULL_CACHE_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition
    xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    r:id="rId1"
    refreshedBy="Excel"
    refreshedDate="44927.5"
    createdVersion="4"
    refreshedVersion="4"
    minRefreshableVersion="3"
    recordCount="100">
  <cacheSource type="worksheet">
    <worksheetSource ref="A1:D101" sheet="SourceData"/>
  </cacheSource>
  <cacheFields count="4">
    <cacheField name="Region"   numFmtId="0"><sharedItems count="4"/></cacheField>
    <cacheField name="Product"  numFmtId="0"><sharedItems count="5"/></cacheField>
    <cacheField name="Category" numFmtId="0"><sharedItems count="3"/></cacheField>
    <cacheField name="Sales"    numFmtId="0"><sharedItems containsNumber="1" minValue="0" maxValue="9999"/></cacheField>
  </cacheFields>
</pivotCacheDefinition>"#;

    const NO_SHEET_ATTR_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cacheSource type="worksheet">
    <worksheetSource ref="B2:F50"/>
  </cacheSource>
  <cacheFields count="2">
    <cacheField name="Alpha" numFmtId="0"/>
    <cacheField name="Beta"  numFmtId="0"/>
  </cacheFields>
</pivotCacheDefinition>"#;

    const NO_REF_ATTR_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cacheSource type="worksheet">
    <worksheetSource sheet="DataSheet"/>
  </cacheSource>
  <cacheFields count="1">
    <cacheField name="Value" numFmtId="0"/>
  </cacheFields>
</pivotCacheDefinition>"#;

    const EXTERNAL_SOURCE_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cacheSource type="external">
    <consolidation autoPage="1"/>
  </cacheSource>
  <cacheFields count="2">
    <cacheField name="X" numFmtId="0"/>
    <cacheField name="Y" numFmtId="0"/>
  </cacheFields>
</pivotCacheDefinition>"#;

    const EMPTY_FIELDS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cacheSource type="worksheet">
    <worksheetSource ref="A1:A1" sheet="Sheet1"/>
  </cacheSource>
  <cacheFields count="0"/>
</pivotCacheDefinition>"#;

    // ── source_sheet ──────────────────────────────────────────────────────────

    #[test]
    fn extracts_source_sheet() {
        let r = parse_bytes(FULL_CACHE_XML.as_bytes()).unwrap();
        assert_eq!(r.source_sheet.as_deref(), Some("SourceData"));
    }

    #[test]
    fn missing_sheet_attr_is_none() {
        let r = parse_bytes(NO_SHEET_ATTR_XML.as_bytes()).unwrap();
        assert!(r.source_sheet.is_none());
    }

    #[test]
    fn external_source_has_no_sheet() {
        let r = parse_bytes(EXTERNAL_SOURCE_XML.as_bytes()).unwrap();
        assert!(r.source_sheet.is_none());
    }

    #[test]
    fn no_ref_but_sheet_present() {
        let r = parse_bytes(NO_REF_ATTR_XML.as_bytes()).unwrap();
        assert_eq!(r.source_sheet.as_deref(), Some("DataSheet"));
    }

    // ── source_range ──────────────────────────────────────────────────────────

    #[test]
    fn extracts_source_range() {
        let r = parse_bytes(FULL_CACHE_XML.as_bytes()).unwrap();
        assert_eq!(r.source_range.as_deref(), Some("A1:D101"));
    }

    #[test]
    fn extracts_range_without_sheet() {
        let r = parse_bytes(NO_SHEET_ATTR_XML.as_bytes()).unwrap();
        assert_eq!(r.source_range.as_deref(), Some("B2:F50"));
    }

    #[test]
    fn missing_ref_attr_is_none() {
        let r = parse_bytes(NO_REF_ATTR_XML.as_bytes()).unwrap();
        assert!(r.source_range.is_none());
    }

    // ── field_names ───────────────────────────────────────────────────────────

    #[test]
    fn extracts_four_field_names() {
        let r = parse_bytes(FULL_CACHE_XML.as_bytes()).unwrap();
        assert_eq!(r.field_names.len(), 4);
    }

    #[test]
    fn field_names_in_order() {
        let r = parse_bytes(FULL_CACHE_XML.as_bytes()).unwrap();
        assert_eq!(r.field_names[0], "Region");
        assert_eq!(r.field_names[1], "Product");
        assert_eq!(r.field_names[2], "Category");
        assert_eq!(r.field_names[3], "Sales");
    }

    #[test]
    fn two_field_names_extracted() {
        let r = parse_bytes(NO_SHEET_ATTR_XML.as_bytes()).unwrap();
        assert_eq!(r.field_names, vec!["Alpha", "Beta"]);
    }

    #[test]
    fn single_field_name() {
        let r = parse_bytes(NO_REF_ATTR_XML.as_bytes()).unwrap();
        assert_eq!(r.field_names, vec!["Value"]);
    }

    #[test]
    fn external_source_still_has_field_names() {
        let r = parse_bytes(EXTERNAL_SOURCE_XML.as_bytes()).unwrap();
        assert_eq!(r.field_names, vec!["X", "Y"]);
    }

    #[test]
    fn empty_cache_fields_returns_empty_vec() {
        let r = parse_bytes(EMPTY_FIELDS_XML.as_bytes()).unwrap();
        assert!(r.field_names.is_empty());
    }

    // ── worksheetSource must be inside cacheSource ────────────────────────────

    #[test]
    fn worksheet_source_outside_cache_source_ignored() {
        // A stray <worksheetSource> that is NOT inside <cacheSource> must not
        // be parsed (the `in_cache_source` guard must hold).
        let xml = r#"<?xml version="1.0"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <worksheetSource ref="STRAY:REF" sheet="StraySheet"/>
  <cacheSource type="worksheet">
    <worksheetSource ref="A1:B2" sheet="Real"/>
  </cacheSource>
  <cacheFields count="0"/>
</pivotCacheDefinition>"#;
        let r = parse_bytes(xml.as_bytes()).unwrap();
        assert_eq!(
            r.source_sheet.as_deref(),
            Some("Real"),
            "stray worksheetSource outside cacheSource should be ignored"
        );
        assert_eq!(r.source_range.as_deref(), Some("A1:B2"));
    }
}
