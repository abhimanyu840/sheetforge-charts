//! Pivot table metadata attached to pivot-backed charts.
//!
//! ## OOXML relationship chain
//!
//! ```text
//! xl/charts/chart1.xml
//!   └─[.rels pivotTable]──► xl/pivotTables/pivotTable1.xml
//!                                └─[.rels pivotCacheDefinition]──►
//!                                         xl/pivotCache/pivotCacheDefinition1.xml
//! ```
//!
//! * `chart1.xml` carries `<c:pivotSource><c:name>Sheet1!PivotTable1</c:name>` —
//!   already parsed into [`crate::model::chart::Chart::pivot_table_name`] in Phase 10.
//! * `pivotTable1.xml` (`<pivotTableDefinition>`) gives the canonical name and the
//!   ordered list of pivot fields.
//! * `pivotCacheDefinition1.xml` (`<pivotCacheDefinition>`) gives the data source:
//!   worksheet name and cell range.

use serde::{Deserialize, Serialize};

use crate::model::series::Series;

// ── PivotField ────────────────────────────────────────────────────────────────

/// A single field (column) in a pivot table, from `<pivotField>` inside
/// `<pivotFields>` in `pivotTableDefinition.xml`.
///
/// The `name` attribute is carried on `<cacheField>` in the cache definition,
/// not on `<pivotField>` itself — the parser joins them positionally.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct PivotField {
    /// Field name from `<cacheField name="…"/>`, indexed in the same order as
    /// `<pivotField>` entries in `pivotTableDefinition.xml`.
    pub name: String,
}

// ── PivotTableMeta ────────────────────────────────────────────────────────────

/// Full pivot table metadata resolved from the chart → pivotTable →
/// pivotCacheDefinition relationship chain.
///
/// Attached to [`crate::model::chart::Chart`] as `pivot_meta` when
/// [`crate::model::chart::Chart::is_pivot_chart`] is `true` **and** the
/// relationship chain can be fully resolved.  May be `None` even for pivot
/// charts if the pivot table or cache parts are missing or malformed.
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PivotTableMeta {
    /// Canonical name from `<pivotTableDefinition name="…"/>`.
    ///
    /// This is the name shown in Excel's PivotTable Field List and matches the
    /// name embedded in `<c:pivotSource><c:name>`.  For the chart-source
    /// variant the name may be prefixed with a sheet name separated by `!`.
    pub pivot_table_name: String,

    /// Ordered list of fields from the pivot cache, one per `<cacheField>` in
    /// `pivotCacheDefinition.xml`.
    ///
    /// The order matches the column order in the source range.  An empty `Vec`
    /// means the cache definition was present but contained no `<cacheField>`
    /// children (unusual but not illegal).
    pub pivot_fields: Vec<PivotField>,

    /// Source worksheet name from
    /// `<cacheSource><worksheetSource sheet="…"/></cacheSource>`.
    ///
    /// `None` when the cache source is not a worksheet (e.g. external data,
    /// OLAP, or scenario cache) or when the `sheet` attribute is absent.
    pub source_sheet: Option<String>,

    /// Source cell range from
    /// `<cacheSource><worksheetSource ref="…"/></cacheSource>`.
    ///
    /// Written as a standard A1-notation range string (e.g. `"A1:D100"`).
    /// `None` when absent (some sources use a named table instead of a range).
    pub source_range: Option<String>,

    /// Aggregated series built from `pivotCacheRecordsN.xml`.
    ///
    /// Each [`Series`] corresponds to one combination of column-field value and
    /// data-field:
    /// * `series.name` — column-field value (e.g. `"Widget"`) or data-field
    ///   display name when there are no column fields.
    /// * `series.category_values` — ordered unique row-field values (the
    ///   category axis labels).
    /// * `series.value_cache` — aggregated values aligned to category order.
    /// * `series.value_cache_state = CacheState::Complete`.
    ///
    /// Empty when the cache-records file is absent, unreadable, or when the
    /// pivot table has no data fields configured.
    pub pivot_series: Vec<Series>,
}

// ── Unit tests ────────────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;

    fn make_meta(name: &str, fields: &[&str]) -> PivotTableMeta {
        PivotTableMeta {
            pivot_table_name: name.to_owned(),
            pivot_fields: fields
                .iter()
                .map(|n| PivotField {
                    name: n.to_string(),
                })
                .collect(),
            source_sheet: Some("Sheet1".to_owned()),
            source_range: Some("A1:D10".to_owned()),
            pivot_series: vec![],
        }
    }

    #[test]
    fn pivot_field_name_roundtrip() {
        let f = PivotField {
            name: "Sales".to_owned(),
        };
        assert_eq!(f.name, "Sales");
    }

    #[test]
    fn meta_field_count() {
        let m = make_meta("PivotTable1", &["Region", "Product", "Sales"]);
        assert_eq!(m.pivot_fields.len(), 3);
    }

    #[test]
    fn meta_field_names_in_order() {
        let m = make_meta("PivotTable1", &["Region", "Product", "Sales"]);
        assert_eq!(m.pivot_fields[0].name, "Region");
        assert_eq!(m.pivot_fields[1].name, "Product");
        assert_eq!(m.pivot_fields[2].name, "Sales");
    }

    #[test]
    fn meta_source_sheet() {
        let m = make_meta("T1", &[]);
        assert_eq!(m.source_sheet.as_deref(), Some("Sheet1"));
    }

    #[test]
    fn meta_source_range() {
        let m = make_meta("T1", &[]);
        assert_eq!(m.source_range.as_deref(), Some("A1:D10"));
    }

    #[test]
    fn meta_no_source_sheet() {
        let m = PivotTableMeta {
            pivot_table_name: "T1".into(),
            pivot_fields: vec![],
            source_sheet: None,
            source_range: None,
            pivot_series: vec![],
        };
        assert!(m.source_sheet.is_none());
        assert!(m.source_range.is_none());
    }

    #[test]
    fn empty_fields_vec() {
        let m = make_meta("Empty", &[]);
        assert!(m.pivot_fields.is_empty());
    }

    #[test]
    fn pivot_field_equality() {
        let a = PivotField {
            name: "Revenue".into(),
        };
        let b = PivotField {
            name: "Revenue".into(),
        };
        let c = PivotField {
            name: "Cost".into(),
        };
        assert_eq!(a, b);
        assert_ne!(a, c);
    }
}
