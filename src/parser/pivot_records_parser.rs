//! Parser for `xl/pivotCache/pivotCacheRecordsN.xml`.
//!
//! Reads the raw source-data rows cached by Excel and aggregates them into
//! [`Series`] objects using the pivot-table layout from
//! [`crate::parser::pivot_table_parser::PivotTableRaw`] and the shared-item
//! lookup tables from [`crate::parser::pivot_cache_parser::PivotCacheRaw`].
//!
//! ## Record format
//!
//! ```xml
//! <pivotCacheRecords xmlns="…" count="6">
//!   <r>
//!     <x v="0"/>    <!-- field 0: index into sharedItems[0] → "North" -->
//!     <x v="0"/>    <!-- field 1: index into sharedItems[1] → "Widget" -->
//!     <n v="1500"/> <!-- field 2: inline numeric value 1500 -->
//!   </r>
//!   <r>…</r>
//! </pivotCacheRecords>
//! ```
//!
//! Cell value tags:
//! * `<x v="N"/>` — shared-items index reference; decoded via the lookup table.
//! * `<s v="…"/>` — inline string value (field has no shared items).
//! * `<n v="…"/>` — inline numeric value.
//! * `<b v="0|1"/>` — boolean; stored as `"0"` or `"1"`.
//! * `<e v="…"/>` — error string (e.g. `"#N/A"`).
//! * `<m/>`        — missing / null value; stored as empty string or NaN.
//!
//! ## Aggregation
//!
//! After parsing all rows, this module performs an in-memory GROUP BY:
//!
//! ```text
//! row_key  = join(row_field_values, "|")   e.g. "North"
//! col_key  = join(col_field_values, "|")   e.g. "Widget"
//! ```
//!
//! The data field is aggregated (sum by default) per `(row_key, col_key)` cell.
//!
//! The result is a `Vec<Series>`:
//! * One `Series` per unique `col_key` (or one series named after the data
//!   field when there are no column fields).
//! * `series.category_values` = unique row-key strings in first-seen order.
//! * `series.value_cache`     = aggregated values aligned to category order.
//! * `series.name`            = col_key string (or data-field display name).
//! * `series.value_cache_state = CacheState::Complete`.

use std::collections::HashMap;

use anyhow::{Context, Result};
use quick_xml::{events::Event, Reader};

use crate::{
    archive::zip_reader::{read_entry_bytes, XlsxArchive},
    model::series::{CacheState, DataValues, Series, StringValues},
    parser::{
        pivot_cache_parser::PivotCacheRaw,
        pivot_table_parser::{DataFieldDef, PivotTableRaw},
    },
};

// ── Raw cell value ────────────────────────────────────────────────────────────

/// A decoded cell value from a single `<r>` element in the records file.
#[derive(Debug, Clone)]
enum CellVal {
    Str(String),
    Num(f64),
    Missing,
}

impl CellVal {
    fn as_str(&self) -> String {
        match self {
            CellVal::Str(s)  => s.clone(),
            CellVal::Num(n)  => {
                // Format without trailing ".0" for integers
                if n.fract() == 0.0 && n.abs() < 1e15 {
                    format!("{}", *n as i64)
                } else {
                    format!("{n}")
                }
            }
            CellVal::Missing => String::new(),
        }
    }

    fn as_f64(&self) -> f64 {
        match self {
            CellVal::Num(n)  => *n,
            CellVal::Str(s)  => s.parse().unwrap_or(f64::NAN),
            CellVal::Missing => f64::NAN,
        }
    }
}

// ── Entry points ──────────────────────────────────────────────────────────────

/// Parse records from the archive and aggregate into `Series` objects.
pub fn parse_and_aggregate(
    archive:    &mut XlsxArchive,
    path:       &str,
    cache_def:  &PivotCacheRaw,
    pivot_def:  &PivotTableRaw,
) -> Result<Vec<Series>> {
    let bytes = read_entry_bytes(archive, path)
        .with_context(|| format!("Cannot read pivot cache records: {path}"))?;
    parse_bytes_and_aggregate(&bytes, cache_def, pivot_def)
        .with_context(|| format!("Failed to parse pivot records: {path}"))
}

/// Pure-bytes version — used in unit tests without a real archive.
pub fn parse_bytes_and_aggregate(
    bytes:      &[u8],
    cache_def:  &PivotCacheRaw,
    pivot_def:  &PivotTableRaw,
) -> Result<Vec<Series>> {
    let records = parse_records(bytes, cache_def)?;
    Ok(aggregate(records, cache_def, pivot_def))
}

// ── Record parser ─────────────────────────────────────────────────────────────

fn parse_records(bytes: &[u8], cache_def: &PivotCacheRaw) -> Result<Vec<Vec<CellVal>>> {
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().trim_text(true);

    let mut records: Vec<Vec<CellVal>> = Vec::new();
    let mut current_row: Option<Vec<CellVal>> = None;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) | Event::Empty(ref e) => {
                let local = e.local_name();
                let dec   = reader.decoder();
                match local.as_ref() {
                    b"r" => {
                        current_row = Some(Vec::new());
                    }
                    tag => {
                        if let Some(row) = current_row.as_mut() {
                            let cell = decode_cell(tag, e, dec, &cache_def.shared_items, row.len())?;
                            row.push(cell);
                        }
                    }
                }
            }
            Event::End(ref e) => {
                if e.local_name().as_ref() == b"r" {
                    if let Some(row) = current_row.take() {
                        records.push(row);
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(records)
}

/// Decode one cell element within a record row.
fn decode_cell(
    tag:          &[u8],
    e:            &quick_xml::events::BytesStart<'_>,
    dec:          quick_xml::Decoder,
    shared_items: &[Vec<String>],
    col_idx:      usize,
) -> Result<CellVal> {
    match tag {
        // Shared-items index reference: look up the string from the cache
        b"x" => {
            let mut idx: usize = 0;
            for attr in e.attributes() {
                let attr = attr.context("Malformed attr in <x>")?;
                if attr.key.local_name().as_ref() == b"v" {
                    idx = attr.decode_and_unescape_value(dec)?.parse().unwrap_or(0);
                }
            }
            let val = shared_items
                .get(col_idx)
                .and_then(|items| items.get(idx))
                .cloned()
                .unwrap_or_default();
            // Try to parse as number (numeric sharedItems use <n> but are
            // sometimes stored via <x> with numeric strings in the table)
            if let Ok(n) = val.parse::<f64>() {
                Ok(CellVal::Num(n))
            } else {
                Ok(CellVal::Str(val))
            }
        }
        // Inline string
        b"s" => {
            let mut val = String::new();
            for attr in e.attributes() {
                let attr = attr.context("Malformed attr in <s>")?;
                if attr.key.local_name().as_ref() == b"v" {
                    val = attr.decode_and_unescape_value(dec)?.into_owned();
                }
            }
            Ok(CellVal::Str(val))
        }
        // Inline numeric
        b"n" => {
            let mut val = 0.0_f64;
            for attr in e.attributes() {
                let attr = attr.context("Malformed attr in <n>")?;
                if attr.key.local_name().as_ref() == b"v" {
                    val = attr.decode_and_unescape_value(dec)?.parse().unwrap_or(f64::NAN);
                }
            }
            Ok(CellVal::Num(val))
        }
        // Boolean (0/1)
        b"b" => {
            let mut val = String::new();
            for attr in e.attributes() {
                let attr = attr.context("Malformed attr in <b>")?;
                if attr.key.local_name().as_ref() == b"v" {
                    val = attr.decode_and_unescape_value(dec)?.into_owned();
                }
            }
            Ok(CellVal::Str(val))
        }
        // Error value
        b"e" => {
            let mut val = String::new();
            for attr in e.attributes() {
                let attr = attr.context("Malformed attr in <e>")?;
                if attr.key.local_name().as_ref() == b"v" {
                    val = attr.decode_and_unescape_value(dec)?.into_owned();
                }
            }
            Ok(CellVal::Str(val))
        }
        // Missing / null
        b"m" => Ok(CellVal::Missing),
        // Unknown tag inside a row — skip gracefully
        _ => Ok(CellVal::Missing),
    }
}

// ── Aggregation ───────────────────────────────────────────────────────────────

fn aggregate(
    records:   Vec<Vec<CellVal>>,
    cache_def: &PivotCacheRaw,
    pivot_def: &PivotTableRaw,
) -> Vec<Series> {
    if pivot_def.data_fields.is_empty() {
        return vec![];
    }

    let row_idxs  = &pivot_def.row_field_idxs;
    let col_idxs  = &pivot_def.col_field_idxs;
    let data_flds = &pivot_def.data_fields;
    let n_fields  = cache_def.field_names.len();

    // ── Build ordered unique keys and aggregation map ─────────────────────────
    // For each data field, accumulate sums keyed by (row_key, col_key).
    // Use insertion-ordered Vecs to preserve natural source order.

    let mut unique_cats:   Vec<String>                        = Vec::new();
    let mut unique_series: Vec<String>                        = Vec::new();
    // agg_map[data_field_idx][(cat_idx, ser_idx)] = (sum, count)
    let mut agg_maps: Vec<HashMap<(usize, usize), (f64, u64)>> =
        (0..data_flds.len()).map(|_| HashMap::new()).collect();

    for row in &records {
        if row.len() < n_fields { continue; }

        // Build composite row key and col key
        let row_key = composite_key(row, row_idxs);
        let col_key = composite_key(row, col_idxs);

        // Find or insert category index (preserve insertion order)
        let cat_idx = find_or_push(&mut unique_cats, &row_key);

        // Find or insert series index
        let ser_idx = find_or_push(&mut unique_series, &col_key);

        // Accumulate data fields
        for (di, df) in data_flds.iter().enumerate() {
            if df.field_idx >= row.len() { continue; }
            let v = row[df.field_idx].as_f64();
            if !v.is_nan() {
                let entry = agg_maps[di].entry((cat_idx, ser_idx)).or_insert((0.0, 0));
                entry.0 += v;
                entry.1 += 1;
            }
        }
    }

    // ── Build final Series objects ────────────────────────────────────────────
    let n_cats = unique_cats.len();
    let n_ser  = unique_series.len().max(1); // at least 1 series

    let mut out: Vec<Series> = Vec::with_capacity(data_flds.len() * n_ser);

    for (di, df) in data_flds.iter().enumerate() {
        for si in 0..n_ser {
            let ser_name = if unique_series.is_empty() {
                df.name.clone()
            } else {
                format!("{} — {}", df.name, &unique_series[si])
            };

            // Build value array aligned to categories
            let values: Vec<f64> = (0..n_cats)
                .map(|ci| {
                    agg_maps[di]
                        .get(&(ci, si))
                        .map(|(sum, _)| apply_subtotal(df, *sum, agg_maps[di].get(&(ci, si)).map(|e| e.1).unwrap_or(0)))
                        .unwrap_or(f64::NAN)
                })
                .collect();

            let mut s = Series::new((di * n_ser + si) as u32);
            s.name = Some(ser_name);
            s.category_values = Some(StringValues {
                pt_count: Some(n_cats),
                values:   unique_cats.clone(),
            });
            s.value_cache = Some(DataValues {
                pt_count:    Some(n_cats),
                values,
                format_code: None,
            });
            s.value_cache_state       = CacheState::Complete;
            s.category_cache_state    = CacheState::Complete;
            out.push(s);
        }
    }

    out
}

/// Build a composite key string from the values of specific column indices.
fn composite_key(row: &[CellVal], idxs: &[usize]) -> String {
    if idxs.is_empty() {
        return String::new();
    }
    idxs.iter()
        .map(|&i| if i < row.len() { row[i].as_str() } else { String::new() })
        .collect::<Vec<_>>()
        .join("|")
}

/// Find `key` in `vec`, returning its index.  If not found, push and return new index.
fn find_or_push(vec: &mut Vec<String>, key: &str) -> usize {
    if let Some(i) = vec.iter().position(|x| x == key) {
        return i;
    }
    vec.push(key.to_owned());
    vec.len() - 1
}

/// Apply the subtotal function for a data field.
fn apply_subtotal(df: &DataFieldDef, sum: f64, count: u64) -> f64 {
    match df.subtotal.to_ascii_lowercase().as_str() {
        "average" | "avg" => if count > 0 { sum / count as f64 } else { f64::NAN },
        "count"           => count as f64,
        _                 => sum,   // "sum" and everything else
    }
}

// ── Tests ─────────────────────────────────────────────────────────────────────

#[cfg(test)]
pub(crate) mod tests {
    use super::*;
    use crate::parser::{
        pivot_cache_parser::PivotCacheRaw,
        pivot_table_parser::{DataFieldDef, PivotTableRaw},
    };

    // ── Shared helpers ────────────────────────────────────────────────────────

    fn make_cache(field_names: &[&str], shared_items: Vec<Vec<&str>>) -> PivotCacheRaw {
        PivotCacheRaw {
            source_sheet: Some("Data".into()),
            source_range: Some("A1:D10".into()),
            field_names: field_names.iter().map(|s| s.to_string()).collect(),
            shared_items: shared_items.into_iter()
                .map(|v| v.into_iter().map(|s| s.to_string()).collect())
                .collect(),
        }
    }

    fn make_pivot(
        row_field_idxs: Vec<usize>,
        col_field_idxs: Vec<usize>,
        data_fields: Vec<(usize, &str, &str)>, // (field_idx, name, subtotal)
    ) -> PivotTableRaw {
        PivotTableRaw {
            name: "TestPivot".into(),
            field_count: 0,
            row_field_idxs,
            col_field_idxs,
            data_fields: data_fields.into_iter()
                .map(|(fi, n, sub)| DataFieldDef {
                    field_idx: fi,
                    name: n.to_string(),
                    subtotal: sub.to_string(),
                })
                .collect(),
        }
    }

    // ── XML fixtures ──────────────────────────────────────────────────────────

    /// 6 records: Region(x-ref) x Sales(numeric), no column field.
    /// Region sharedItems: [North(0), South(1), East(2)]
    pub(crate) const SIMPLE_RECORDS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="6">
  <r><x v="0"/><n v="1500"/></r>
  <r><x v="0"/><n v="2300"/></r>
  <r><x v="1"/><n v="800"/> </r>
  <r><x v="1"/><n v="1200"/></r>
  <r><x v="2"/><n v="3100"/></r>
  <r><x v="2"/><n v="900"/> </r>
</pivotCacheRecords>"#;

    /// 6 records: Region x Product (col field) x Sales.
    pub(crate) const MATRIX_RECORDS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="6">
  <r><x v="0"/><x v="0"/><n v="1500"/></r>
  <r><x v="0"/><x v="1"/><n v="2300"/></r>
  <r><x v="1"/><x v="0"/><n v="800"/> </r>
  <r><x v="1"/><x v="1"/><n v="1200"/></r>
  <r><x v="2"/><x v="0"/><n v="3100"/></r>
  <r><x v="2"/><x v="1"/><n v="900"/> </r>
</pivotCacheRecords>"#;

    /// Uses inline <s> and <n> (no shared items).
    const INLINE_RECORDS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="3">
  <r><s v="Alpha"/><n v="10"/></r>
  <r><s v="Beta"/> <n v="20"/></r>
  <r><s v="Alpha"/><n v="30"/></r>
</pivotCacheRecords>"#;

    const EMPTY_RECORDS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0">
</pivotCacheRecords>"#;

    // ── record parsing ────────────────────────────────────────────────────────

    #[test]
    fn parses_six_records() {
        let cache = make_cache(&["Region", "Sales"],
            vec![vec!["North", "South", "East"], vec![]]);
        let records = parse_records(SIMPLE_RECORDS_XML.as_bytes(), &cache).unwrap();
        assert_eq!(records.len(), 6);
    }

    #[test]
    fn x_ref_decoded_to_string() {
        let cache = make_cache(&["Region", "Sales"],
            vec![vec!["North", "South", "East"], vec![]]);
        let records = parse_records(SIMPLE_RECORDS_XML.as_bytes(), &cache).unwrap();
        assert_eq!(records[0][0].as_str(), "North");
        assert_eq!(records[2][0].as_str(), "South");
        assert_eq!(records[4][0].as_str(), "East");
    }

    #[test]
    fn numeric_cell_parsed() {
        let cache = make_cache(&["Region", "Sales"],
            vec![vec!["North", "South", "East"], vec![]]);
        let records = parse_records(SIMPLE_RECORDS_XML.as_bytes(), &cache).unwrap();
        assert_eq!(records[0][1].as_f64(), 1500.0);
        assert_eq!(records[1][1].as_f64(), 2300.0);
    }

    #[test]
    fn inline_string_cell_parsed() {
        let cache = make_cache(&["Label", "Value"], vec![vec![], vec![]]);
        let records = parse_records(INLINE_RECORDS_XML.as_bytes(), &cache).unwrap();
        assert_eq!(records[0][0].as_str(), "Alpha");
        assert_eq!(records[1][0].as_str(), "Beta");
    }

    #[test]
    fn empty_records_parses_zero_rows() {
        let cache = make_cache(&["A"], vec![vec![]]);
        let records = parse_records(EMPTY_RECORDS_XML.as_bytes(), &cache).unwrap();
        assert_eq!(records.len(), 0);
    }

    // ── aggregation: no col field (single series) ─────────────────────────────

    #[test]
    fn aggregate_single_series_count() {
        let cache = make_cache(&["Region", "Sales"],
            vec![vec!["North", "South", "East"], vec![]]);
        let pivot = make_pivot(vec![0], vec![], vec![(1, "Sum of Sales", "sum")]);
        let series = parse_bytes_and_aggregate(SIMPLE_RECORDS_XML.as_bytes(), &cache, &pivot).unwrap();
        assert_eq!(series.len(), 1);
    }

    #[test]
    fn aggregate_single_series_name() {
        let cache = make_cache(&["Region", "Sales"],
            vec![vec!["North", "South", "East"], vec![]]);
        let pivot = make_pivot(vec![0], vec![], vec![(1, "Sum of Sales", "sum")]);
        let series = parse_bytes_and_aggregate(SIMPLE_RECORDS_XML.as_bytes(), &cache, &pivot).unwrap();
        assert_eq!(series[0].name.as_deref(), Some("Sum of Sales"));
    }

    #[test]
    fn aggregate_categories_in_order() {
        let cache = make_cache(&["Region", "Sales"],
            vec![vec!["North", "South", "East"], vec![]]);
        let pivot = make_pivot(vec![0], vec![], vec![(1, "Sum of Sales", "sum")]);
        let series = parse_bytes_and_aggregate(SIMPLE_RECORDS_XML.as_bytes(), &cache, &pivot).unwrap();
        let cats = &series[0].category_values.as_ref().unwrap().values;
        assert_eq!(cats, &vec!["North", "South", "East"]);
    }

    #[test]
    fn aggregate_sums_correctly() {
        let cache = make_cache(&["Region", "Sales"],
            vec![vec!["North", "South", "East"], vec![]]);
        let pivot = make_pivot(vec![0], vec![], vec![(1, "Sum of Sales", "sum")]);
        let series = parse_bytes_and_aggregate(SIMPLE_RECORDS_XML.as_bytes(), &cache, &pivot).unwrap();
        let vals = &series[0].value_cache.as_ref().unwrap().values;
        // North: 1500+2300=3800, South: 800+1200=2000, East: 3100+900=4000
        assert_eq!(vals, &vec![3800.0, 2000.0, 4000.0]);
    }

    #[test]
    fn aggregate_value_cache_state_is_complete() {
        let cache = make_cache(&["Region", "Sales"],
            vec![vec!["North", "South", "East"], vec![]]);
        let pivot = make_pivot(vec![0], vec![], vec![(1, "Sum of Sales", "sum")]);
        let series = parse_bytes_and_aggregate(SIMPLE_RECORDS_XML.as_bytes(), &cache, &pivot).unwrap();
        assert_eq!(series[0].value_cache_state, CacheState::Complete);
        assert_eq!(series[0].category_cache_state, CacheState::Complete);
    }

    // ── aggregation: with col field (multi-series) ────────────────────────────

    #[test]
    fn aggregate_matrix_series_count() {
        // Region=row, Product=col, Sales=data
        let cache = make_cache(&["Region", "Product", "Sales"],
            vec![
                vec!["North", "South", "East"],   // Region sharedItems
                vec!["Widget", "Gadget"],           // Product sharedItems
                vec![],                             // Sales: numeric
            ]);
        let pivot = make_pivot(vec![0], vec![1], vec![(2, "Sum of Sales", "sum")]);
        let series = parse_bytes_and_aggregate(MATRIX_RECORDS_XML.as_bytes(), &cache, &pivot).unwrap();
        // 1 data field × 2 col values = 2 series
        assert_eq!(series.len(), 2);
    }

    #[test]
    fn aggregate_matrix_series_names() {
        let cache = make_cache(&["Region", "Product", "Sales"],
            vec![
                vec!["North", "South", "East"],
                vec!["Widget", "Gadget"],
                vec![],
            ]);
        let pivot = make_pivot(vec![0], vec![1], vec![(2, "Sum of Sales", "sum")]);
        let series = parse_bytes_and_aggregate(MATRIX_RECORDS_XML.as_bytes(), &cache, &pivot).unwrap();
        assert_eq!(series[0].name.as_deref(), Some("Sum of Sales — Widget"));
        assert_eq!(series[1].name.as_deref(), Some("Sum of Sales — Gadget"));
    }

    #[test]
    fn aggregate_matrix_categories_shared() {
        let cache = make_cache(&["Region", "Product", "Sales"],
            vec![
                vec!["North", "South", "East"],
                vec!["Widget", "Gadget"],
                vec![],
            ]);
        let pivot = make_pivot(vec![0], vec![1], vec![(2, "Sum of Sales", "sum")]);
        let series = parse_bytes_and_aggregate(MATRIX_RECORDS_XML.as_bytes(), &cache, &pivot).unwrap();
        // Both series must have the same 3 categories
        let cats0 = &series[0].category_values.as_ref().unwrap().values;
        let cats1 = &series[1].category_values.as_ref().unwrap().values;
        assert_eq!(cats0, &vec!["North", "South", "East"]);
        assert_eq!(cats0, cats1);
    }

    #[test]
    fn aggregate_matrix_widget_values() {
        let cache = make_cache(&["Region", "Product", "Sales"],
            vec![vec!["North","South","East"], vec!["Widget","Gadget"], vec![]]);
        let pivot = make_pivot(vec![0], vec![1], vec![(2, "Sum of Sales", "sum")]);
        let series = parse_bytes_and_aggregate(MATRIX_RECORDS_XML.as_bytes(), &cache, &pivot).unwrap();
        let vals = &series[0].value_cache.as_ref().unwrap().values;
        // Widget: North=1500, South=800, East=3100
        assert_eq!(vals, &vec![1500.0, 800.0, 3100.0]);
    }

    #[test]
    fn aggregate_matrix_gadget_values() {
        let cache = make_cache(&["Region", "Product", "Sales"],
            vec![vec!["North","South","East"], vec!["Widget","Gadget"], vec![]]);
        let pivot = make_pivot(vec![0], vec![1], vec![(2, "Sum of Sales", "sum")]);
        let series = parse_bytes_and_aggregate(MATRIX_RECORDS_XML.as_bytes(), &cache, &pivot).unwrap();
        let vals = &series[1].value_cache.as_ref().unwrap().values;
        // Gadget: North=2300, South=1200, East=900
        assert_eq!(vals, &vec![2300.0, 1200.0, 900.0]);
    }

    // ── subtotal functions ────────────────────────────────────────────────────

    #[test]
    fn average_subtotal() {
        let cache = make_cache(&["Region", "Sales"],
            vec![vec!["North", "South", "East"], vec![]]);
        let pivot = make_pivot(vec![0], vec![], vec![(1, "Avg Sales", "average")]);
        let series = parse_bytes_and_aggregate(SIMPLE_RECORDS_XML.as_bytes(), &cache, &pivot).unwrap();
        let vals = &series[0].value_cache.as_ref().unwrap().values;
        // North: (1500+2300)/2=1900, South: (800+1200)/2=1000, East: (3100+900)/2=2000
        assert_eq!(vals, &vec![1900.0, 1000.0, 2000.0]);
    }

    #[test]
    fn count_subtotal() {
        let cache = make_cache(&["Region", "Sales"],
            vec![vec!["North", "South", "East"], vec![]]);
        let pivot = make_pivot(vec![0], vec![], vec![(1, "Count", "count")]);
        let series = parse_bytes_and_aggregate(SIMPLE_RECORDS_XML.as_bytes(), &cache, &pivot).unwrap();
        let vals = &series[0].value_cache.as_ref().unwrap().values;
        // Each region appears 2 times
        assert_eq!(vals, &vec![2.0, 2.0, 2.0]);
    }

    // ── edge cases ────────────────────────────────────────────────────────────

    #[test]
    fn empty_records_produces_empty_series() {
        let cache = make_cache(&["Region", "Sales"],
            vec![vec!["North"], vec![]]);
        let pivot = make_pivot(vec![0], vec![], vec![(1, "Sum", "sum")]);
        let series = parse_bytes_and_aggregate(EMPTY_RECORDS_XML.as_bytes(), &cache, &pivot).unwrap();
        // One series but with 0 categories
        assert_eq!(series.len(), 1);
        assert!(series[0].value_cache.as_ref().unwrap().values.is_empty());
    }

    #[test]
    fn no_data_fields_returns_empty() {
        let cache = make_cache(&["Region"], vec![vec!["North"]]);
        let pivot = make_pivot(vec![0], vec![], vec![]);
        let series = parse_bytes_and_aggregate(SIMPLE_RECORDS_XML.as_bytes(), &cache, &pivot).unwrap();
        assert!(series.is_empty());
    }

    #[test]
    fn inline_string_aggregate() {
        let cache = make_cache(&["Label", "Value"], vec![vec![], vec![]]);
        let pivot = make_pivot(vec![0], vec![], vec![(1, "Sum of Value", "sum")]);
        let series = parse_bytes_and_aggregate(INLINE_RECORDS_XML.as_bytes(), &cache, &pivot).unwrap();
        let cats = &series[0].category_values.as_ref().unwrap().values;
        // "Alpha" appears twice (rows 0 and 2), "Beta" once
        assert_eq!(cats, &vec!["Alpha", "Beta"]);
        let vals = &series[0].value_cache.as_ref().unwrap().values;
        assert_eq!(vals, &vec![40.0, 20.0]);  // Alpha: 10+30=40, Beta: 20
    }
}