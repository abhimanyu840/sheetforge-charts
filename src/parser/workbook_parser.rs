//! Parser for `xl/workbook.xml`.

use anyhow::{Context, Result};
use quick_xml::{events::Event, Reader};

use crate::{
    archive::zip_reader::{read_entry_to_string, XlsxArchive},
    model::workbook::{SheetCharts, WorkbookCharts},
};

pub fn parse(archive: &mut XlsxArchive, workbook_path: &str) -> Result<WorkbookCharts> {
    let xml = read_entry_to_string(archive, workbook_path)
        .with_context(|| format!("Cannot read workbook part: {workbook_path}"))?;

    let sheets = parse_sheets_xml(&xml)
        .with_context(|| format!("Failed to parse sheets in: {workbook_path}"))?;

    Ok(WorkbookCharts {
        source_path: workbook_path.to_owned(),
        sheets,
        theme: None,
    })
}

pub(crate) fn parse_sheets_xml(xml: &str) -> Result<Vec<SheetCharts>> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);

    let mut sheets: Vec<SheetCharts> = Vec::new();
    let mut index: usize = 0;

    loop {
        match reader.read_event()? {
            Event::Empty(ref e) | Event::Start(ref e) => {
                if e.local_name().as_ref() == b"sheet" {
                    let decoder = reader.decoder();
                    if let Some(sheet) = parse_sheet_element(e, decoder, index)? {
                        sheets.push(sheet);
                        index += 1;
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(sheets)
}

fn parse_sheet_element(
    e: &quick_xml::events::BytesStart<'_>,
    decoder: quick_xml::Decoder,
    index: usize,
) -> Result<Option<SheetCharts>> {
    let mut name = String::new();
    let mut r_id = String::new();

    for attr in e.attributes() {
        let attr = attr.context("Malformed attribute in <sheet> element")?;
        match attr.key.local_name().as_ref() {
            b"name" => name = attr.decode_and_unescape_value(decoder)?.into_owned(),
            b"id" => r_id = attr.decode_and_unescape_value(decoder)?.into_owned(),
            _ => {}
        }
    }

    if name.is_empty() {
        return Ok(None);
    }

    Ok(Some(SheetCharts::new(name, r_id, index)))
}

#[cfg(test)]
mod tests {
    use super::*;

    const WORKBOOK_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sales"    sheetId="1" r:id="rId1"/>
    <sheet name="Expenses" sheetId="2" r:id="rId2"/>
    <sheet name="Charts"   sheetId="3" r:id="rId3"/>
  </sheets>
</workbook>"#;

    #[test]
    fn parses_three_sheets() {
        let sheets = parse_sheets_xml(WORKBOOK_XML).unwrap();
        assert_eq!(sheets.len(), 3);
    }

    #[test]
    fn sheet_names_and_indices_correct() {
        let sheets = parse_sheets_xml(WORKBOOK_XML).unwrap();
        assert_eq!(sheets[0].name, "Sales");
        assert_eq!(sheets[0].index, 0);
        assert_eq!(sheets[2].name, "Charts");
        assert_eq!(sheets[2].index, 2);
    }

    #[test]
    fn relationship_ids_preserved() {
        let sheets = parse_sheets_xml(WORKBOOK_XML).unwrap();
        assert_eq!(sheets[0].relationship_id, "rId1");
        assert_eq!(sheets[2].relationship_id, "rId3");
    }

    #[test]
    fn sheets_start_with_empty_chart_list() {
        let sheets = parse_sheets_xml(WORKBOOK_XML).unwrap();
        for sheet in &sheets {
            assert!(sheet.charts.is_empty());
        }
    }
}
