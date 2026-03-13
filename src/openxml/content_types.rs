//! Parser for `[Content_Types].xml`.

use anyhow::{Context, Result};
use quick_xml::{events::Event, Reader};

use crate::archive::zip_reader::{read_entry_to_string, XlsxArchive};

const CONTENT_TYPES_PATH: &str = "[Content_Types].xml";

const WORKBOOK_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";

const WORKBOOK_MACRO_CONTENT_TYPE: &str =
    "application/vnd.ms-excel.sheet.macroEnabled.main+xml";

#[derive(Debug, Default)]
pub struct ContentTypes {
    overrides: Vec<(String, String)>,
}

impl ContentTypes {
    pub fn workbook_path(&self) -> Result<String> {
        self.overrides
            .iter()
            .find(|(_, ct)| ct == WORKBOOK_CONTENT_TYPE || ct == WORKBOOK_MACRO_CONTENT_TYPE)
            .map(|(path, _)| path.clone())
            .with_context(|| {
                "No workbook part found in [Content_Types].xml — is this a valid .xlsx file?"
            })
    }

    pub fn content_type_for(&self, part_path: &str) -> Option<&str> {
        self.overrides
            .iter()
            .find(|(p, _)| p == part_path)
            .map(|(_, ct)| ct.as_str())
    }
}

pub fn parse(archive: &mut XlsxArchive) -> Result<ContentTypes> {
    let xml = read_entry_to_string(archive, CONTENT_TYPES_PATH)
        .context("Cannot read [Content_Types].xml")?;
    parse_xml(&xml)
}

pub(crate) fn parse_xml(xml: &str) -> Result<ContentTypes> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut ct = ContentTypes::default();

    loop {
        match reader.read_event()? {
            Event::Empty(ref e) | Event::Start(ref e) => {
                if e.local_name().as_ref() == b"Override" {
                    let decoder = reader.decoder();
                    let (part_name, content_type) =
                        extract_override_attrs(e, decoder)?;
                    if !part_name.is_empty() && !content_type.is_empty() {
                        let normalised = part_name.trim_start_matches('/').to_owned();
                        ct.overrides.push((normalised, content_type));
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(ct)
}

fn extract_override_attrs(
    e: &quick_xml::events::BytesStart<'_>,
    decoder: quick_xml::Decoder,
) -> Result<(String, String)> {
    let mut part_name = String::new();
    let mut content_type = String::new();

    for attr in e.attributes() {
        let attr = attr.context("Malformed attribute in [Content_Types].xml")?;
        match attr.key.local_name().as_ref() {
            b"PartName" => {
                part_name = attr.decode_and_unescape_value(decoder)?.into_owned();
            }
            b"ContentType" => {
                content_type = attr.decode_and_unescape_value(decoder)?.into_owned();
            }
            _ => {}
        }
    }

    Ok((part_name, content_type))
}

#[cfg(test)]
mod tests {
    use super::*;

    const SAMPLE_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/xl/workbook.xml"
            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/charts/chart1.xml"
            ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
</Types>"#;

    #[test]
    fn workbook_path_resolved() {
        let ct = parse_xml(SAMPLE_XML).unwrap();
        assert_eq!(ct.workbook_path().unwrap(), "xl/workbook.xml");
    }

    #[test]
    fn chart_part_recognised() {
        let ct = parse_xml(SAMPLE_XML).unwrap();
        assert_eq!(
            ct.content_type_for("xl/charts/chart1.xml"),
            Some("application/vnd.openxmlformats-officedocument.drawingml.chart+xml")
        );
    }

    #[test]
    fn missing_workbook_returns_error() {
        let ct = parse_xml("<Types/>").unwrap();
        assert!(ct.workbook_path().is_err());
    }
}
