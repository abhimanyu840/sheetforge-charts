//! Parser for OpenXML relationship files (`.rels`) and the
//! [`RelationshipResolver`] that chains them together.

use std::collections::HashMap;
use std::path::{Path, PathBuf};

use anyhow::{Context, Result};
use quick_xml::{events::Event, Reader};

use crate::archive::zip_reader::{read_entry_to_string, XlsxArchive};

// ── Relationship type URI constants ───────────────────────────────────────────

pub mod rel_type {
    pub const WORKSHEET: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
    pub const CHARTSHEET: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
    pub const SHARED_STRINGS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
    pub const STYLES: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    pub const DRAWING: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
    pub const CHART: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
    pub const THEME: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
    /// Relationship from a chart part to its backing pivot table definition.
    /// Present in `xl/charts/_rels/chartN.xml.rels` for pivot charts.
    pub const PIVOT_TABLE: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";
    /// Relationship from a pivot table definition to its cache definition.
    /// Present in `xl/pivotTables/_rels/pivotTableN.xml.rels`.
    pub const PIVOT_CACHE_DEF: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";
    /// Relationship from a pivot cache definition to its cache records.
    /// Present in `xl/pivotCache/_rels/pivotCacheDefinitionN.xml.rels`.
    pub const PIVOT_CACHE_RECORDS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords";
}

// ── Core types ────────────────────────────────────────────────────────────────

#[derive(Debug, Clone, PartialEq)]
pub struct Relationship {
    pub id: String,
    pub rel_type: String,
    pub target: String,
}

impl Relationship {
    pub fn is_type(&self, rt: &str) -> bool {
        self.rel_type == rt
    }
}

#[derive(Debug, Default)]
pub struct Relationships {
    entries: Vec<Relationship>,
    id_index: HashMap<String, usize>,
}

impl Relationships {
    fn push(&mut self, rel: Relationship) {
        let idx = self.entries.len();
        self.id_index.insert(rel.id.clone(), idx);
        self.entries.push(rel);
    }

    pub fn all(&self) -> &[Relationship] {
        &self.entries
    }

    pub fn by_type<'a>(&'a self, rel_type: &'a str) -> impl Iterator<Item = &'a Relationship> {
        self.entries.iter().filter(move |r| r.rel_type == rel_type)
    }

    pub fn by_id(&self, id: &str) -> Option<&Relationship> {
        self.id_index.get(id).map(|&i| &self.entries[i])
    }

    pub fn resolve_target(&self, rel: &Relationship, owner_part: &str) -> String {
        resolve_relative(owner_part, &rel.target)
    }

    pub fn resolve_id(&self, rel_id: &str, owner_part: &str) -> Option<String> {
        self.by_id(rel_id)
            .map(|rel| self.resolve_target(rel, owner_part))
    }
}

// ── RelationshipResolver ──────────────────────────────────────────────────────

#[derive(Debug, Default)]
pub struct RelationshipResolver {
    cache: HashMap<String, Relationships>,
}

impl RelationshipResolver {
    pub fn new() -> Self {
        Self::default()
    }

    pub fn resolve_target(
        &mut self,
        archive: &mut XlsxArchive,
        owner_part: &str,
        rel_id: &str,
    ) -> Result<String> {
        self.load(archive, owner_part)?;
        self.cache
            .get(owner_part)
            .unwrap()
            .resolve_id(rel_id, owner_part)
            .with_context(|| {
                format!("Relationship '{rel_id}' not found in rels for '{owner_part}'")
            })
    }

    pub fn relationships_for(
        &mut self,
        archive: &mut XlsxArchive,
        owner_part: &str,
    ) -> Result<&Relationships> {
        self.load(archive, owner_part)?;
        Ok(self.cache.get(owner_part).unwrap())
    }

    pub fn targets_of_type(
        &mut self,
        archive: &mut XlsxArchive,
        owner_part: &str,
        rel_type_uri: &str,
    ) -> Result<Vec<String>> {
        self.load(archive, owner_part)?;
        let resolved: Vec<String> = self
            .cache
            .get(owner_part)
            .unwrap()
            .by_type(rel_type_uri)
            .map(|r| resolve_relative(owner_part, &r.target))
            .collect();
        Ok(resolved)
    }

    pub fn follow_chain(
        &mut self,
        archive: &mut XlsxArchive,
        start: &str,
        steps: &[&str],
    ) -> Result<String> {
        let mut current = start.to_owned();
        for &rel_id in steps {
            current = self.resolve_target(archive, &current, rel_id)?;
        }
        Ok(current)
    }

    fn load(&mut self, archive: &mut XlsxArchive, owner_part: &str) -> Result<()> {
        if self.cache.contains_key(owner_part) {
            return Ok(());
        }
        let rels = parse_for_part(archive, owner_part)?;
        self.cache.insert(owner_part.to_owned(), rels);
        Ok(())
    }
}

// ── Free-standing parser functions ───────────────────────────────────────────

pub fn rels_path_for(part_path: &str) -> String {
    let path = Path::new(part_path);
    let dir = path.parent().unwrap_or_else(|| Path::new(""));
    let file_name = path.file_name().unwrap_or_default().to_string_lossy();

    let mut rels = PathBuf::from(dir);
    rels.push("_rels");
    rels.push(format!("{}.rels", file_name));

    rels.to_string_lossy().replace('\\', "/")
}

pub fn parse_for_part(archive: &mut XlsxArchive, part_path: &str) -> Result<Relationships> {
    let rels_path = rels_path_for(part_path);
    let xml = match read_entry_to_string(archive, &rels_path) {
        Ok(s) => s,
        Err(_) => return Ok(Relationships::default()),
    };
    parse_xml(&xml).with_context(|| format!("Failed to parse: {rels_path}"))
}

pub(crate) fn parse_xml(xml: &str) -> Result<Relationships> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(true);
    let mut rels = Relationships::default();

    loop {
        match reader.read_event()? {
            Event::Empty(ref e) | Event::Start(ref e) => {
                if e.local_name().as_ref() == b"Relationship" {
                    let decoder = reader.decoder();
                    if let Some(rel) = parse_relationship_element(e, decoder)? {
                        rels.push(rel);
                    }
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(rels)
}

// ── Path resolution helper ────────────────────────────────────────────────────

pub(crate) fn resolve_relative(owner_part: &str, target: &str) -> String {
    let base = Path::new(owner_part)
        .parent()
        .unwrap_or_else(|| Path::new(""));

    let mut components: Vec<&str> = base
        .to_str()
        .unwrap_or("")
        .split('/')
        .filter(|s| !s.is_empty())
        .collect();

    for segment in target.split('/') {
        match segment {
            "" | "." => {}
            ".." => {
                components.pop();
            }
            s => components.push(s),
        }
    }

    components.join("/")
}

// ── Attribute parsing helper ──────────────────────────────────────────────────

fn parse_relationship_element(
    e: &quick_xml::events::BytesStart<'_>,
    decoder: quick_xml::Decoder,
) -> Result<Option<Relationship>> {
    let mut id = String::new();
    let mut rel_type = String::new();
    let mut target = String::new();

    for attr in e.attributes() {
        let attr = attr.context("Malformed attribute in .rels file")?;
        match attr.key.local_name().as_ref() {
            b"Id" => id = attr.decode_and_unescape_value(decoder)?.into_owned(),
            b"Type" => rel_type = attr.decode_and_unescape_value(decoder)?.into_owned(),
            b"Target" => target = attr.decode_and_unescape_value(decoder)?.into_owned(),
            _ => {}
        }
    }

    if id.is_empty() || rel_type.is_empty() || target.is_empty() {
        return Ok(None);
    }

    Ok(Some(Relationship {
        id,
        rel_type,
        target,
    }))
}

// ── Tests ─────────────────────────────────────────────────────────────────────

#[cfg(test)]
mod tests {
    use super::*;

    const WORKBOOK_RELS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
    Target="sharedStrings.xml"/>
</Relationships>"#;

    const SHEET_RELS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
    Target="../drawings/drawing1.xml"/>
</Relationships>"#;

    const DRAWING_RELS_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
    Target="../charts/chart1.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
    Target="../charts/chart2.xml"/>
</Relationships>"#;

    #[test]
    fn rels_path_workbook() {
        assert_eq!(
            rels_path_for("xl/workbook.xml"),
            "xl/_rels/workbook.xml.rels"
        );
    }

    #[test]
    fn rels_path_worksheet() {
        assert_eq!(
            rels_path_for("xl/worksheets/sheet1.xml"),
            "xl/worksheets/_rels/sheet1.xml.rels"
        );
    }

    #[test]
    fn rels_path_drawing() {
        assert_eq!(
            rels_path_for("xl/drawings/drawing1.xml"),
            "xl/drawings/_rels/drawing1.xml.rels"
        );
    }

    #[test]
    fn parses_worksheet_relationships() {
        let rels = parse_xml(WORKBOOK_RELS_XML).unwrap();
        let sheets: Vec<_> = rels.by_type(rel_type::WORKSHEET).collect();
        assert_eq!(sheets.len(), 2);
        assert_eq!(sheets[0].id, "rId1");
        assert_eq!(sheets[1].target, "worksheets/sheet2.xml");
    }

    #[test]
    fn lookup_by_id_succeeds() {
        let rels = parse_xml(WORKBOOK_RELS_XML).unwrap();
        let rel = rels.by_id("rId2").expect("rId2 must exist");
        assert_eq!(rel.target, "worksheets/sheet2.xml");
    }

    #[test]
    fn lookup_by_id_missing_returns_none() {
        let rels = parse_xml(WORKBOOK_RELS_XML).unwrap();
        assert!(rels.by_id("rId99").is_none());
    }

    #[test]
    fn resolve_simple_relative_target() {
        assert_eq!(
            resolve_relative("xl/workbook.xml", "worksheets/sheet1.xml"),
            "xl/worksheets/sheet1.xml"
        );
    }

    #[test]
    fn resolve_dotdot_segment() {
        assert_eq!(
            resolve_relative("xl/drawings/drawing1.xml", "../charts/chart1.xml"),
            "xl/charts/chart1.xml"
        );
    }

    #[test]
    fn resolve_dotdot_from_worksheets() {
        assert_eq!(
            resolve_relative("xl/worksheets/sheet1.xml", "../drawings/drawing1.xml"),
            "xl/drawings/drawing1.xml"
        );
    }

    #[test]
    fn resolve_target_workbook_to_sheet() {
        let rels = parse_xml(WORKBOOK_RELS_XML).unwrap();
        let rel = rels.by_id("rId1").unwrap();
        assert_eq!(
            rels.resolve_target(rel, "xl/workbook.xml"),
            "xl/worksheets/sheet1.xml"
        );
    }

    #[test]
    fn resolve_target_sheet_to_drawing() {
        let rels = parse_xml(SHEET_RELS_XML).unwrap();
        let rel = rels.by_id("rId1").unwrap();
        assert_eq!(
            rels.resolve_target(rel, "xl/worksheets/sheet1.xml"),
            "xl/drawings/drawing1.xml"
        );
    }

    #[test]
    fn resolve_target_drawing_to_chart() {
        let rels = parse_xml(DRAWING_RELS_XML).unwrap();
        let rel = rels.by_id("rId1").unwrap();
        assert_eq!(
            rels.resolve_target(rel, "xl/drawings/drawing1.xml"),
            "xl/charts/chart1.xml"
        );
    }

    #[test]
    fn full_chain_workbook_to_chart() {
        let wb_rels = parse_xml(WORKBOOK_RELS_XML).unwrap();
        let sh_rels = parse_xml(SHEET_RELS_XML).unwrap();
        let dr_rels = parse_xml(DRAWING_RELS_XML).unwrap();

        let sheet_path = wb_rels.resolve_id("rId1", "xl/workbook.xml").unwrap();
        let drawing_path = sh_rels.resolve_id("rId1", &sheet_path).unwrap();
        let chart_path = dr_rels.resolve_id("rId1", &drawing_path).unwrap();

        assert_eq!(sheet_path, "xl/worksheets/sheet1.xml");
        assert_eq!(drawing_path, "xl/drawings/drawing1.xml");
        assert_eq!(chart_path, "xl/charts/chart1.xml");
    }

    #[test]
    fn drawing_with_two_charts() {
        let dr_rels = parse_xml(DRAWING_RELS_XML).unwrap();
        let owner = "xl/drawings/drawing1.xml";
        let chart_paths: Vec<String> = dr_rels
            .by_type(rel_type::CHART)
            .map(|r| dr_rels.resolve_target(r, owner))
            .collect();
        assert_eq!(
            chart_paths,
            vec!["xl/charts/chart1.xml", "xl/charts/chart2.xml"]
        );
    }
}
