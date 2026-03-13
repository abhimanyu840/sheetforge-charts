//! Streaming parser for `xl/theme/theme1.xml`.
//!
//! ## XML structure
//!
//! ```xml
//! <a:theme xmlns:a="…" name="Office Theme">
//!   <a:themeElements>
//!     <a:clrScheme name="Office">
//!       <a:dk1>
//!         <a:sysClr val="windowText" lastClr="000000"/>
//!       </a:dk1>
//!       <a:lt1>
//!         <a:sysClr val="window" lastClr="FFFFFF"/>
//!       </a:lt1>
//!       <a:dk2><a:srgbClr val="44546A"/></a:dk2>
//!       <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
//!       <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
//!       <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
//!       <a:accent3><a:srgbClr val="A9D18E"/></a:accent3>
//!       <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
//!       <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
//!       <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
//!       <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
//!       <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
//!     </a:clrScheme>
//!   </a:themeElements>
//! </a:theme>
//! ```

use anyhow::{Context, Result};
use quick_xml::{events::Event, Reader};

use crate::{
    archive::zip_reader::{read_entry_bytes, XlsxArchive},
    model::{
        color::{Rgb, ThemeColorSlot},
        theme::Theme,
    },
};

// ── Public entry points ───────────────────────────────────────────────────────

/// Parse the theme at `theme_path` from the archive.
pub fn parse(archive: &mut XlsxArchive, theme_path: &str) -> Result<Theme> {
    let bytes = read_entry_bytes(archive, theme_path)
        .with_context(|| format!("Cannot read theme part: {theme_path}"))?;
    parse_bytes(&bytes).with_context(|| format!("Failed to parse theme XML: {theme_path}"))
}

/// Parse from raw bytes — unit-test friendly.
pub(crate) fn parse_bytes(bytes: &[u8]) -> Result<Theme> {
    let mut reader = Reader::from_reader(bytes);
    reader.config_mut().trim_text(true);

    let mut theme = Theme::default();
    let mut current_slot: Option<ThemeColorSlot> = None;

    loop {
        match reader.read_event()? {
            Event::Start(ref e) | Event::Empty(ref e) => {
                let ln = e.local_name();
                let tag = std::str::from_utf8(ln.as_ref()).unwrap_or("");
                let dec = reader.decoder();

                match tag {
                    // Theme name
                    "theme" => {
                        if let Some(n) = attr(e, b"name", dec)? {
                            theme.name = Some(n);
                        }
                    }

                    // Color slot containers — one of the 12 named slots
                    t if ThemeColorSlot::from_str(t).is_some() => {
                        current_slot = ThemeColorSlot::from_str(t);
                    }

                    // srgbClr — direct hex value
                    "srgbClr" => {
                        if let Some(slot) = current_slot {
                            if let Some(v) = attr(e, b"val", dec)? {
                                if let Some(rgb) = Rgb::from_hex(&v) {
                                    theme.set(slot, rgb);
                                }
                            }
                        }
                    }

                    // sysClr — use lastClr as the resolved fallback
                    "sysClr" => {
                        if let Some(slot) = current_slot {
                            if let Some(v) = attr(e, b"lastClr", dec)? {
                                if let Some(rgb) = Rgb::from_hex(&v) {
                                    theme.set(slot, rgb);
                                }
                            }
                        }
                    }

                    _ => {}
                }
            }

            // When a slot element closes, clear the pending slot so that
            // nested srgbClr / sysClr inside a different element can't
            // accidentally overwrite it.
            Event::End(ref e) => {
                let ln = e.local_name();
                let tag = std::str::from_utf8(ln.as_ref()).unwrap_or("");
                if ThemeColorSlot::from_str(tag).is_some() {
                    current_slot = None;
                }
            }

            Event::Eof => break,
            _ => {}
        }
    }

    Ok(theme)
}

// ── Attribute helper ──────────────────────────────────────────────────────────

fn attr(
    e: &quick_xml::events::BytesStart<'_>,
    name: &[u8],
    dec: quick_xml::Decoder,
) -> Result<Option<String>> {
    for a in e.attributes() {
        let a = a.context("Malformed attribute in theme XML")?;
        if a.key.local_name().as_ref() == name {
            return Ok(Some(a.decode_and_unescape_value(dec)?.into_owned()));
        }
    }
    Ok(None)
}

// ── Tests ─────────────────────────────────────────────────────────────────────

#[cfg(test)]
pub(crate) mod tests {
    use super::*;
    use crate::model::color::{Rgb, ThemeColorSlot};

    pub const OFFICE_THEME_XML: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window"     lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
      <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
      <a:accent3><a:srgbClr val="A9D18E"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
  </a:themeElements>
</a:theme>"#;

    fn parse_office() -> Theme {
        parse_bytes(OFFICE_THEME_XML.as_bytes()).unwrap()
    }

    #[test]
    fn theme_name_parsed() {
        assert_eq!(parse_office().name.as_deref(), Some("Office Theme"));
    }

    #[test]
    fn dk1_from_sysclr_lastclr() {
        assert_eq!(parse_office().dk1(), Some(Rgb::BLACK));
    }

    #[test]
    fn lt1_from_sysclr_lastclr() {
        assert_eq!(parse_office().lt1(), Some(Rgb::WHITE));
    }

    #[test]
    fn accent1_blue() {
        assert_eq!(
            parse_office().accent1(),
            Some(Rgb::from_hex("4472C4").unwrap())
        );
    }

    #[test]
    fn accent2_orange() {
        assert_eq!(
            parse_office().accent2(),
            Some(Rgb::from_hex("ED7D31").unwrap())
        );
    }

    #[test]
    fn accent3_green() {
        assert_eq!(
            parse_office().accent3(),
            Some(Rgb::from_hex("A9D18E").unwrap())
        );
    }

    #[test]
    fn accent4_yellow() {
        assert_eq!(
            parse_office().accent4(),
            Some(Rgb::from_hex("FFC000").unwrap())
        );
    }

    #[test]
    fn accent5_light_blue() {
        assert_eq!(
            parse_office().accent5(),
            Some(Rgb::from_hex("5B9BD5").unwrap())
        );
    }

    #[test]
    fn accent6_light_green() {
        assert_eq!(
            parse_office().accent6(),
            Some(Rgb::from_hex("70AD47").unwrap())
        );
    }

    #[test]
    fn dk2_parsed() {
        assert_eq!(
            parse_office().color(ThemeColorSlot::Dk2),
            Some(Rgb::from_hex("44546A").unwrap())
        );
    }

    #[test]
    fn all_12_slots_populated() {
        let t = parse_office();
        assert_eq!(t.all_colors().len(), 12);
    }

    #[test]
    fn empty_xml_returns_empty_theme() {
        let t = parse_bytes(b"<a:theme/>").unwrap();
        assert!(t.is_empty());
    }
}
