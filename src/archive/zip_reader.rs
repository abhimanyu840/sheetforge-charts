//! ZIP archive helpers.

use std::{
    fs::File,
    io::{BufReader, Read},
    path::Path,
};

use anyhow::{Context, Result};
use zip::ZipArchive;

pub type XlsxArchive = ZipArchive<BufReader<File>>;

pub fn open_xlsx(path: &str) -> Result<XlsxArchive> {
    let path = Path::new(path);
    let file = File::open(path).with_context(|| format!("Cannot open file: {}", path.display()))?;
    ZipArchive::new(BufReader::new(file))
        .with_context(|| format!("File is not a valid ZIP / XLSX archive: {}", path.display()))
}

/// Read a ZIP entry into a `String`.  Used for small XML parts.
pub fn read_entry_to_string(archive: &mut XlsxArchive, entry_path: &str) -> Result<String> {
    let mut entry = archive
        .by_name(entry_path)
        .with_context(|| format!("Entry not found in archive: {entry_path}"))?;
    let capacity = entry.size() as usize;
    let mut buf = String::with_capacity(capacity);
    entry
        .read_to_string(&mut buf)
        .with_context(|| format!("Entry is not valid UTF-8: {entry_path}"))?;
    Ok(buf)
}

/// Read a ZIP entry into a `Vec<u8>`.
///
/// Preferred for chart XML parts because `quick_xml::Reader` can work directly
/// on `&[u8]` without any UTF-8 validation overhead — the parser handles
/// encoding itself.
pub fn read_entry_bytes(archive: &mut XlsxArchive, entry_path: &str) -> Result<Vec<u8>> {
    let mut entry = archive
        .by_name(entry_path)
        .with_context(|| format!("Entry not found in archive: {entry_path}"))?;
    let capacity = entry.size() as usize;
    let mut buf = Vec::with_capacity(capacity);
    entry
        .read_to_end(&mut buf)
        .with_context(|| format!("Cannot read entry: {entry_path}"))?;
    Ok(buf)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn open_nonexistent_file_returns_error() {
        let result = open_xlsx("does_not_exist.xlsx");
        assert!(result.is_err());
        let msg = format!("{}", result.unwrap_err());
        assert!(msg.contains("Cannot open file"));
    }
}
