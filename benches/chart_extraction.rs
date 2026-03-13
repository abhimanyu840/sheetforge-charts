//! Criterion benchmarks for chart extraction throughput.
//!
//! ## Workloads
//!
//! | Benchmark            | Charts | Purpose                              |
//! |----------------------|--------|--------------------------------------|
//! | `extract_1_chart`    | 1      | baseline / single-thread cost        |
//! | `extract_10_charts`  | 10     | light multi-chart file               |
//! | `extract_50_charts`  | 50     | representative real-world workbook   |
//! | `extract_100_charts` | 100    | stress test / parallelism ceiling    |
//!
//! ## Running
//!
//! ```sh
//! cargo bench
//! # or with HTML report:
//! cargo bench -- --output-format html
//! ```
//!
//! ## What is measured
//!
//! Each iteration calls `extract_charts(path)` end-to-end: file open, ZIP
//! decompression, relationship walking, theme parsing, and parallel chart-XML
//! parsing.  The fixture file is created once per benchmark group and written
//! to a `tempfile`; Criterion's iteration loop re-opens it each time so the OS
//! page cache influence is realistic (files stay warm after the first sample).

use std::io::Write;

use criterion::{criterion_group, criterion_main, BenchmarkId, Criterion, Throughput};
use sheetforge_charts::extract_charts;

// ── Fixture generation ────────────────────────────────────────────────────────

/// Build a minimal but structurally valid `.xlsx` ZIP in memory.
///
/// Generates `n_charts` charts spread across `n_charts` sheets (one per
/// sheet — the worst case for relationship walking).  Each chart is a bar
/// chart with two series and a three-point numeric cache, matching the
/// structure of a real Excel file.
fn build_xlsx_bytes(n_charts: usize) -> Vec<u8> {
    use std::io::Cursor;
    use zip::{write::SimpleFileOptions, ZipWriter};

    let buf = Cursor::new(Vec::<u8>::new());
    let mut zw = ZipWriter::new(buf);
    let opts = SimpleFileOptions::default();

    // ── [Content_Types].xml ───────────────────────────────────────────────
    let mut ct = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml"
    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
"#,
    );
    for i in 1..=n_charts {
        ct.push_str(&format!(
            "  <Override PartName=\"/xl/worksheets/sheet{i}.xml\" \
             ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\n"
        ));
        ct.push_str(&format!(
            "  <Override PartName=\"/xl/drawings/drawing{i}.xml\" \
             ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/>\n"
        ));
        ct.push_str(&format!(
            "  <Override PartName=\"/xl/charts/chart{i}.xml\" \
             ContentType=\"application/vnd.openxmlformats-officedocument.drawingml.chart+xml\"/>\n"
        ));
    }
    ct.push_str("</Types>");
    zw.start_file("[Content_Types].xml", opts).unwrap();
    zw.write_all(ct.as_bytes()).unwrap();

    // ── _rels/.rels ───────────────────────────────────────────────────────
    zw.start_file("_rels/.rels", opts).unwrap();
    zw.write_all(
        br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="xl/workbook.xml"/>
</Relationships>"#,
    )
    .unwrap();

    // ── xl/workbook.xml ───────────────────────────────────────────────────
    let mut wb = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
"#,
    );
    for i in 1..=n_charts {
        wb.push_str(&format!(
            "    <sheet name=\"Sheet{i}\" sheetId=\"{i}\" r:id=\"rId{i}\"/>\n"
        ));
    }
    wb.push_str("  </sheets>\n</workbook>");
    zw.start_file("xl/workbook.xml", opts).unwrap();
    zw.write_all(wb.as_bytes()).unwrap();

    // ── xl/_rels/workbook.xml.rels ────────────────────────────────────────
    let mut wb_rels = String::from(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
"#,
    );
    for i in 1..=n_charts {
        wb_rels.push_str(&format!(
            "  <Relationship Id=\"rId{i}\" \
             Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" \
             Target=\"worksheets/sheet{i}.xml\"/>\n"
        ));
    }
    wb_rels.push_str("</Relationships>");
    zw.start_file("xl/_rels/workbook.xml.rels", opts).unwrap();
    zw.write_all(wb_rels.as_bytes()).unwrap();

    // ── Per-chart parts ───────────────────────────────────────────────────
    for i in 1..=n_charts {
        let sheet_name = format!("Sheet{i}");

        // xl/worksheets/sheetN.xml  (empty — we only care about rels)
        zw.start_file(format!("xl/worksheets/sheet{i}.xml"), opts)
            .unwrap();
        zw.write_all(
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#,
        )
        .unwrap();

        // xl/worksheets/_rels/sheetN.xml.rels
        let sh_rels = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
    Target="../drawings/drawing{i}.xml"/>
</Relationships>"#
        );
        zw.start_file(format!("xl/worksheets/_rels/sheet{i}.xml.rels"), opts)
            .unwrap();
        zw.write_all(sh_rels.as_bytes()).unwrap();

        // xl/drawings/drawingN.xml
        let drawing = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr
  xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <xdr:twoCellAnchor>
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>8</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>15</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:graphicFrame>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"#
        );
        zw.start_file(format!("xl/drawings/drawing{i}.xml"), opts)
            .unwrap();
        zw.write_all(drawing.as_bytes()).unwrap();

        // xl/drawings/_rels/drawingN.xml.rels
        let dr_rels = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
    Target="../charts/chart{i}.xml"/>
</Relationships>"#
        );
        zw.start_file(format!("xl/drawings/_rels/drawing{i}.xml.rels"), opts)
            .unwrap();
        zw.write_all(dr_rels.as_bytes()).unwrap();

        // xl/charts/chartN.xml  — a realistic bar chart with 2 series × 3 pts
        let chart = format!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/>
      <a:p><a:r><a:t>Chart {i}</a:t></a:r></a:p>
    </c:rich></c:tx><c:overlay val="0"/></c:title>
    <c:autoTitleDeleted val="0"/>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:tx><c:strRef><c:f>{sheet_name}!$B$1</c:f></c:strRef></c:tx>
          <c:spPr>
            <a:solidFill><a:srgbClr val="4472C4"/></a:solidFill>
          </c:spPr>
          <c:cat>
            <c:strRef>
              <c:f>{sheet_name}!$A$2:$A$4</c:f>
              <c:strCache>
                <c:ptCount val="3"/>
                <c:pt idx="0"><c:v>Jan</c:v></c:pt>
                <c:pt idx="1"><c:v>Feb</c:v></c:pt>
                <c:pt idx="2"><c:v>Mar</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:f>{sheet_name}!$B$2:$B$4</c:f>
              <c:numCache>
                <c:formatCode>General</c:formatCode>
                <c:ptCount val="3"/>
                <c:pt idx="0"><c:v>1000</c:v></c:pt>
                <c:pt idx="1"><c:v>1500</c:v></c:pt>
                <c:pt idx="2"><c:v>1200</c:v></c:pt>
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
        <c:ser>
          <c:idx val="1"/><c:order val="1"/>
          <c:tx><c:strRef><c:f>{sheet_name}!$C$1</c:f></c:strRef></c:tx>
          <c:spPr>
            <a:gradFill>
              <a:gsLst>
                <a:gs pos="0"><a:srgbClr val="ED7D31"/></a:gs>
                <a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs>
              </a:gsLst>
              <a:lin ang="5400000" scaled="0"/>
            </a:gradFill>
          </c:spPr>
          <c:cat>
            <c:strRef>
              <c:f>{sheet_name}!$A$2:$A$4</c:f>
              <c:strCache>
                <c:ptCount val="3"/>
                <c:pt idx="0"><c:v>Jan</c:v></c:pt>
                <c:pt idx="1"><c:v>Feb</c:v></c:pt>
                <c:pt idx="2"><c:v>Mar</c:v></c:pt>
              </c:strCache>
            </c:strRef>
          </c:cat>
          <c:val>
            <c:numRef>
              <c:f>{sheet_name}!$C$2:$C$4</c:f>
              <c:numCache>
                <c:formatCode>General</c:formatCode>
                <c:ptCount val="3"/>
                <c:pt idx="0"><c:v>50</c:v></c:pt>
                <c:pt idx="1"><c:v>75</c:v></c:pt>
                <c:pt idx="2"><c:v>60</c:v></c:pt>
              </c:numCache>
            </c:numRef>
          </c:val>
        </c:ser>
        <c:axId val="1"/>
        <c:axId val="2"/>
      </c:barChart>
      <c:catAx>
        <c:axId val="1"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:numFmt formatCode="General" sourceLinked="0"/>
        <c:crossAx val="2"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="2"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="l"/>
        <c:numFmt formatCode="General" sourceLinked="0"/>
        <c:crossAx val="1"/>
      </c:valAx>
    </c:plotArea>
    <c:legend><c:legendPos val="b"/></c:legend>
  </c:chart>
  <c:style val="2"/>
</c:chartSpace>"#
        );
        zw.start_file(format!("xl/charts/chart{i}.xml"), opts)
            .unwrap();
        zw.write_all(chart.as_bytes()).unwrap();
    }

    let cursor = zw.finish().unwrap();
    cursor.into_inner()
}

/// Write XLSX bytes to a tempfile, return the path (file stays alive via handle).
fn write_temp_xlsx(bytes: &[u8]) -> (tempfile::NamedTempFile, String) {
    let mut tmp = tempfile::NamedTempFile::new().unwrap();
    tmp.write_all(bytes).unwrap();
    let path = tmp.path().to_string_lossy().into_owned();
    (tmp, path)
}

// ── Benchmark groups ──────────────────────────────────────────────────────────

fn bench_extract(c: &mut Criterion) {
    let sizes: &[usize] = &[1, 10, 50, 100];

    let mut group = c.benchmark_group("extract_charts");

    for &n in sizes {
        // Build the fixture once per size; hold the NamedTempFile so it isn't
        // deleted while the benchmark loop runs.
        let bytes = build_xlsx_bytes(n);
        let (_tmp, path) = write_temp_xlsx(&bytes);

        group.throughput(Throughput::Elements(n as u64));
        group.bench_with_input(BenchmarkId::from_parameter(n), &path, |b, path| {
            b.iter(|| {
                let wb = extract_charts(path).expect("extraction failed in bench");
                // Black-box the result so the compiler cannot optimise away
                // the call.
                criterion::black_box(wb.chart_count())
            });
        });
    }

    group.finish();
}

/// Separate group that isolates pure XML parse time (no ZIP I/O).
///
/// Generates the raw chart XML bytes and calls `parse_bytes` directly so
/// we can see how much time is spent inside the parser vs. I/O.
fn bench_parse_bytes(c: &mut Criterion) {
    use sheetforge_charts::parser::chart_parser::parse_bytes;

    let chart_xml = build_xlsx_bytes(1); // use fixture generator for valid XML
                                         // Extract just one chart's XML — build it directly as a string instead
    let xml = build_single_chart_xml(1);

    let mut group = c.benchmark_group("parse_bytes");

    for &n in &[1usize, 10, 50, 100] {
        // n copies in a Vec simulates a parallel workload
        let xmls: Vec<Vec<u8>> = (1..=n)
            .map(|i| build_single_chart_xml(i).into_bytes())
            .collect();

        group.throughput(Throughput::Elements(n as u64));
        group.bench_with_input(BenchmarkId::from_parameter(n), &xmls, |b, xmls| {
            b.iter(|| {
                let parsed: Vec<_> = xmls
                    .iter()
                    .map(|bytes| parse_bytes(bytes, "xl/charts/chart.xml").unwrap())
                    .collect();
                criterion::black_box(parsed.len())
            });
        });
    }

    group.finish();
    let _ = xml; // suppress unused warning
    let _ = chart_xml;
}

fn build_single_chart_xml(i: usize) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart>
    <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/>
      <a:p><a:r><a:t>Chart {i}</a:t></a:r></a:p>
    </c:rich></c:tx><c:overlay val="0"/></c:title>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/><c:order val="0"/>
          <c:spPr><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></c:spPr>
          <c:cat><c:strRef><c:f>Sheet{i}!$A$2:$A$4</c:f>
            <c:strCache><c:ptCount val="3"/>
              <c:pt idx="0"><c:v>Jan</c:v></c:pt>
              <c:pt idx="1"><c:v>Feb</c:v></c:pt>
              <c:pt idx="2"><c:v>Mar</c:v></c:pt>
            </c:strCache>
          </c:strRef></c:cat>
          <c:val><c:numRef><c:f>Sheet{i}!$B$2:$B$4</c:f>
            <c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/>
              <c:pt idx="0"><c:v>1000</c:v></c:pt>
              <c:pt idx="1"><c:v>1500</c:v></c:pt>
              <c:pt idx="2"><c:v>1200</c:v></c:pt>
            </c:numCache>
          </c:numRef></c:val>
        </c:ser>
        <c:axId val="1"/>
        <c:axId val="2"/>
      </c:barChart>
      <c:catAx>
        <c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/>
      </c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>"#
    )
}

criterion_group!(benches, bench_extract, bench_parse_bytes);
criterion_main!(benches);
