# xlsx-chart-extractor

A high-performance Rust library for extracting chart metadata from Excel (`.xlsx`) files.  
Designed as the native core for Python (PyO3) and TypeScript/JavaScript (napi-rs) bindings.

---

## Project status

| Phase | Scope | Status |
|-------|-------|--------|
| **1** | Rust foundation — archive reading, content-types, workbook discovery | ✅ Complete |
| **2** | Chart XML parsing — type, title, series, axes | 🔜 Next |
| **3** | Python bindings (PyO3 + maturin) | 🔜 Planned |
| **4** | TypeScript/Node.js bindings (napi-rs) | 🔜 Planned |

---

## Architecture

```
src/
├── lib.rs                     ← public API: extract_charts()
├── archive/
│   └── zip_reader.rs          ← open XLSX as ZIP, read entries
├── openxml/
│   ├── content_types.rs       ← parse [Content_Types].xml
│   └── relationships.rs       ← parse .rels files
├── model/
│   ├── workbook.rs            ← WorkbookCharts, SheetCharts
│   ├── chart.rs               ← Chart, ChartType
│   ├── series.rs              ← Series, DataReference, DataValues
│   └── axis.rs                ← Axis, AxisType, AxisPosition
└── parser/
    └── workbook_parser.rs     ← parse workbook.xml → sheet list
```

### Data model

```
WorkbookCharts
  source_path: String
  sheets: Vec<SheetCharts>
    SheetCharts
      name: String
      relationship_id: String
      index: usize
      charts: Vec<Chart>
        Chart
          chart_path: String
          chart_type: ChartType
          title: Option<String>
          series: Vec<Series>
            Series
              index: u32
              name: Option<String>
              category_ref / value_ref: Option<DataReference>
              value_cache: Option<DataValues>
          axes: Vec<Axis>
            Axis
              id: u32
              axis_type: AxisType
              title / position / number_format
```

---

## Quick start

```rust
use xlsx_chart_extractor::extract_charts;

fn main() -> anyhow::Result<()> {
    let workbook = extract_charts("path/to/file.xlsx")?;

    println!("Sheets: {}", workbook.sheets.len());
    for sheet in &workbook.sheets {
        println!("  {} — {} chart(s)", sheet.name, sheet.charts.len());
        for chart in &sheet.charts {
            println!("    {:?} — {:?}", chart.chart_type, chart.title);
        }
    }
    Ok(())
}
```

---

## Build

```bash
# Standard library build
cargo build --release

# Run all tests (including unit tests inside each module)
cargo test

# Check for compile errors without linking
cargo check
```

---

## Future: Python bindings (Phase 3)

```bash
pip install maturin
maturin develop --features python
```

```python
import xlsx_chart_extractor
wb = xlsx_chart_extractor.extract_charts("file.xlsx")
print(wb.sheets)
```

## Future: Node.js bindings (Phase 4)

```bash
npm install -g @napi-rs/cli
napi build --features nodejs --release
```

```typescript
import { extractCharts } from './index.js'
const wb = extractCharts('file.xlsx')
console.log(wb.sheets)
```

---

## Dependencies

| Crate | Purpose |
|-------|---------|
| `zip` | Read `.xlsx` as a ZIP archive |
| `quick-xml` | Streaming, zero-copy XML parsing |
| `serde` + `serde_json` | Serialise the model to JSON for bindings |
| `anyhow` | Ergonomic, context-rich error handling |
