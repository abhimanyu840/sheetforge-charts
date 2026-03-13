# sheetforge-charts

A high-performance Rust library for extracting chart metadata from Excel (`.xlsx`) files.  
Designed as the native core for Python (PyO3) and TypeScript/Node.js (napi-rs) bindings.

---

## Project status

| Phase | Scope | Status |
|-------|-------|--------|
| **1** | ZIP archive reading, content-types, workbook/sheet discovery | ✅ Complete |
| **2** | Relationship chain resolver, drawing parser, sheet→chart linking | ✅ Complete |
| **3** | Chart XML streaming parser — type, title, legend, style, grouping, series refs+caches, axes | ✅ Complete |
| **4** | Cache extraction — sparse `pt idx`, `ptCount` pre-alloc, slot-aware routing | ✅ Complete |
| **5** | Theme parsing, color/gradient fill extraction, `schemeClr` resolution | ✅ Complete |
| **6** | Drawing anchor parsing — `twoCellAnchor` → `ChartAnchor` on `Chart` | ✅ Complete |
| **7** | Performance — Rayon parallel chart parsing + Criterion benchmarks | ✅ Complete |
| **8** | Full 3-D chart support — `Chart3DView`, all 3-D `ChartType` variants, `<c:view3D>` parser | ✅ Complete |
| **9** | 3-D geometry surfaces — `Chart3DSurface`, `<c:floor>`, `<c:sideWall>`, `<c:backWall>` fills | ✅ Complete |
| **10** | Python bindings (PyO3 + maturin) | 🔜 Next |
| **11** | TypeScript/Node.js bindings (napi-rs) | 🔜 Planned |
| **12** | Resolve worksheet cell ranges to actual data values | 🔜 Planned |

---

## Architecture

```
src/
├── lib.rs                          ← public API: extract_charts()
│                                     two-phase pipeline: serial I/O → parallel parse
├── bin/
│   └── inspect.rs                  ← CLI tool: dump all metadata for any .xlsx
├── archive/
│   └── zip_reader.rs               ← open XLSX as ZIP, stream entries
├── openxml/
│   ├── content_types.rs            ← parse [Content_Types].xml
│   ├── relationships.rs            ← parse .rels files, resolve chains
│   ├── drawing.rs                  ← parse xl/drawings/drawingN.xml → ChartAnchor
│   └── theme_parser.rs             ← parse xl/theme/theme1.xml → Theme
├── model/
│   ├── workbook.rs                 ← WorkbookCharts, SheetCharts
│   ├── chart.rs                    ← Chart, ChartType, PlotArea, ChartAnchor,
│   │                                  Chart3DView, Chart3DSurface
│   ├── color.rs                    ← Rgb, ColorSpec, Fill, Gradient, ColorMod,
│   │                                  ThemeColorSlot
│   ├── series.rs                   ← Series, DataReference, DataValues, StringValues
│   ├── axis.rs                     ← Axis, AxisType, AxisPosition
│   └── theme.rs                    ← Theme (slot → Rgb map)
└── parser/
    ├── workbook_parser.rs          ← parse workbook.xml → sheet list
    ├── sheet_parser.rs             ← walk sheet/drawing/chart rel chains
    └── chart_parser.rs             ← streaming chartN.xml state machine
```

### Data model

```
WorkbookCharts
  source_path : String
  theme       : Option<Theme>
  sheets      : Vec<SheetCharts>
    name            : String
    index           : usize
    charts          : Vec<Chart>
      chart_path      : String
      chart_type      : ChartType          // Bar, Line3D, Surface3D, …
      title           : Option<String>
      legend_position : Option<LegendPosition>
      style           : Option<u32>
      chart_fill      : Option<Fill>
      anchor          : Option<ChartAnchor>    // twoCellAnchor col/row bounds
      view_3d         : Option<Chart3DView>    // rotX, rotY, rAngAx, perspective
      surface         : Option<Chart3DSurface> // floor/sideWall/backWall fills
      plot_area       : PlotArea
        chart_type    : ChartType
        grouping      : Option<Grouping>
        fill          : Option<Fill>
        series        : Vec<Series>
          index            : u32
          name             : Option<String>
          name_ref         : Option<DataReference>
          category_ref     : Option<DataReference>
          value_ref        : Option<DataReference>
          value_cache      : Option<DataValues>    // embedded numeric cache
          category_values  : Option<StringValues>  // embedded label cache
          fill             : Option<Fill>
        axes          : Vec<Axis>
          id, axis_type, position, cross_axis_id, title, number_format
```

---

## Quick start

```rust
use sheetforge_charts::extract_charts;

fn main() -> anyhow::Result<()> {
    let wb = extract_charts("path/to/file.xlsx")?;

    if let Some(theme) = &wb.theme {
        println!("accent1 = {:?}", theme.accent1());
    }

    for sheet in &wb.sheets {
        for chart in &sheet.charts {
            println!("{}: {:?}", chart.chart_type, chart.title);

            if let Some(v) = &chart.view_3d {
                println!("  3D — rotX={:?} rotY={:?}", v.rotation_x, v.rotation_y);
            }
            if let Some(s) = &chart.surface {
                println!("  floor={:?}", s.floor_fill);
            }
        }
    }
    Ok(())
}
```

---

## Build & test

```bash
# Build
cargo build --release

# Run all tests (unit + integration)
cargo test

# Inspect any .xlsx file
cargo run --bin inspect -- path/to/file.xlsx

# Benchmarks
cargo bench
```

---

## Color & theme resolution

`Fill` values may contain `ColorSpec::Scheme(slot, mods)` references that require the workbook theme to resolve to a concrete `Rgb`. Pass `wb.theme.as_ref()` to any resolver:

```rust
use sheetforge_charts::model::color::Fill;

if let Some(fill) = &chart.chart_fill {
    if let Some(rgb) = fill.solid_rgb(wb.theme.as_ref()) {
        println!("chart background = #{}", rgb.to_hex());
    }
}
```

---

## Performance design

The pipeline is deliberately split into two phases:

**Phase A — Serial I/O**  
All ZIP reads happen on a single thread. `ZipArchive` is not `Send`, so the archive cannot be shared. Every chart's raw XML bytes are loaded into memory here.

**Phase B — Parallel parse**  
Raw bytes are handed to Rayon's `par_iter`. Each worker calls `chart_parser::parse_bytes` independently — pure CPU, no I/O, no shared state. On a 50-chart workbook, Phase B completes in ≈ `1/N_CORES` of serial time.

---

## Roadmap

**Phase 10 — Python bindings**
```bash
pip install maturin
maturin develop --features python
```
```python
import sheetforge_charts
wb = sheetforge_charts.extract_charts("file.xlsx")
```

**Phase 11 — Node.js/TypeScript bindings**
```bash
npm install -g @napi-rs/cli
napi build --features nodejs --release
```
```typescript
import { extractCharts } from './index.js'
const wb = extractCharts('file.xlsx')
```

---

## Dependencies

| Crate | Purpose |
|-------|---------|
| `zip` | Read `.xlsx` as a ZIP archive |
| `quick-xml` | Streaming, zero-copy XML parsing |
| `serde` + `serde_json` | Serialise the model to JSON for language bindings |
| `anyhow` | Context-rich error handling |
| `rayon` | Data-parallel chart parsing |
| `criterion` *(dev)* | Micro-benchmark harness |
