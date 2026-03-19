<div align="center">

<img src="https://img.shields.io/badge/sheetforge--charts-v0.1.0-f59e0b?style=for-the-badge&labelColor=0d0d12" alt="version"/>
<img src="https://img.shields.io/badge/python-3.8%2B-60a5fa?style=for-the-badge&labelColor=0d0d12" alt="python"/>
<img src="https://img.shields.io/badge/built%20with-Rust%20%F0%9F%A6%80-e85d04?style=for-the-badge&labelColor=0d0d12" alt="rust"/>
<img src="https://img.shields.io/badge/license-MIT-34d399?style=for-the-badge&labelColor=0d0d12" alt="license"/>
<img src="https://img.shields.io/badge/tests-656%20passing-34d399?style=for-the-badge&labelColor=0d0d12" alt="tests"/>

<br/><br/>

# sheetforge-charts

**Extract chart metadata, series data, theme colors, and placement from any Excel `.xlsx` file — in milliseconds.**

No Excel. No LibreOffice. No Java. Just a fast Rust core with clean Python bindings.

[**Quick Start**](#-quick-start) · [**API Reference**](#-api-reference) · [**Color System**](#-color-system) · [**Rendering**](#-rendering-charts) · [**Install**](#-installation)

</div>

---

## What does it do?

You hand it an Excel file. It gives you back everything you need to recreate every chart programmatically:

```python
import sheetforge_charts as sc

wb    = sc.extract_charts("sales_report.xlsx")
theme = wb.theme

for sheet in wb.sheets:
    for chart in sheet.charts:
        print(f"[{sheet.name}] {chart.chart_type}: {chart.title}")
        # [Sales] Bar: Sales Overview
        # [Expenses] Line: Monthly Expenses

        print(f"  position: {chart.position.top_left}:{chart.position.bottom_right}")
        # position: B2:J17

        for s in chart.series:
            # Resolve the exact color Excel uses for this series
            fill  = s.fill(theme)
            color = fill.color if fill else theme.color(f"accent{(s.index % 6) + 1}")
            print(f"  {s.name}: {s.values} | color={color}")
            # Revenue: [1000.0, 1500.0, 1200.0] | color=#4472C4
```

---

## ✨ Features

| Feature | Details |
|---|---|
| **15+ chart types** | Bar, Line, Pie, Area, Scatter, Bubble, Radar, Stock, 3-D variants, Combo |
| **Full color resolution** | Resolves theme accents, sRGB hex, gradients, and DrawingML modifiers (lumMod, tint, shade) |
| **Combo chart layers** | Multi-layer charts split into `ChartLayer` objects — each with its own type and series |
| **Secondary axis** | `series.is_secondary_axis` flag + `series.axis_id` for right/top axis detection |
| **Chart placement** | A1-notation cell range (`B2:J17`) and EMU dimensions from both anchor types |
| **Pivot chart metadata** | Full relationship chain: chart → pivotTable → pivotCacheDefinition → field names |
| **3-D chart support** | `view_3d` rotation/perspective + floor/sideWall/backWall surface fills |
| **Parallel parsing** | Rayon workers — 50-chart workbook parses in ≈1/N_CORES of serial time |
| **Single wheel** | `abi3-py38` — one `.whl` works on Python 3.8 through 3.14 |

---

## 📦 Installation

```bash
pip install sheetforge-charts
```

**Requirements:** Python ≥ 3.8. The wheel ships a pre-compiled Rust binary — no Rust toolchain needed.

<details>
<summary>Install in a virtual environment (recommended)</summary>

```bash
python -m venv .venv
source .venv/bin/activate        # macOS / Linux
source .venv/Scripts/activate    # Windows (MINGW64 / PowerShell)
pip install sheetforge-charts
```

</details>

<details>
<summary>Build from source</summary>

```bash
git clone https://github.com/your-org/sheetforge-charts
cd sheetforge-charts
python -m venv .venv && source .venv/Scripts/activate
pip install maturin
PYO3_USE_ABI3_FORWARD_COMPATIBILITY=1 python -m maturin develop --features python
```

> **Note:** On Python 3.14 you must prefix maturin commands with `PYO3_USE_ABI3_FORWARD_COMPATIBILITY=1` until PyO3 0.22 ships with native 3.14 support.

</details>

---

## 🚀 Quick Start

```python
import sheetforge_charts as sc

wb    = sc.extract_charts("your_file.xlsx")
theme = wb.theme

print(f"Found {wb.chart_count} charts across {len(wb.sheets)} sheets")

for sheet in wb.sheets:
    for chart in sheet.charts:

        # ── Basic info ────────────────────────────────────────────────
        print(chart.chart_type)       # "Bar", "Line", "Pie", "Combo" …
        print(chart.title)            # "Sales Overview" or None
        print(chart.style)            # Excel style index 1–48

        # ── Worksheet placement ───────────────────────────────────────
        if chart.position:
            p = chart.position
            print(p.sheet)            # "Sales"
            print(p.top_left)         # "B2"
            print(p.bottom_right)     # "J17"

        # ── Combo charts: per-layer breakdown ─────────────────────────
        for layer in chart.layers:
            print(layer.chart_type)   # "Bar" / "Line" / …
            print(layer.grouping)     # "Clustered" / "Stacked" / …

        # ── Series: data + colors ─────────────────────────────────────
        for s in chart.series:
            print(s.name)             # "Revenue"
            print(s.values)           # [1000.0, 1500.0, 1200.0]
            print(s.categories)       # ["Jan", "Feb", "Mar"]
            print(s.value_ref)        # "Sheet1!$B$2:$B$4"
            print(s.is_secondary_axis)# False

            # Exact color used in Excel:
            fill = s.fill(theme)
            if fill and fill.fill_type == "solid":
                print(fill.color)     # "#4472C4"
                print(fill.color_raw) # "theme:accent1"
            else:
                # No explicit fill → Excel cycles through accent1–6
                slot  = f"accent{(s.index % 6) + 1}"
                color = theme.color(slot)
                print(color)          # "#4472C4"
```

---

## 📖 API Reference

### `extract_charts(path: str) → WorkbookCharts`

Entry point. Reads and parses all chart metadata from a `.xlsx` file.

```python
wb = sc.extract_charts("report.xlsx")
```

Raises `ValueError` if the file cannot be read or parsed.

---

### `WorkbookCharts`

| Property | Type | Description |
|---|---|---|
| `source_path` | `str` | Absolute path to the parsed file |
| `chart_count` | `int` | Total charts across all sheets |
| `sheets` | `list[SheetCharts]` | In Excel tab order |
| `theme` | `Theme \| None` | Workbook color theme |

---

### `SheetCharts`

| Property | Type | Description |
|---|---|---|
| `name` | `str` | Sheet display name |
| `index` | `int` | 0-based tab order index |
| `charts` | `list[Chart]` | All charts on this sheet |

---

### `Chart`

| Property | Type | Description |
|---|---|---|
| `chart_type` | `str` | `"Bar"` `"Line"` `"Pie"` `"Area"` `"Combo"` `"Bar3D"` … |
| `title` | `str \| None` | Chart title text |
| `layers` | `list[ChartLayer]` | 1 for single-type, 2+ for combo |
| `series` | `list[Series]` | All series (flat, all layers combined) |
| `position` | `ChartPosition \| None` | A1-notation cell range |
| `anchor` | `ChartAnchor \| None` | Raw 0-based row/col offsets |
| `style` | `int \| None` | Excel style index (1–48) |
| `is_pivot_chart` | `bool` | `True` when backed by a PivotTable |
| `pivot_table_name` | `str \| None` | e.g. `"Sheet1!PivotTable1"` |

**Methods:**

```python
chart.chart_fill(theme)      # → Fill | None — chart-space background
chart.plot_area_fill(theme)  # → Fill | None — plot area background
```

---

### `ChartLayer`

One chart-type element inside `<c:plotArea>`. Combo charts have 2+ layers.

| Property | Type | Description |
|---|---|---|
| `chart_type` | `str` | `"Bar"` `"Line"` `"Area"` … |
| `grouping` | `str \| None` | `"Clustered"` `"Stacked"` `"Standard"` … |
| `bar_horizontal` | `bool` | `True` for horizontal bar charts |
| `axis_ids` | `list[int]` | Axis IDs this layer references |
| `series` | `list[Series]` | Only the series in this layer |

---

### `Series`

| Property | Type | Description |
|---|---|---|
| `index` | `int` | 0-based series index |
| `order` | `int` | Plot order |
| `name` | `str \| None` | Series name |
| `values` | `list[float]` | Cached numeric values (NaN gaps → `0.0`) |
| `categories` | `list[str]` | Cached category label strings |
| `value_ref` | `str \| None` | Cell-range formula, e.g. `"Sheet1!$B$2:$B$12"` |
| `category_ref` | `str \| None` | Category cell-range formula |
| `is_secondary_axis` | `bool` | `True` when on right/top axis |
| `axis_id` | `int \| None` | Numeric ID of the value axis |

**Methods:**

```python
fill = series.fill(theme)    # → Fill | None
```

---

### `ChartPosition`

| Property | Type | Description |
|---|---|---|
| `sheet` | `str` | Worksheet name |
| `top_left` | `str` | A1-notation, e.g. `"B2"` |
| `bottom_right` | `str` | A1-notation, e.g. `"J17"` |
| `width_emu` | `int \| None` | EMU width (`oneCellAnchor` only) |
| `height_emu` | `int \| None` | EMU height (`oneCellAnchor` only) |

---

## 🎨 Color System

Excel colors come from three sources. `sheetforge-charts` resolves all of them to `"#RRGGBB"` hex strings.

| Source | Example raw | Resolved hex |
|---|---|---|
| Direct hex | `srgb:#4472C4` | `#4472C4` |
| Theme slot | `theme:accent1` | `#4472C4` (from theme) |
| Preset name | `preset:red` | `#FF0000` |

### `Theme`

```python
theme = wb.theme

# All 12 slots at once
print(theme.colors())
# {"accent1": "#4472C4", "accent2": "#ED7D31", "dk1": "#000000", ...}

# Individual slots
print(theme.accent1)            # "#4472C4"
print(theme.accent2)            # "#ED7D31"
print(theme.color("accent3"))   # "#A9D18E"
```

Available slots: `accent1`–`accent6`, `dk1`, `dk2`, `lt1`, `lt2`, `hlink`, `folHlink`

### `Fill`

```python
fill = series.fill(theme)

if fill and fill.fill_type == "solid":
    print(fill.color)           # "#4472C4"   — resolved hex
    print(fill.color_raw)       # "theme:accent1"  — before resolution

elif fill and fill.fill_type == "gradient":
    print(fill.gradient_angle)  # 90.0  (degrees)
    for stop in fill.gradient_stops:
        print(f"{stop.position:.0%}: {stop.color}")
        # 0%:   #4472C4
        # 100%: #2F528F

elif fill is None:
    # Excel uses automatic accent cycle — no explicit fill set
    slot  = f"accent{(series.index % 6) + 1}"
    color = theme.color(slot)
```

### Default series colors

When `series.fill(theme)` returns `None`, Excel assigns colors from the theme accent cycle automatically. Reproduce it with:

```python
DEFAULT_SLOTS = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"]

def series_color(series, theme):
    fill = series.fill(theme)
    if fill and fill.fill_type == "solid":
        return fill.color
    return theme.color(DEFAULT_SLOTS[series.index % 6])
```

---

## 📊 Rendering Charts

`sheetforge-charts` gives you the metadata. You choose the renderer.

### With Plotly (recommended — interactive + PNG export)

```bash
pip install plotly kaleido
```

```python
import sheetforge_charts as sc
import plotly.graph_objects as go

wb    = sc.extract_charts("report.xlsx")
theme = wb.theme

for sheet in wb.sheets:
    for i, chart in enumerate(sheet.charts):
        fig = go.Figure()

        for s in chart.series:
            fill  = s.fill(theme)
            color = fill.color if fill else theme.color(f"accent{(s.index%6)+1}")
            x     = s.categories or list(range(len(s.values)))

            if chart.chart_type in ("Bar", "HorizontalBar"):
                fig.add_trace(go.Bar(x=x, y=s.values, name=s.name, marker_color=color))

            elif chart.chart_type == "Line":
                fig.add_trace(go.Scatter(
                    x=x, y=s.values, name=s.name,
                    mode="lines+markers", line=dict(color=color)
                ))

            elif chart.chart_type == "Pie":
                fig.add_trace(go.Pie(labels=s.categories, values=s.values))

        fig.update_layout(title=chart.title, barmode="group")
        fig.write_html(f"{sheet.name}_chart{i}.html")   # interactive
        fig.write_image(f"{sheet.name}_chart{i}.png")   # static PNG
```

### With Matplotlib

```python
import sheetforge_charts as sc
import matplotlib.pyplot as plt

wb    = sc.extract_charts("report.xlsx")
theme = wb.theme

for sheet in wb.sheets:
    for i, chart in enumerate(sheet.charts):
        fig, ax = plt.subplots(figsize=(10, 5))
        ax.set_title(chart.title or chart.chart_type)

        for s in chart.series:
            fill  = s.fill(theme)
            color = fill.color if fill else theme.color(f"accent{(s.index%6)+1}")
            x     = s.categories or range(len(s.values))

            if chart.chart_type == "Bar":
                ax.bar(x, s.values, label=s.name, color=color)
            elif chart.chart_type == "Line":
                ax.plot(x, s.values, label=s.name, color=color, marker="o")

        ax.legend()
        plt.tight_layout()
        plt.savefig(f"{sheet.name}_chart{i}.png", dpi=150)
        plt.close()
```

### With LibreOffice (pixel-exact, no Excel needed)

```bash
# Install LibreOffice from https://libreoffice.org (free, ~350MB)
```

```python
import subprocess, sheetforge_charts as sc
from PIL import Image  # pip install pillow

wb = sc.extract_charts("report.xlsx")

# Render the entire workbook to PNG sheets
subprocess.run([
    "soffice", "--headless",
    "--convert-to", "png",
    "--outdir", "output/",
    "report.xlsx"
], check=True)

# Crop each chart using ChartAnchor pixel coordinates
COL_PX, ROW_PX = 64, 20  # approximate px per column / row

for sheet in wb.sheets:
    img = Image.open(f"output/report_{sheet.name}.png")
    for i, chart in enumerate(sheet.charts):
        if not chart.anchor: continue
        a = chart.anchor
        crop = (a.col_start*COL_PX, a.row_start*ROW_PX,
                a.col_end*COL_PX,   a.row_end*ROW_PX)
        img.crop(crop).save(f"output/{sheet.name}_chart{i}.png")
```

---

## 🏗️ How It Works

The pipeline is split into two deliberate phases:

```
Phase A  ─ Serial I/O  ──────────────────────────────────────────────────
  ZipArchive (not Send) reads every chart's raw XML bytes into memory.
  Relationship chains are walked: workbook → sheet → drawing → chart.
  Each chart gets a skeleton with path + anchor + pivot refs.

Phase B  ─ Parallel parse  ──────────────────────────────────────────────
  Rayon par_iter hands each (path, bytes) pair to a worker.
  Workers call chart_parser::parse_bytes independently.
  Pure CPU, zero I/O, zero shared state.
  Results collected back in original order.

  ┌─ Worker 1: chart1.xml ──┐
  ├─ Worker 2: chart2.xml ──┤  → Vec<Chart> in original order
  └─ Worker N: chartN.xml ──┘

Phase A5 ─ Pivot metadata (serial) ──────────────────────────────────────
  For pivot charts only: walk pivotTable → pivotCacheDefinition → records.
  Aggregates cache records into pivot_series.
```

On a 50-chart workbook, Phase B completes in ≈ `1/N_CORES` of serial time.

---

## 🔢 Type Reference

### Chart types returned by `chart.chart_type`

| Value | Excel chart |
|---|---|
| `"Bar"` | Clustered / stacked column chart |
| `"HorizontalBar"` | Horizontal bar chart |
| `"Line"` | Line chart |
| `"Pie"` | Pie chart |
| `"Area"` | Area chart |
| `"Scatter"` | XY scatter chart |
| `"Bubble"` | Bubble chart |
| `"Radar"` | Radar / spider chart |
| `"Doughnut"` | Doughnut chart |
| `"Bar3D"` | 3-D column chart |
| `"HorizontalBar3D"` | 3-D horizontal bar chart |
| `"Line3D"` | 3-D line chart |
| `"Area3D"` | 3-D area chart |
| `"Surface3D"` | 3-D surface chart |
| `"Combo"` | Mixed types (see `chart.layers`) |
| `"Unknown"` | Unrecognised chart element |

### Fill types returned by `fill.fill_type`

| Value | Meaning |
|---|---|
| `"solid"` | Single color — use `fill.color` |
| `"gradient"` | Gradient — use `fill.gradient_stops` |
| `"none"` | Explicit transparent (no fill) |

---

## 📦 Building & Distribution

```bash
# Development install (editable, fastest rebuild)
PYO3_USE_ABI3_FORWARD_COMPATIBILITY=1 python -m maturin develop --features python

# Release wheel for your current platform
PYO3_USE_ABI3_FORWARD_COMPATIBILITY=1 python -m maturin build --features python --release

# Publish to PyPI
PYO3_USE_ABI3_FORWARD_COMPATIBILITY=1 python -m maturin publish --features python \
  --username __token__ --password pypi-YOUR_TOKEN_HERE
```

For multi-platform wheels (Windows + Linux + macOS), use the GitHub Actions workflow in `.github/workflows/release.yml`.

---

## 🧪 Testing

```bash
# Full test suite (385 unit + 271 integration tests)
cargo test

# Rust tests only
cargo test --lib

# Integration tests only
cargo test --test integration_test

# With output
cargo test -- --nocapture
```

---

## 📁 Project Structure

```
sheetforge-charts/
├── src/
│   ├── lib.rs                  ← extract_charts() — two-phase pipeline
│   ├── python_bindings.rs      ← PyO3 module (_core)
│   ├── bin/inspect.rs          ← CLI: dump metadata for any .xlsx
│   ├── archive/zip_reader.rs   ← XLSX ZIP reading
│   ├── openxml/
│   │   ├── drawing.rs          ← twoCellAnchor / oneCellAnchor parser
│   │   ├── relationships.rs    ← .rels chain resolver
│   │   └── theme_parser.rs     ← xl/theme/theme1.xml parser
│   ├── model/
│   │   ├── chart.rs            ← Chart, ChartType, ChartLayer, ChartPosition
│   │   ├── series.rs           ← Series, DataValues, StringValues
│   │   ├── color.rs            ← Fill, ColorSpec, Rgb, Gradient
│   │   ├── axis.rs             ← Axis, AxisType, AxisPosition
│   │   ├── theme.rs            ← Theme (slot → Rgb)
│   │   ├── pivot.rs            ← PivotTableMeta, PivotField
│   │   └── workbook.rs         ← WorkbookCharts, SheetCharts
│   └── parser/
│       ├── chart_parser.rs     ← streaming chartN.xml state machine
│       ├── sheet_parser.rs     ← sheet/drawing/chart rel chain walker
│       ├── pivot_table_parser.rs
│       ├── pivot_cache_parser.rs
│       └── pivot_records_parser.rs
├── tests/
│   ├── integration_test.rs     ← 271 integration tests
│   └── fixtures/               ← test .xlsx files
├── python/
│   └── sheetforge_charts/
│       ├── __init__.py         ← Python package wrapper
│       └── __init__.pyi        ← Type stubs for IDE autocomplete
├── Cargo.toml
└── pyproject.toml
```

---

## 📄 License

MIT — see [LICENSE](LICENSE).

---

<div align="center">

Built with 🦀 Rust · PyO3 · Rayon · quick-xml

**[PyPI](https://pypi.org/project/sheetforge-charts/)** · **[GitHub](https://github.com/your-org/sheetforge-charts)** · **[Docs](https://your-org.github.io/sheetforge-charts/)**

</div>
