"""Type stubs for sheetforge_charts — enables IDE auto-complete and mypy."""

from typing import Optional

def extract_charts(path: str) -> WorkbookCharts: ...

class GradientStop:
    position: float
    color: Optional[str]
    color_raw: Optional[str]

class Fill:
    fill_type: str  # "solid" | "gradient" | "none"
    color: Optional[str]
    color_raw: Optional[str]
    gradient_stops: list[GradientStop]
    gradient_angle: Optional[float]

class Theme:
    name: Optional[str]
    accent1: Optional[str]
    accent2: Optional[str]
    accent3: Optional[str]
    accent4: Optional[str]
    accent5: Optional[str]
    accent6: Optional[str]
    dk1: Optional[str]
    lt1: Optional[str]
    dk2: Optional[str]
    lt2: Optional[str]
    def colors(self) -> dict[str, str]: ...
    def color(self, slot: str) -> Optional[str]: ...

class WorkbookCharts:
    source_path: str
    chart_count: int
    sheets: list[SheetCharts]
    theme: Optional[Theme]

class SheetCharts:
    name: str
    index: int
    charts: list[Chart]

class Chart:
    chart_path: str
    chart_type: str
    title: Optional[str]
    style: Optional[int]
    is_pivot_chart: bool
    pivot_table_name: Optional[str]
    layers: list[ChartLayer]
    series: list[Series]
    position: Optional[ChartPosition]
    anchor: Optional[ChartAnchor]
    def chart_fill(self, theme: Optional[Theme] = None) -> Optional[Fill]: ...
    def plot_area_fill(self, theme: Optional[Theme] = None) -> Optional[Fill]: ...

class ChartLayer:
    chart_type: str
    grouping: Optional[str]
    bar_horizontal: bool
    axis_ids: list[int]
    series: list[Series]

class Series:
    index: int
    order: int
    name: Optional[str]
    is_secondary_axis: bool
    axis_id: Optional[int]
    values: list[float]
    categories: list[str]
    value_ref: Optional[str]
    category_ref: Optional[str]
    def fill(self, theme: Optional[Theme] = None) -> Optional[Fill]: ...

class ChartPosition:
    sheet: str
    top_left: str
    bottom_right: str
    width_emu: Optional[int]
    height_emu: Optional[int]

class ChartAnchor:
    col_start: int
    row_start: int
    col_end: int
    row_end: int
    col_span: int
    row_span: int
