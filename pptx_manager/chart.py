from typing import Optional, List, Dict
from collections import defaultdict
from pptx import Presentation
from pptx.chart.chart import Chart
from pptx.dml.color import ColorFormat
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.dml import MSO_COLOR_TYPE, MSO_THEME_COLOR
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml import parse_xml
from pptx.slide import Slide
from pydantic import BaseModel
import colorsys

class ChartSeriesPoint(BaseModel):
    value: Optional[float]
    color: Optional[tuple[int, int, int]]

class ChartSeries(BaseModel):
    name: str
    values: List[Optional[float]]
    color: Optional[tuple[int, int, int]]
    points: List[ChartSeriesPoint]

class ChartAxis(BaseModel):
    title: Optional[str]
    visible: bool
    font_color: tuple[int, int, int]
    has_major_grid: bool
    has_minor_grid: bool

class ChartLegend(BaseModel):
    position: Optional[str]
    font_color: tuple[int, int, int]

class ChartData(BaseModel):
    shape_id: int
    chart_type: str
    title: Optional[str]
    legend: Optional[ChartLegend]
    categories: List[str]
    series: List[ChartSeries]
    category_axis: ChartAxis
    value_axis: ChartAxis

class ChartManager:
    def __init__(self):
        self.charts = defaultdict(list[ChartData])

    def _get_theme_color(self, theme_color: MSO_THEME_COLOR, slide: Slide) -> Optional[str]:
        try:
            master = slide.slide_layout.slide_master
            theme = master.part.related_parts[RT.THEME]._element
            return theme.xpath(f"//a:clrScheme/a:{theme_color.xml_value}/a:srgbClr/@val")[0]
        except (IndexError, KeyError, AttributeError):
            return None

    def _resolve_color(self, color: ColorFormat, slide: Slide) -> Optional[tuple[int, int, int]]:
        if color.type == MSO_COLOR_TYPE.RGB:
            return (color.rgb.r, color.rgb.g, color.rgb.b)
        
        if color.type == MSO_COLOR_TYPE.SCHEME:
            hex_color = self._get_theme_color(color.theme_color, slide)
            if hex_color:
                return self._apply_brightness(hex_color, color.brightness)
        
        return None

    def _apply_brightness(self, hex_color: str, brightness: float) -> tuple[int, int, int]:
        r, g, b = (int(hex_color[i:i+2], 16) for i in (0, 2, 4))
        h, l, s = colorsys.rgb_to_hls(r/255, g/255, b/255)
        new_l = max(0.0, min(1.0, l + brightness))
        r, g, b = colorsys.hls_to_rgb(h, new_l, s)
        return (int(r*255), int(g*255), int(b*255))

    def _parse_axis(self, axis, slide: Slide) -> ChartAxis:
        return ChartAxis(
            title=axis.axis_title.text_frame.text if axis.has_title else None,
            visible=axis.visible,
            font_color=self._resolve_color(axis.tick_labels.font.color, slide) or (0,0,0),
            has_major_grid=axis.has_major_gridlines,
            has_minor_grid=axis.has_minor_gridlines
        )

    def _get_chart_type_name(self, chart: Chart) -> str:
        try:
            return chart.chart_type.name
        except AttributeError:
            return "UNKNOWN"

    def _get_categories(self, chart: Chart) -> List[str]:
        try:
            return [
                str(cat.label)
                for cat in chart.plots[0].categories
                if cat.label is not None
            ]
        except Exception as e:
            print(f"Error getting categories: {e}")
            return []

    def _parse_series(self, chart: Chart, slide: Slide) -> List[ChartSeries]:
        series_list = []
        try:
            for series in chart.series:
                series_color = self._resolve_color(series.format.fill.fore_color, slide)
                
                points = []
                for point in series.points:
                    point_color = self._resolve_color(point.format.fill.fore_color, slide)
                    points.append(ChartSeriesPoint(
                        value=point.data_value,
                        color=point_color
                    ))
                
                series_list.append(ChartSeries(
                    name=series.name,
                    values=series.values,
                    color=series_color,
                    points=points
                ))
        except Exception as e:
            print(f"Error parsing series: {e}")
        return series_list

    def parse_chart(self, chart: Chart, slide: Slide, shape_id: int) -> Optional[ChartData]:
        try:
            return ChartData(
                shape_id=shape_id,
                chart_type=self._get_chart_type_name(chart),
                title=chart.chart_title.text_frame.text if chart.has_title else None,
                legend=self._parse_legend(chart.legend, slide) if chart.has_legend else None,
                categories=self._get_categories(chart),
                series=self._parse_series(chart, slide),
                category_axis=self._parse_axis(chart.category_axis, slide),
                value_axis=self._parse_axis(chart.value_axis, slide),
            )
        except Exception as e:
            print(f"Error parsing chart: {e}")
            return None

    def _parse_legend(self, legend, slide: Slide) -> ChartLegend:
        return ChartLegend(
            position=legend.position.name if legend.position else None,
            font_color=self._resolve_color(legend.font.color, slide) or (0,0,0)
        )

def test(file_path: str):
    pres = Presentation(file_path)
    manager = ChartManager()
    
    for slide in pres.slides:
        chart_count = 1
        for shape in slide.shapes:
            if shape.has_chart:
                chart = shape.chart
                chart_data = manager.parse_chart(chart, slide, chart_count)
                if chart_data:
                    manager.charts[slide.slide_id].append(chart_data)
                    chart_count += 1
    
    # Пример вывода
    for slide_id, charts in manager.charts.items():
        print(f"\nSlide {slide_id}:")
        for chart in charts:
            print(f"\nChart #{chart.shape_id}")
            print(f"Type: {chart.chart_type}")
            print(f"Title: {chart.title or 'No title'}")
            print(f"Categories: {chart.categories}")
            print(f"Series: {[s.name for s in chart.series]}")

if __name__ == "__main__":
    test("test_sources/test_dit.pptx")