import logging
from functools import wraps
from typing import Any, Dict, List, Optional, Tuple, Union
from enum import Enum
from pptx import Presentation
from pptx.chart.axis import CategoryAxis, ValueAxis
from pptx.chart.chart import Chart
from pptx.chart.data import ChartData as PptxChartData
from pptx.chart.series import XySeries
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_MARKER_STYLE
from pptx.enum.dml import MSO_FILL, MSO_LINE
from pptx.shapes.base import BaseShape
from pptx.slide import Slide
from pptx.util import Inches, Pt
from pydantic import BaseModel

# region Data Models
class ChartType(str, Enum):
    BAR = "bar"
    COLUMN = "column"
    LINE = "line"
    PIE = "pie"
    DOUGHNUT = "doughnut"
    AREA = "area"
    SCATTER = "scatter"
    BUBBLE = "bubble"
    RADAR = "radar"
    SURFACE = "surface"
    COMBO = "combo"

class AxisMetadata(BaseModel):
    title: Optional[str]
    labels: List[str]
    min: Optional[float]
    max: Optional[float]
    number_format: Optional[str]
    font: Optional[Dict[str, Any]]
    is_linked: Optional[bool]
    units: Optional[str]
    axis_type: Optional[str]
    visible: bool = True

class LegendEntry(BaseModel):
    name: str
    color: Optional[str]

class LegendMetadata(BaseModel):
    position: Optional[str]
    font: Optional[Dict[str, Any]]
    entries: List[LegendEntry]
    visible: bool = True

class ChartSeries(BaseModel):
    name: str
    values: List[Union[float, str]]
    series_type: ChartType 
    color: Optional[Dict[str, Any]]
    marker: Optional[Dict[str, Any]]
    trendline: Optional[Dict[str, Any]]
    error_bars: Optional[Dict[str, Any]]
    data_labels: Optional[Dict[str, Any]]

class ChartData(BaseModel):
    categories: List[Union[str, Tuple[str, ...]]]
    series: List[ChartSeries]

class ChartStyle(BaseModel):
    fill: Optional[Dict[str, Any]]
    line: Optional[Dict[str, Any]]
    effects: Optional[Dict[str, Any]]

class ChartMetadata(BaseModel):
    title: Optional[Dict[str, Any]]
    axes: Dict[str, AxisMetadata]
    legend: Optional[LegendMetadata]
    plot_area: Dict[str, Any]
    style: ChartStyle
    size: Dict[str, Any]
    slide_id: int
    shape_id: int

class ChartShape(BaseModel):
    chart_type: ChartType
    data: ChartData
    metadata: ChartMetadata
# endregion

class ChartManager:
    def __init__(self, presentation_path: str):
        self.prs = Presentation(presentation_path)
        self.logger = logging.getLogger(__name__)
        
        self.type_map = {
            XL_CHART_TYPE.BAR_CLUSTERED: ChartType.BAR,
            XL_CHART_TYPE.BAR_STACKED: ChartType.BAR,
            XL_CHART_TYPE.COLUMN_CLUSTERED: ChartType.COLUMN,
            XL_CHART_TYPE.COLUMN_STACKED: ChartType.COLUMN,
            XL_CHART_TYPE.LINE: ChartType.LINE,
            XL_CHART_TYPE.LINE_MARKERS: ChartType.LINE,
            XL_CHART_TYPE.PIE: ChartType.PIE,
            XL_CHART_TYPE.DOUGHNUT: ChartType.DOUGHNUT,
            XL_CHART_TYPE.AREA: ChartType.AREA,
            XL_CHART_TYPE.XY_SCATTER: ChartType.SCATTER,
            XL_CHART_TYPE.BUBBLE: ChartType.BUBBLE,
            XL_CHART_TYPE.RADAR: ChartType.RADAR,
            XL_CHART_TYPE.SURFACE: ChartType.SURFACE,
        }
        self.reverse_type_map = {v: k for k, v in self.type_map.items()}

    # region Public API
    def get_chart_by_ids(self, slide_id: int, shape_id: int) -> Optional[ChartShape]:
        """Get chart by slide ID and shape ID"""
        slide = self._get_slide_by_id(slide_id)
        if not slide:
            return None
            
        shape = self._get_shape_by_id(slide, shape_id)
        if shape and shape.has_chart:
            return self.parse_chart(shape, slide_id, shape_id)
        return None

    def parse_all_charts(self) -> List[ChartShape]:
        """Parse all charts in presentation"""
        charts = []
        for slide_idx, slide in enumerate(self.prs.slides):
            for shape in slide.shapes:
                if shape.has_chart:
                    charts.append(
                        self.parse_chart(shape, slide_idx, shape.shape_id)
                    )
        return charts

    def create_chart(self, slide_id: int, chart_data: Dict[str, Any]) -> Optional[ChartShape]:
        """Create new chart on specified slide"""
        slide = self._get_slide_by_id(slide_id)
        if not slide:
            return None

        try:
            chart_type = self._get_chart_type(chart_data['chart_type'])
            chart = slide.shapes.add_chart(
                chart_type,
                Inches(chart_data['position']['left']),
                Inches(chart_data['position']['top']),
                Inches(chart_data['position']['width']),
                Inches(chart_data['position']['height']),
                self._prepare_chart_data(chart_data['data'])
            ).chart
            
            if 'style' in chart_data:
                self._apply_chart_style(chart, chart_data['style'])
            
            return self.parse_chart(chart.parent, slide_id, chart.parent.shape_id)
        except Exception as e:
            self.logger.error(f"Chart creation failed: {str(e)}")
            return None
    # endregion

    # region Parsing Logic
    def parse_chart(self, shape: BaseShape, slide_id: int, shape_id: int) -> ChartShape:
        """Main parsing method"""
        chart = shape.chart
        return ChartShape(
            chart_type=self._parse_chart_type(chart),
            data=self._parse_chart_data(chart),
            metadata=ChartMetadata(
                title=self._parse_title(chart),
                axes=self._parse_axes(chart),
                legend=self._parse_legend(chart),
                plot_area=self._parse_plot_area(chart),
                style=self._parse_chart_style(chart),
                size=self._parse_size(shape),
                slide_id=slide_id,
                shape_id=shape_id
            )
        )

    def _parse_chart_type(self, chart: Chart) -> ChartType:
        chart_type = self.type_map.get(chart.chart_type, ChartType.COLUMN)
        if chart_type == ChartType.COMBO:
            return self._handle_combo_chart(chart)
        return chart_type

    def _parse_chart_data(self, chart: Chart) -> ChartData:
        return ChartData(
            categories=self._parse_categories(chart),
            series=self._parse_series(chart)
        )

    def _parse_categories(self, chart: Chart) -> List[Union[str, Tuple]]:
        try:
            if chart.chart_type in [XL_CHART_TYPE.PIE, XL_CHART_TYPE.DOUGHNUT]:
                return [str(point.data_label.text) for point in chart.series[0].points]
                
            if chart.chart_type == XL_CHART_TYPE.XY_SCATTER:
                return list(zip(
                    [str(x) for x in chart.series[0].x_values],
                    [str(y) for y in chart.series[0].y_values]
                ))
                
            return [
                str(cat) 
                for cat in chart.plots[0].categories
                if chart.plots[0].categories
            ]
        except Exception as e:
            self.logger.error(f"Categories parsing error: {str(e)}")
            return []

    def _parse_series(self, chart: Chart) -> List[ChartSeries]:
        return [self._parse_single_series(series, chart) for series in chart.series]

    def _parse_single_series(self, series: XySeries, chart: Chart) -> ChartSeries:
        return ChartSeries(
            name=series.name or "Unnamed Series",
            values=[float(v) if v is not None else 0.0 for v in series.values],
            series_type=self.type_map.get(chart.chart_type, ChartType.COLUMN),
            color=self._parse_series_style(series),
            marker=self._parse_marker(series),
            trendline=self._parse_trendline(series),
            error_bars=self._parse_error_bars(series),
            data_labels=self._parse_data_labels(series)
        )

    def _parse_series_style(self, series: XySeries) -> Dict[str, Any]:
        return {
            'fill': self._parse_fill(series.format.fill),
            'line': self._parse_line(series.format.line)
        }

    def _parse_marker(self, series: XySeries) -> Optional[Dict[str, Any]]:
        if not hasattr(series, 'marker') or series.marker.style == XL_MARKER_STYLE.NONE:
            return None
            
        return {
            'style': series.marker.style.name,
            'size': series.marker.size,
            'color': self._parse_color(series.marker.format.fill.fore_color),
            'line': self._parse_line(series.marker.format.line)
        }
    # endregion

    # region Helper Methods
    def _get_slide_by_id(self, slide_id: int) -> Optional[Slide]:
        try:
            return self.prs.slides[slide_id]
        except IndexError:
            self.logger.error(f"Slide ID {slide_id} not found")
            return None

    def _get_shape_by_id(self, slide: Slide, shape_id: int) -> Optional[BaseShape]:
        for shape in slide.shapes:
            if shape.shape_id == shape_id:
                return shape
        self.logger.error(f"Shape ID {shape_id} not found in slide")
        return None

    def _parse_color(self, color_obj) -> Optional[str]:
        if isinstance(color_obj, RGBColor):
            return f"#{color_obj.rgb:06X}"
        return None

    def _parse_fill(self, fill) -> Dict[str, Any]:
        if fill.type == MSO_FILL.SOLID:
            return {
                'type': 'solid',
                'color': self._parse_color(fill.fore_color)
            }
        elif fill.type == MSO_FILL.GRADIENT:
            return {
                'type': 'gradient',
                'colors': [self._parse_color(stop.color) for stop in fill.gradient_stops]
            }
        return {'type': 'none'}

    def _parse_line(self, line) -> Dict[str, Any]:
        if line.visible:
            return {
                'color': self._parse_color(line.color),
                'width': line.width.pt,
                'style': line.dash_style.name if line.dash_style else 'solid'
            }
        return {'visible': False}

    def _convert_to_rgb(self, color: Union[str, tuple]) -> RGBColor:
        """Convert color string or tuple to RGBColor"""
        if isinstance(color, str):
            return RGBColor.from_string(color.lstrip('#'))
        elif isinstance(color, tuple):
            return RGBColor(*color)
        raise ValueError(f"Invalid color: {color}")
    # endregion
    # region Style Application
    def _apply_chart_style(self, chart: Chart, style: Dict):
        """Apply complete styling to chart"""
        try:
            if 'title' in style:
                self._apply_title_style(chart, style['title'])
            
            if 'legend' in style:
                self._apply_legend_style(chart, style['legend'])
            
            if 'series_styles' in style:
                for idx, series_style in enumerate(style['series_styles']):
                    if idx < len(chart.series):
                        self._apply_series_style(chart.series[idx], series_style)
            
            if 'axes' in style:
                self._apply_axes_style(chart, style['axes'])
            
            if 'plot_area' in style:
                self._apply_plot_area_style(chart, style['plot_area'])
            
            if 'color_palette' in style:
                self._apply_color_palette(chart, style['color_palette'])
            
            if 'rotation_3d' in style:
                self._apply_3d_rotation(chart, style['rotation_3d'])
            
            if 'data_labels' in style:
                self._apply_data_labels(chart, style['data_labels'])
            
        except Exception as e:
            self.logger.error(f"Error applying chart style: {str(e)}")

    def _apply_title_style(self, chart: Chart, style: Dict):
        """Apply title formatting"""
        if not chart.has_title:
            if 'text' in style:
                chart.has_title = True
            else:
                return

        title = chart.chart_title
        try:
            if 'text' in style:
                title.text_frame.text = style['text']
            
            if 'font' in style:
                for paragraph in title.text_frame.paragraphs:
                    for run in paragraph.runs:
                        self._apply_font_style(run.font, style['font'])
            
            if 'fill_color' in style:
                title.format.fill.solid()
                title.format.fill.fore_color.rgb = self._convert_to_rgb(
                    style['fill_color']
                )
            
            if 'border' in style:
                self._apply_line_style(title.format.line, style['border'])
        
        except Exception as e:
            self.logger.error(f"Title styling failed: {str(e)}")

    def _apply_legend_style(self, chart: Chart, style: Dict):
        """Configure legend appearance"""
        try:
            chart.has_legend = style.get('visible', True)
            
            if 'position' in style:
                chart.legend.position = getattr(
                    XL_LEGEND_POSITION, 
                    style['position'].upper()
                )
            
            if 'font' in style:
                for entry in chart.legend.legend_entries:
                    self._apply_font_style(entry.font, style['font'])
            
            if 'border' in style:
                self._apply_line_style(chart.legend.format.line, style['border'])
            
            if 'background_color' in style:
                chart.legend.format.fill.solid()
                chart.legend.format.fill.fore_color.rgb = self._convert_to_rgb(
                    style['background_color']
                )
        
        except Exception as e:
            self.logger.error(f"Legend styling failed: {str(e)}")

    def _apply_series_style(self, series: XySeries, style: Dict):
        """Style individual data series"""
        try:
            if 'fill' in style:
                series.format.fill.solid()
                series.format.fill.fore_color.rgb = self._convert_to_rgb(
                    style['fill']
                )
            
            if 'line' in style:
                self._apply_line_style(series.format.line, style['line'])
            
            if 'marker' in style:
                self._apply_marker_style(series, style['marker'])
            
            if 'data_labels' in style:
                self._apply_series_data_labels(series, style['data_labels'])
            
            if 'trendline' in style:
                self._apply_trendline(series, style['trendline'])
        
        except Exception as e:
            self.logger.error(f"Series styling failed: {str(e)}")

    def _apply_axes_style(self, chart: Chart, axes_config: Dict):
        """Configure chart axes"""
        try:
            # Primary axes
            if 'category_axis' in axes_config:
                self._apply_axis_style(
                    chart.category_axis,
                    axes_config['category_axis']
                )
            
            if 'value_axis' in axes_config:
                self._apply_axis_style(
                    chart.value_axis,
                    axes_config['value_axis']
                )
            
            # Secondary axis for combo charts
            if 'secondary_axis' in axes_config and chart.chart_type == XL_CHART_TYPE.COMBO:
                self._apply_axis_style(
                    chart.secondary_value_axis,
                    axes_config['secondary_axis']
                )
        
        except Exception as e:
            self.logger.error(f"Axes styling failed: {str(e)}")

    def _apply_axis_style(self, axis, style: Dict):
        """Style individual axis"""
        try:
            if 'title' in style:
                axis.axis_title.text_frame.text = style['title']
            
            if 'font' in style:
                self._apply_font_style(axis.tick_labels.font, style['font'])
            
            if 'number_format' in style:
                axis.tick_labels.number_format = style['number_format']
            
            if 'min' in style:
                axis.minimum_scale = style['min']
            if 'max' in style:
                axis.maximum_scale = style['max']
            
            if 'line' in style:
                self._apply_line_style(axis.format.line, style['line'])
            
            if 'major_gridlines' in style:
                self._apply_gridlines_style(
                    axis.major_gridlines, 
                    style['major_gridlines']
                )
        
        except Exception as e:
            self.logger.error(f"Axis styling failed: {str(e)}")

    def _apply_font_style(self, font, style: Dict):
        """Generic font styling method"""
        try:
            if 'name' in style:
                font.name = style['name']
            if 'size' in style:
                font.size = Pt(style['size'])
            if 'bold' in style:
                font.bold = style['bold']
            if 'italic' in style:
                font.italic = style['italic']
            if 'color' in style:
                font.color.rgb = self._convert_to_rgb(style['color'])
            if 'underline' in style:
                font.underline = style['underline']
        except Exception as e:
            self.logger.error(f"Font styling failed: {str(e)}")

    # region Combo Chart Handling
    def _handle_combo_chart(self, chart: Chart) -> ChartType:
        """Determine combo chart type based on series"""
        series_types = set()
        for series in chart.series:
            series_types.add(self.type_map.get(series.chart_type, ChartType.COLUMN))
        
        if len(series_types) > 1:
            return ChartType.COMBO
        return series_types.pop() if series_types else ChartType.COLUMN

    def _apply_combo_chart_settings(self, chart: Chart, settings: Dict):
        """Configure combo chart secondary axis"""
        try:
            for idx, series_settings in enumerate(settings.get('series_settings', [])):
                if idx >= len(chart.series):
                    continue
                
                series = chart.series[idx]
                if 'axis' in series_settings:
                    axis_type = series_settings['axis']
                    if axis_type == 'secondary':
                        series.axis_group = 2
        except Exception as e:
            self.logger.error(f"Combo chart configuration failed: {str(e)}")
    # endregion

    # region Additional Parsers
    def _parse_legend(self, chart: Chart) -> Optional[LegendMetadata]:
        if not chart.has_legend:
            return None
        
        legend = chart.legend
        entries = []
        for entry in legend.legend_entries:
            entries.append(
                LegendEntry(
                    name=entry.text,
                    color=self._parse_color(entry.font.color)
                ))
        
        return LegendMetadata(
            position=legend.position.name,
            font=self._parse_font(legend.font),
            entries=entries,
            visible=True
        )

    def _parse_plot_area(self, chart: Chart) -> Dict[str, Any]:
        return {
            'fill': self._parse_fill(chart.plots[0].format.fill),
            'line': self._parse_line(chart.plots[0].format.line)
        }

    def _parse_chart_style(self, chart: Chart) -> ChartStyle:
        return ChartStyle(
            fill=self._parse_fill(chart.chart_area.format.fill),
            line=self._parse_line(chart.chart_area.format.line),
            effects={}  # Placeholder for future implementation
        )

    def _parse_size(self, shape: BaseShape) -> Dict[str, Any]:
        return {
            'width': shape.width.inches,
            'height': shape.height.inches,
            'left': shape.left.inches,
            'top': shape.top.inches
        }

    def _parse_font(self, font) -> Dict[str, Any]:
        return {
            'name': font.name,
            'size': font.size.pt,
            'bold': font.bold,
            'italic': font.italic,
            'color': self._parse_color(font.color),
            'underline': font.underline
        }
    # endregion

    # region Error Handling
    def _get_chart_type(self, chart_type_str: str) -> XL_CHART_TYPE:
        try:
            return self.reverse_type_map[chart_type_str]
        except KeyError:
            self.logger.warning(f"Unknown chart type {chart_type_str}, using default")
            return XL_CHART_TYPE.COLUMN_CLUSTERED
    # endregion