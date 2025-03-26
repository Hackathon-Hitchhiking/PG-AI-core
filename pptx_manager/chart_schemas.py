from enum import Enum
from typing import Any, Dict, List, Optional, Union

from pydantic import BaseModel


class ChartType(str, Enum):
    BAR = "bar"
    COLUMN = "column"
    LINE = "line"
    PIE = "pie"
    AREA = "area"
    SCATTER = "scatter"
    DOUGHNUT = "doughnut"
    COMBO = "combo"
    UNKNOWN = "unknown"

class ChartSeries(BaseModel):
    name: str
    values: list[Union[float, str, None]]
    series_type: Optional[ChartType] = None
    color: Optional[str] = None
    marker: Optional[dict] = None
    style: Optional[Dict[str, Any]] = None  # Включает цвет, маркеры, линии
    data_labels: Optional[Dict[str, Any]] = None

class AxisMetadata(BaseModel):
    title: Optional[str] = None
    labels: List[str] = []
    min: Optional[float] = None
    max: Optional[float] = None
    number_format: Optional[str] = None
    font: Optional[Dict[str, Any]] = None
    offset: Optional[int] = None
    is_linked: Optional[bool] = None
    units: Optional[str] = None
    axis_type: Optional[str] = None 

class LegendMetadata(BaseModel):
    position: Optional[str] = None
    entries: List[str] = []

class ChartData(BaseModel):
    categories: List[Union[str, tuple]] = []
    series: List[ChartSeries] = []

class ChartMetadata(BaseModel):
    title: Optional[str] = None
    subtitle: Optional[str] = None
    axes: Dict[str, AxisMetadata] = {}
    legend: Optional[LegendMetadata] = None
    formats: Dict[str, Any] = {}
    gridlines: Optional[Dict[str, Any]] = None
    number_precision: Optional[int] = None
    display_units: Optional[Dict[str, bool]] = None

class ChartShape(BaseModel):
    shape_id: int
    coordinates: Dict[str, float]
    chart_type: ChartType
    data: ChartData
    metadata: ChartMetadata
    plot_area: Optional[Dict[str, Any]] = None
    wall: Optional[Dict[str, Any]] = None