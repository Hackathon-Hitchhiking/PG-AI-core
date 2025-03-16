from typing import List, Optional, Literal
from pydantic import BaseModel

class Coordinates(BaseModel):
    x: float
    y: float
    width: float
    height: float

class FontStyle(BaseModel):
    name: str
    size: float
    color: str  # HEX формат цвета
    bold: bool
    italic: bool

class TextShape(BaseModel):
    type: Literal["text"]
    text: str
    coordinates: Coordinates
    font_style: Optional[FontStyle]

class ImageShape(BaseModel):
    type: Literal["image"]
    image_path: str
    coordinates: Coordinates

class ChartSeries(BaseModel):
    name: str
    values: List[float]

class ChartShape(BaseModel):
    type: Literal["chart"]
    chart_type: str
    coordinates: Coordinates
    categories: List[str]
    series: List[ChartSeries]

Shape = TextShape | ImageShape | ChartShape

class Slide(BaseModel):
    slide_number: int
    shapes: List[Shape]

class PresentationInfo(BaseModel):
    width: float
    height: float
    slides: List[Slide]
