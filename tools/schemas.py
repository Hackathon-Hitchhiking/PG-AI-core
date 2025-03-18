from datetime import datetime
from enum import IntEnum
from typing import Literal

from pydantic import BaseModel, Field


class FillType(IntEnum):
    """PowerPoint fill type enumeration"""
    SOLID = 1  # Solid fill
    THEME = 2  # Theme color fill 
    PICTURE = 3  # Picture fill
    PATTERN = 4  # Pattern fill

# Базовые модели данных
class Coordinates(BaseModel):
    """Координаты и размеры элемента"""
    x: float | None = Field(None, description="Горизонтальная позиция")
    y: float | None = Field(None, description="Вертикальная позиция")
    width: float = Field(..., description="Ширина элемента")
    height: float = Field(..., description="Высота элемента")

class FontStyle(BaseModel):
    """Стиль шрифта текстового элемента"""
    name: str = Field(default="Calibri", description="Название шрифта")
    size: float = Field(default=12.0, description="Размер шрифта в пунктах")
    color: str = Field(default="#000000", description="Цвет в HEX-формате")
    bold: bool = Field(default=False, description="Жирное начертание")
    italic: bool = Field(default=False, description="Курсивное начертание")
    underline: bool = Field(default=False, description="Подчеркнутый текст")
    alignment: Literal["left", "center", "right", "justify"] = Field("left", description="Выравнивание текста")

class Background(BaseModel):
    """Фон слайда с поддержкой отсутствующего фона"""
    type: Literal["color", "image", "pattern", "none"] = Field("none", description="Тип фона")
    value: str | None = Field(None, description="HEX-код, путь к файлу или паттерн")
    transparency: float = Field(1.0, ge=0.0, le=1.0, description="Прозрачность 0-1")


class SlideLayout(BaseModel):
    name: str
    display_name: str
    idx: int


# Модели элементов слайда
class TextContent(BaseModel):
    text: str = Field(..., description="Текстовое содержимое")
    bullet: bool = Field(default=False, description="Маркированный список")
    bullet_level: int = Field(0, ge=0, description="Уровень вложенности списка")

class TextShape(BaseModel):
    type: Literal["text"] = "text"
    id: str = Field(..., description="Уникальный идентификатор элемента")
    content: TextContent
    coordinates: Coordinates
    style: FontStyle

class ImageShape(BaseModel):
    type: Literal["image"] = "image"
    id: str
    path: str = Field(..., description="Путь к файлу или base64-строка")
    coordinates: Coordinates
    crop: Coordinates | None = Field(None, description="Область обрезки")
    brightness: float = Field(1.0, ge=0.0, le=2.0)
    contrast: float = Field(1.0, ge=0.0, le=2.0)

class ChartSeries(BaseModel):
    name: str
    values: list[float]

class ChartData(BaseModel):
    categories: list[str]
    series: list[ChartSeries]

class ChartShape(BaseModel):
    type: Literal["chart"] = "chart"
    id: str
    chart_type: Literal["bar", "line", "pie", "area"]
    data: ChartData
    coordinates: Coordinates
    title: str | None = None

class TableCell(BaseModel):
    row: int
    col: int
    content: str
    span_rows: int = 1
    span_cols: int = 1

class TableShape(BaseModel):
    type: Literal["table"] = "table"
    id: str
    rows: int
    cols: int
    cells: list[TableCell]
    coordinates: Coordinates

class GeometricShape(BaseModel):
    type: Literal["shape"] = "shape"
    id: str
    shape_type: Literal["rectangle", "ellipse", "line", "arrow"]
    coordinates: Coordinates
    fill: str = "#FFFFFF"
    outline: str = "#000000"
    outline_width: float = 1.0

Shape = TextShape | ImageShape | ChartShape | TableShape | GeometricShape

# Модели слайдов и презентации
class SlideDesign(BaseModel):
    theme_colors: list[str]
    theme_fonts: dict[str, str]
    default_margin: float = 0.5

class Slide(BaseModel):
    number: int = Field(..., ge=1, description="Порядковый номер слайда")
    layout: SlideLayout
    background: Background
    shapes: list[Shape]
    design: SlideDesign

class PresentationMeta(BaseModel):
    title: str | None = None
    author: str | None = None
    created: datetime | None = None
    modified: datetime | None = None
    template: str | None = Field(None, description="Используемый шаблон")

class PresentationInfo(BaseModel):
    meta: PresentationMeta
    dimensions: Coordinates = Field(..., description="Ширина и высота слайдов")
    slides_count: int
    slides: list[Slide]
    master_layouts: list[SlideLayout]
