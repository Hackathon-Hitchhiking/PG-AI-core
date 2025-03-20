from datetime import datetime
from enum import IntEnum
from typing import Literal, Union

from pydantic import BaseModel, Field


#
# ENUMS / CONSTANTS
#
class FillType(IntEnum):
    """PowerPoint fill type enumeration"""

    SOLID = 1  # Solid fill
    THEME = 2  # Theme color fill
    PICTURE = 3  # Picture fill
    PATTERN = 4  # Pattern fill
    # Removed NO_FILL because older python-pptx does not define an equivalent,
    # and we often see 'None' if there's no fill at all.


#
# COMMON MODELS
#
class Coordinates(BaseModel):
    """Coordinates and size of an element in inches."""

    x: float | None = Field(None, description='Horizontal position in inches')
    y: float | None = Field(None, description='Vertical position in inches')
    width: float = Field(..., description='Width in inches')
    height: float = Field(..., description='Height in inches')


class FontStyle(BaseModel):
    """Text styling."""

    name: str = Field(default='Calibri', description='Font name')
    size: float = Field(default=12.0, description='Font size in points')
    color: str = Field(default='#000000', description='Text color in HEX format')
    bold: bool = Field(default=False, description='Bold text')
    italic: bool = Field(default=False, description='Italic text')
    underline: bool = Field(default=False, description='Underlined text')
    alignment: Literal['left', 'center', 'right', 'justify'] = Field('left', description='Text alignment')


class Background(BaseModel):
    """Slide background, can be color, image, pattern, or none."""

    type: Literal['color', 'image', 'pattern', 'none'] = Field('none', description='Background type')
    value: str | None = Field(None, description='HEX color code, image path, or pattern name')
    transparency: float = Field(1.0, ge=0.0, le=1.0, description='Transparency (0=opaque, 1=fully transparent)')


class SlideLayout(BaseModel):
    """Details about a single slide layout."""

    name: str
    display_name: str
    idx: int


#
# SHAPES / CONTENT
#
class TextContent(BaseModel):
    text: str = Field(..., description='Text content')
    bullet: bool = Field(default=False, description='Bulleted list?')
    bullet_level: int = Field(0, ge=0, description='Bulleted list indent level')


class TextShape(BaseModel):
    """Represents a text box shape."""

    type: Literal['text'] = 'text'
    id: str = Field(..., description='Unique shape identifier')
    content: TextContent
    coordinates: Coordinates
    style: FontStyle


class ImageShape(BaseModel):
    """Represents an image shape."""

    type: Literal['image'] = 'image'
    id: str
    path: str = Field(..., description='Path to the image file (or base64 string)')
    coordinates: Coordinates
    crop: Coordinates | None = Field(None, description='Crop area coordinates if any')
    brightness: float = Field(1.0, ge=0.0, le=2.0, description='Image brightness multiplier')
    contrast: float = Field(1.0, ge=0.0, le=2.0, description='Image contrast multiplier')


class ChartSeries(BaseModel):
    name: str
    values: list[float]


class ChartData(BaseModel):
    """Data for charts."""

    categories: list[str]
    series: list[ChartSeries]


class ChartShape(BaseModel):
    """Represents a chart shape."""

    type: Literal['chart'] = 'chart'
    id: str
    chart_type: Literal['bar', 'line', 'pie', 'area']  # You can expand this if needed
    data: ChartData
    coordinates: Coordinates
    title: str | None = None


class TableCell(BaseModel):
    """Represents a single table cell, including spans."""

    row: int
    col: int
    content: str
    span_rows: int = 1
    span_cols: int = 1


class TableShape(BaseModel):
    """Represents a table shape."""

    type: Literal['table'] = 'table'
    id: str
    rows: int
    cols: int
    cells: list[TableCell]
    coordinates: Coordinates


class GeometricShape(BaseModel):
    """
    Represents an auto-shape (rectangle, ellipse, arrow, etc.).
    Instead of restricting shape_type to a few literals, we allow any string
    so 'rounded_rectangle' won't fail.
    Fill/outline can be None if the shape has no fill/outline.
    """

    type: Literal['shape'] = 'shape'
    id: str
    shape_type: str = Field('rectangle', description="Auto shape type (e.g. 'rectangle', 'rounded_rectangle', ...)")
    coordinates: Coordinates
    fill: str | None = Field(None, description='Fill color in HEX or None if no fill')
    outline: str | None = Field(None, description='Outline color in HEX or None if no outline')
    outline_width: float = Field(0.0, ge=0.0, description='Outline width in inches')


# Union of all recognized shapes
Shape = Union[TextShape, ImageShape, ChartShape, TableShape, GeometricShape]


#
# MODELS FOR SLIDES / PRESENTATION
#
class SlideDesign(BaseModel):
    """Theme info extracted from a slide."""

    theme_colors: list[str]
    theme_fonts: dict[str, str]
    default_margin: float = 0.5


class Slide(BaseModel):
    """Represents a single slide, including layout, background, and shape data."""

    number: int = Field(..., ge=1, description='1-based slide number')
    layout: SlideLayout
    background: Background
    shapes: list[Shape]  # We skip shapes that can't be parsed, so they won't be None
    design: SlideDesign


class PresentationMeta(BaseModel):
    """Presentation metadata from core properties."""

    title: str | None = None
    author: str | None = None
    created: datetime | None = None
    modified: datetime | None = None
    template: str | None = Field(None, description='Template name if any')


class PresentationInfo(BaseModel):
    """High-level presentation info, including all slides."""

    meta: PresentationMeta
    dimensions: Coordinates
    slides_count: int
    slides: list[Slide]
    master_layouts: list[SlideLayout]
