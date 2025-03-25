from typing import Any

from pydantic import BaseModel


class TextFrameShape(BaseModel):
    shape_id: int

    text: str

    font_name: str | None = None
    font_size: float | None = None
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    strike: bool | None = None
    color: tuple | None = None

    # the manager of the shape block. import pptx.text.text.TextFrame
    text_manager: Any
    # the manager of the shape.paragraph.runs. import pptx.text.text.Font
    font_manager: Any


class UpdateTextFrameOpts(BaseModel):
    text: str | None = None
    new_color: tuple[int] | None = None


class RunElement(BaseModel):
    text: str
    font_name: str | None = None
    font_size: float | None = None
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    strike: bool | None = None
    color: tuple | None = None


class ParagraphRunElement(BaseModel):
    alignment: str | None = None
    fonts: list[RunElement]


class ShapeRunElement(BaseModel):
    shape_id: int | None = None
    shape_name: str | None = None
    paragraphs: list[ParagraphRunElement]
