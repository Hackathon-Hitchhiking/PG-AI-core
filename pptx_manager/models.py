from typing import Any

from pydantic import BaseModel


class TextFrameShape(BaseModel):
    slide_id: int
    shape_id: int

    text: str

    font_name: str | None = None
    font_size: float | None = None
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    strike: bool | None = None
    color: tuple | None = None

    text_frame: Any


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
