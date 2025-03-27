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
    color: tuple[int, int, int] | None = None
    size: int | None = None
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None


class UpdateImageFrameOpts(BaseModel):
    width: int | None = None
    height: int | None = None

    left: int | None = None
    top: int | None = None


class ImageFrameShape(BaseModel):
    shape_id: int

    width: int
    height: int

    left: int
    top: int

    blob: bytes

    shape_manager: Any
