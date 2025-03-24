from pydantic import BaseModel


class FontElement(BaseModel):
    text: str
    font_name: str | None = None
    font_size: float | None = None
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    strike: bool | None = None
    color: tuple | None = None


class ParagraphFontElement(BaseModel):
    alignment: str | None = None
    fonts: list[FontElement]


class ShapeFontElement(BaseModel):
    shape_id: int | None = None
    shape_name: str | None = None
    paragraphs: list[ParagraphFontElement]
