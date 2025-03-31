import colorsys

from collections import defaultdict
from collections.abc import Iterator

import numpy as np

from loguru import logger
from PIL import ImageColor
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml import parse_xml
from pptx.shapes.autoshape import Shape
from pptx.slide import Slide
from pptx.text.text import Font
from pptx.util import Pt

from pptx_manager.models import TextFrameOpts, TextFrameShape
from pptx_manager.utils import get_slide_from_shape, hex_to_rgb


class TextFrameManager:
    def __init__(self):
        self.pres = None

        self.text_frame_shapes: defaultdict[int, list[TextFrameShape]] = defaultdict(
            list[TextFrameShape]
        )  # slide_id -> text_frame_shape

    def get_text_frame_json(self, slide_id: int) -> list[dict]:
        return [
            shape.model_dump(exclude={'text_manager', 'font_manager', 'shape_manager'})
            for shape in self.text_frame_shapes[slide_id]
        ]

    def get_all_text_frame_json(self) -> dict:
        text_frame_json = {}

        for slide_id, shapes in self.text_frame_shapes.items():
            shapes_json = []
            for shape in shapes:
                shapes_json.append(shape.model_dump(exclude={'text_manager', 'font_manager'}))
            text_frame_json[slide_id] = shapes_json

        return text_frame_json

    def _get_font_color(self, slide: Slide, font: Font) -> tuple:
        # https://stackoverflow.com/questions/54692768/python-pptx-read-font-color
        font_color = font.color

        try:
            srgb = font_color.rgb
        except AttributeError:
            srgb = None

        if srgb is not None:
            return hex_to_rgb(str(srgb))

        try:
            theme_color = font_color.theme_color
        except AttributeError:
            theme_color = None

        # if the theme color and rgb is None its the black color
        if theme_color is None:
            return 0, 0, 0

        brightness = font_color.brightness

        accent = theme_color.xml_value
        xpath = f'a:themeElements/a:clrScheme/a:{accent}/a:srgbClr/@val'

        slide_master_part = slide.slide_layout.slide_master.part
        theme_part = slide_master_part.part_related_by(RT.THEME)
        theme = parse_xml(theme_part.blob)

        try:
            hex_color = theme.xpath(xpath)[0]
        except IndexError:
            return 0, 0, 0

        srgb = np.array(ImageColor.getcolor(f'#{hex_color}', 'RGB'))

        srgb = srgb / 255
        h, luminance, s = colorsys.rgb_to_hls(*srgb)
        lum_mod = 100000 * (1 - brightness)
        lum_off = 100000 * brightness
        luminance = luminance * (lum_mod / 100000) + (lum_off / 100000)
        srgb = np.array(colorsys.hls_to_rgb(h, luminance, s))
        srgb = (srgb * 255).round(0).astype(int)

        return tuple(srgb)

    def _get_font_size(self, font: Font) -> float:
        font_size = 12

        if font.size is not None:
            font_size = font.size.pt

        return font_size

    def update_text_frame_shape(self, slide_id: int, shape_id: int | None, opts: TextFrameOpts | dict):
        """Updates the properties of a text frame shape in a specific slide.

        Modifies various text attributes, including content, color, size, and style (bold, italic, underline).

        Args:
            slide_id (int): The ID of the slide containing the text frame shape to be updated.
            shape_id (int | None): The ID of the shape to be updated. If None, updates all shapes on the slide.
            opts (dict): An object containing the options for updating the text frame.
                -   text (str | None, optional): The new text content for the shape.
                -   color (list[int] | None, optional): The new text color as an RGB tuple.
                -   size (int | None, optional): The new font size for the text.
                -   bold (bool | None, optional): Whether to set the text to bold.
                -   italic (bool | None, optional): Whether to set the text to italic.
                -   underline (bool | None, optional): Whether to underline the text.

        Notes:
            -   Only the attributes specified in `opts` will be updated.
            -   If `opts` is a dictionary, it will be converted to `UpdateTextFrameOpts`.
            -   Each attribute update is handled by a separate internal method.
            -   All attributes in `UpdateTextFrameOpts` are optional and default to None.

        Returns:
            str: Confirmation of the completion of the task
        """
        logger.debug(f'update_text_frame_shape calls with parametrs: {slide_id, shape_id, opts}')

        if isinstance(opts, dict):
            opts = TextFrameOpts(**opts)

        if opts.text is not None:
            self._update_text_text_frame_shape(slide_id, shape_id, opts.text)

        if opts.color is not None:
            self._update_color_text_frame_shape(slide_id, shape_id, opts.color)

        if opts.italic is not None:
            self._update_italic_text_frame_shape(slide_id, shape_id, opts.italic)

        if opts.underline is not None:
            self._update_underline_text_frame_shape(slide_id, shape_id, opts.underline)

        if opts.bold is not None:
            self._update_bold_text_frame_shape(slide_id, shape_id, opts.bold)

        if opts.size is not None:
            self._update_size_text_frame_shape(slide_id, shape_id, opts.size)

        return 'Done'

    def _update_text_text_frame_shape(self, slide_id: int, shape_id: int | None, new_text: str) -> None:
        for frame in self._get_text_frame(slide_id, shape_id):
            frame.text_manager.paragraphs[0].runs[0].text = new_text
            frame.text = new_text

    def _update_color_text_frame_shape(
        self, slide_id: int, shape_id: int | None, new_color: tuple[int, int, int]
    ) -> None:
        for frame in self._get_text_frame(slide_id, shape_id):
            frame.font_manager.color.rgb = RGBColor(new_color[0], new_color[1], new_color[2])
            frame.color = (new_color[0], new_color[1], new_color[2])

    def _update_bold_text_frame_shape(self, slide_id: int, shape_id: int | None, bold: bool) -> None:
        for frame in self._get_text_frame(slide_id, shape_id):
            frame.font_manager.bold = bold
            frame.bold = bold

    def _update_italic_text_frame_shape(self, slide_id: int, shape_id: int | None, italic: bool) -> None:
        for frame in self._get_text_frame(slide_id, shape_id):
            frame.font_manager.italic = italic
            frame.italic = italic

    def _update_underline_text_frame_shape(self, slide_id: int, shape_id: int | None, underline: bool) -> None:
        for frame in self._get_text_frame(slide_id, shape_id):
            frame.font_manager.underline = underline
            frame.underline = underline

    def _update_size_text_frame_shape(self, slide_id: int, shape_id: int | None, new_size: int) -> None:
        for frame in self._get_text_frame(slide_id, shape_id):
            frame.font_manager.size = Pt(new_size)
            frame.font_size = new_size

    def _delete_text_frame_shape(self, slide_id: int, shape_id: int) -> None:
        frames = self.text_frame_shapes[slide_id]
        for index, frame in enumerate(frames):
            if frame.shape_id == shape_id:
                del frames[index]

    def _get_text_frame(self, slide_id: int, shape_id: int | None) -> Iterator[TextFrameShape]:
        if slide_id < 0:
            slide_id = sorted(list(self.text_frame_shapes.keys()))[slide_id]
        frames = self.text_frame_shapes[slide_id]
        for frame in frames:
            if shape_id is not None and shape_id != frame.shape_id:
                continue
            yield frame

    def _copy_font_properties(self, copy_font: Font, base_font: Font, slide: Slide) -> Font:
        copy_font.name = base_font.name
        copy_font.size = Pt(self._get_font_size(base_font))
        copy_font.bold = base_font.bold
        copy_font.italic = base_font.italic
        copy_font.underline = base_font.underline
        copy_font.language_id = base_font.language_id

        copy_font.color.rgb = RGBColor(*self._get_font_color(slide, base_font))

        return copy_font

    def _get_text_frame_shape(self, slide_id: int, shape_id: int | None) -> TextFrameShape | None:
        for frame in self._get_text_frame(slide_id, shape_id):
            return frame

    def _parse_text_shape(self, slide_id: int, shape_id: int, shape: Shape) -> TextFrameShape | None:
        if shape.text == '' and shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
            return None
        try:
            text_frame = shape.text_frame

            first_run_font = None
            full_text_lines = []
            for p_i, paragraph in enumerate(text_frame.paragraphs):
                paragraph_text = ''
                for r_i, run in enumerate(paragraph.runs):
                    if first_run_font is None:
                        first_run_font = run.font
                    paragraph_text += run.text
                full_text_lines.append(paragraph_text)

            merged_text = '\n'.join(full_text_lines)

            if first_run_font is None:
                logger.warning(
                    f'Could not find base font for text frame with slide_id {slide_id}, shape_id {shape_id} using base shape'
                )

            text_frame.clear()

            paragraph = text_frame.paragraphs[0]
            single_run = paragraph.add_run()
            single_run.text = merged_text
            text_frame_font = single_run.font

            slide = get_slide_from_shape(shape)

            if first_run_font is None:
                text_frame_font = self._copy_font_properties(text_frame_font, single_run.font, slide)
            else:
                text_frame_font = self._copy_font_properties(text_frame_font, first_run_font, slide)

            font_size = self._get_font_size(text_frame_font)
            font_color = self._get_font_color(slide, text_frame_font)
            font_name = text_frame_font.name
            bold = text_frame_font.bold
            italic = text_frame_font.italic
            underline = text_frame_font.underline

            text = merged_text

        except Exception as e:
            logger.warning(f'error parsing text_frame = {e}, text_shape = {shape.text}')
            return None

        text_frame_shape = TextFrameShape(
            shape_id=shape_id,
            text=text,
            font_name=font_name,
            font_size=font_size,
            bold=bold,
            italic=italic,
            underline=underline,
            color=font_color,
            text_manager=text_frame,
            font_manager=text_frame_font,
            shape_manager=shape,
        )

        if slide_id < 0:
            slide_id = len(self.text_frame_shapes) - slide_id

        self.text_frame_shapes[slide_id].append(text_frame_shape)

        return text_frame_shape

    def test(self, source):
        self.pres = Presentation(source)

        slide_id = 1
        for slide in self.pres.slides:
            shape_id = 1
            for shape in slide.shapes:
                if shape.has_text_frame:
                    res = self._parse_text_shape(slide_id, shape_id, shape)
                    if res is not None:
                        shape_id += 1
            slide_id += 1

        self.update_text_frame_shape(
            -1,
            2,
            {'text': 'тест', 'italic': True, 'size': 50},
        )

        self.pres.save('test.pptx')


if __name__ == '__main__':
    fm = TextFrameManager()

    fm.test('../test_data/test_dit.pptx')
