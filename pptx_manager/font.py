import colorsys

import numpy as np

from loguru import logger
from PIL import ImageColor
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml import parse_xml
from pptx.shapes.autoshape import Shape
from pptx.slide import Slides
from pptx.text.text import Font

from pptx_manager.models import ParagraphRunElement, RunElement, ShapeRunElement, TextFrameShape
from pptx_manager.utils import get_all_methods, hex_to_rgb


class FontManager:
    def __init__(self):
        self.pres = None

        self.text_frames = []

    def get_all_runs(self, slide: Slides) -> list[ShapeRunElement]:
        shape_info_list = []

        for shape in slide.shapes:
            print(type(shape))
            if not shape.has_text_frame:
                continue

            paragraphs_list = []
            for paragraph in shape.text_frame.paragraphs:
                font_list = []

                logger.debug(f'methods = {get_all_methods(paragraph.text)}')

                for run in paragraph.runs:
                    font = run.font

                    font_size = 12

                    if font.size is not None:
                        font_size = font.size.pt

                    font_color = self._get_font_color(slide, font)

                    run_info = RunElement(
                        text=run.text,
                        font_name=font.name,
                        font_size=font_size,
                        bold=font.bold,
                        italic=font.italic,
                        underline=font.underline,
                        color=font_color,
                    )
                    font_list.append(run_info)

                align_name = str(paragraph.alignment) if paragraph.alignment else None

                paragraphs_list.append(ParagraphRunElement(alignment=align_name, fonts=font_list))

            shape_info = ShapeRunElement(shape_id=shape.shape_id, shape_name=shape.name, paragraphs=paragraphs_list)
            shape_info_list.append(shape_info)

        return shape_info_list

    def _get_font_color(self, slide: Slides, font: Font) -> tuple:
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

    def parse_text_frame_shape(
        self, slide_id: int, shape_id: int, slide: Slides, shape: Shape
    ) -> None | TextFrameShape:
        if shape.text == '':
            return None
        try:
            text_frame = shape.text_frame

            text = shape.text

            paragraph = text_frame.paragraphs[0]
            run = paragraph.runs[0]
            font = run.font

            font_size = self._get_font_size(font)
            font_color = self._get_font_color(slide, font)
            font_name = font.name
            bold = font.bold
            italic = font.italic
            underline = font.underline
        except Exception as e:
            logger.warning(f'error parsing text_frame = {e}')
            return None

        text_frame = TextFrameShape(
            slide_id=slide_id,
            shape_id=shape_id,
            text=text,
            font_name=font_name,
            font_size=font_size,
            bold=bold,
            italic=italic,
            underline=underline,
            color=font_color,
            text_frame=text_frame,
        )

        self.text_frames.append(text_frame)

        return text_frame

    def test(self, source):
        self.pres = Presentation(source)

        shape_id = 1
        slide_id = 1
        for slide in self.pres.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    res = self.parse_text_frame_shape(slide_id, shape_id, slide, shape)
                    if res is not None:
                        shape_id += 1
                        slide_id += 1

        logger.debug(f'frames = {self.text_frames}')


if __name__ == '__main__':
    fm = FontManager()

    fm.test('../test_sources/test_dit.pptx')
