from pptx import Presentation
from pptx.slide import Slides

from pptx_manager.models import FontElement, ParagraphFontElement, ShapeFontElement
from pptx_manager.utils import get_all_methods, hex_to_rgb

from pptx.dml.color import ColorFormat, RGBColor
from pptx.enum.dml import MSO_THEME_COLOR
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml import parse_xml
from PIL import ImageColor
import numpy as np
import colorsys

from loguru import logger

class FontManager:
    def __init__(self):
        self.pres = None

    def get_all_fonts(self, slide: Slides) -> list[ShapeFontElement]:
        shape_info_list = []

        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue

            paragraphs_list = []
            for paragraph in shape.text_frame.paragraphs:
                font_list = []
                for run in paragraph.runs:
                    font = run.font

                    logger.debug(f"font_name = {run.text}")
                    logger.debug(f"methodds = {get_all_methods(run.part)}")

                    font_size = font.size.pt

                    font_color = self._get_font_color(slide, font.color)

                    run_info = FontElement(
                        # TODO maybe here returning the language ID
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

                paragraphs_list.append(ParagraphFontElement(
                    alignment=align_name,
                    fonts=font_list
                ))

            shape_info = ShapeFontElement(
                shape_id=shape.shape_id,
                shape_name=shape.name,
                paragraphs=paragraphs_list
            )
            shape_info_list.append(shape_info)

        return shape_info_list

    def _get_font_size(self):
        # https://github.com/scanny/python-pptx/issues/378#issuecomment-2730305545

    def _get_font_color(self, slide: Slides, font_color: ColorFormat,) -> tuple:
        # https://stackoverflow.com/questions/54692768/python-pptx-read-font-color
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
        xpath = 'a:themeElements/a:clrScheme/a:{}/a:srgbClr/@val'.format(accent)

        slide_master_part = slide.slide_layout.slide_master.part
        theme_part = slide_master_part.part_related_by(RT.THEME)
        theme = parse_xml(theme_part.blob)

        hex_color = theme.xpath(xpath)[0]

        srgb = np.array(ImageColor.getcolor('#{}'.format(hex_color), 'RGB'))

        srgb = srgb / 255
        h, luminance, s = colorsys.rgb_to_hls(*srgb)
        lum_mod = 100000 * (1 - brightness)
        lum_off = 100000 * brightness
        luminance = luminance * (lum_mod / 100000) + (lum_off / 100000)
        srgb = np.array(colorsys.hls_to_rgb(h, luminance, s))
        srgb = (srgb * 255).round(0).astype(int)

        return tuple(srgb)

    def test(self, source):
        self.pres = Presentation(source)

        for slide in self.pres.slides:
            logger.debug(self.get_all_fonts(slide))




if __name__ == '__main__':
    fm = FontManager()

    fm.test("../test_sources/test_dit.pptx")


