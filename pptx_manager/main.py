from loguru import logger
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.shapes.autoshape import Shape

from pptx_manager.image import ImageManager
from pptx_manager.shape import ShapeManager
from pptx_manager.slide import SlideManager
from pptx_manager.table import TableManager
from pptx_manager.text import TextFrameManager


class PPTXManager(
    TextFrameManager,
    ImageManager,
    ShapeManager,
    SlideManager,
    TableManager,
):
    def __init__(self, source):
        super(PPTXManager, self).__init__()
        TextFrameManager.__init__(self)
        ImageManager.__init__(self)
        ShapeManager.__init__(self)
        SlideManager.__init__(self)
        TableManager.__init__(self)

        self.pres = Presentation(source)

        self.parse_choice = {
            MSO_SHAPE_TYPE.PICTURE: self.parse_image_shape,
            MSO_SHAPE_TYPE.AUTO_SHAPE: self.parse_text_shape,
            MSO_SHAPE_TYPE.TEXT_BOX: self.parse_text_shape,
            MSO_SHAPE_TYPE.GROUP: self.parse_shape,
        }

        self.slide_count = len(self.pres.slides)

        self.slide_metadata = {}  # slide_id -> count of the shape id

        self.parse_presentation()

    def parse_presentation(self):
        logger.debug(f'Parsing Presentation, len = {len(self.pres.slides)}')
        slide_id = 1
        for slide in self.pres.slides:
            self.slide_metadata[slide_id] = 1
            for shape in slide.shapes:
                parse_fn = self.parse_choice.get(shape.shape_type)
                if parse_fn is not None:
                    result = parse_fn(slide_id, self.slide_metadata[slide_id], shape)
                    if result is not None:
                        self.slide_metadata[slide_id] += 1
            slide_id += 1

    def parse_shape(self, slide_id: int, shape_id: int, shape: Shape):
        for group_shape in shape.shapes:
            parse_fn = self.parse_choice.get(group_shape.shape_type)
            if parse_fn is not None:
                result = parse_fn(slide_id, self.slide_metadata[slide_id], group_shape)
                if result is not None:
                    self.slide_metadata[slide_id] += 1

    def get_json_schema(self) -> dict:
        slide_json = {}
        for slide_id in range(1, self.slide_count + 1):
            text_json = self.get_text_frame_json(slide_id)
            image_json = self.get_image_json(slide_id)

            slide_json[slide_id] = {
                'text': text_json,
                'image': image_json,
            }

        return slide_json

    def save(self, path):
        self.pres.save(path)


if __name__ == '__main__':
    pr = PPTXManager('../test_data/test_dit.pptx')

    pr.parse_presentation()

    print(pr.get_json_schema())
