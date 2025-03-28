from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

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
        }

        self.slide_count = 1

        self.parse_presentation()

    def parse_presentation(self):
        for slide in self.pres.slides:
            shape_id = 1
            for shape in slide.shapes:
                parse_fn = self.parse_choice.get(shape.shape_type)
                if parse_fn is not None:
                    parse_fn(self.slide_count, shape_id, shape)
                shape_id += 1
            self.slide_count += 1

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


if __name__ == '__main__':
    pr = PPTXManager('../test_sources/test_dit.pptx')

    pr.parse_presentation()

    print(pr.get_json_schema())
