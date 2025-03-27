import io

from collections import defaultdict

from loguru import logger
from PIL import Image as PilImage
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn
from pptx.parts.image import Image
from pptx.shapes.autoshape import Shape
from pptx.slide import Slides

from pptx_manager.models import ImageFrameShape


def image_to_byte_array(image: Image):
    img_byte_arr = io.BytesIO()
    image.save(img_byte_arr, format=image.format)
    img_byte_arr = img_byte_arr.getvalue()
    return img_byte_arr


class ImageManager:
    def __init__(self):
        self._pres = None

        self.image_frame_shapes = defaultdict(list[ImageFrameShape])  # slide_id -> image_frame_shapes

    def get_image_json(self, slide_id: int):
        return [shape.model_dump(exclude={'blob', 'shape_manager'}) for shape in self.image_frame_shapes[slide_id]]

    def replace_image(self, slide: Slides, slide_id: int, shape_id: int, new_picture: bytes):
        shapes = self.image_frame_shapes[slide_id]

        for shape in shapes:
            if shape.shape_id == shape_id:
                shape.blob = new_picture

                slide_part, rId = shape.shape_manager.part, shape.shape_manager._element.blip_rId
                (new_impart_part, new_rid) = slide_part.get_or_add_image_part(io.BytesIO(shape.blob))
                blip_fill = shape.shape_manager._element.blipFill
                blip = blip_fill.find(qn('a:blip'))
                blip.set(qn('r:embed'), new_rid)

    def parse_image_shape(self, slide_id: int, shape_id: int, shape: Shape) -> defaultdict[str, list[ImageFrameShape]]:
        width = shape.width
        height = shape.height
        image = shape.image
        image_bytes = image.blob

        self.image_frame_shapes[slide_id].append(
            ImageFrameShape(
                shape_id=shape_id,
                width=width,
                height=height,
                left=image.left,
                right=image.right,
                blob=image_bytes,
                shape_manager=shape,
            )
        )

        return self.image_frame_shapes

    def test(self, source):
        self._pres = Presentation(source)

        slide_id = 1
        for slide in self._pres.slides:
            shape_id = 1
            for shape in slide.shapes:
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    res = self.parse_image_shape(slide_id, shape_id, shape)
                    if res is not None:
                        shape_id += 1
            slide_id += 1

        logger.debug(self.get_image_json(2))

        self._pres.save('test_new_img.pptx')


if __name__ == '__main__':
    im = ImageManager()

    im.test('../test_sources/test_dit.pptx')

    new_img = PilImage.open('../test_sources/test_image.png')

    im.replace_image(2, 2, image_to_byte_array(new_img))
