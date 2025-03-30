from loguru import logger
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.shapes.autoshape import Shape

from pptx_manager.image import ImageManager
from pptx_manager.models import CreateShapeOpts, CreateTextFrameOpts, ShapeType, TextFrameOpts
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
            MSO_SHAPE_TYPE.AUTO_SHAPE: self._parse_text_shape,
            MSO_SHAPE_TYPE.TEXT_BOX: self._parse_text_shape,
            MSO_SHAPE_TYPE.GROUP: self.parse_group_shape,
        }

        self.slide_count = len(self.pres.slides)

        self.parse_presentation()

    def parse_presentation(self):
        logger.debug(f'parse_presentation calls with len = {len(self.pres.slides)}')
        slide_id = 1

        for slide in self.pres.slides:
            self.parse_slide(slide_id, slide)
            for shape in slide.shapes:
                parse_fn = self.parse_choice.get(shape.shape_type)
                if parse_fn is not None:
                    result = parse_fn(slide_id, self.get_shape_count(slide_id), shape)
                    if result is not None:
                        self.increase_shape_count(slide_id, 1)
            slide_id += 1

    def parse_group_shape(self, slide_id: int, shape_id: int, shape: Shape):
        for group_shape in shape.shapes:
            parse_fn = self.parse_choice.get(group_shape.shape_type)
            if parse_fn is not None:
                result = parse_fn(slide_id, self.get_shape_count(slide_id), group_shape)
                if result is not None:
                    self.increase_shape_count(slide_id, 1)

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

    def create_text_shape(self, slide_id: int, opts: CreateTextFrameOpts) -> str:
        """
        Creates a new text shape on a specified slide with the given options.

        This function adds a text shape to the presentation and applies the specified
        formatting options including position, size, text content, and text styling.

        Args:
            slide_id: The ID of the slide where the text shape will be created.
            opts: Configuration options for the text shape, including:
                - left: Left position of the shape in pixels
                - top: Top position of the shape in pixels
                - height: Height of the shape in pixels
                - width: Width of the shape in pixels
                - text: The text content to display in the shape
                - color: RGB tuple (r,g,b) for text color
                - size: Font size in points
                - bold: Whether text should be bold (True/False)
                - italic: Whether text should be italic (True/False)
                - underline: Whether text should be underlined (True/False)

        Returns:
            str: A status message indicating the result of the operation:
        """
        logger.debug(f'create_text_shape calls with parameters slide_id={slide_id}, opts={opts}')
        shape, shape_id = self._add_shape_on_slide(
            slide_id,
            CreateShapeOpts(
                left=opts.left,
                top=opts.top,
                height=opts.height,
                width=opts.width,
                type=ShapeType.TEXT,
            ),
        )

        self._parse_text_shape(slide_id, shape_id, shape)

        self.update_text_frame_shape(
            slide_id,
            shape_id,
            TextFrameOpts(
                text=opts.text,
                color=opts.color,
                size=opts.size,
                bold=opts.bold,
                italic=opts.italic,
                underline=opts.underline,
            ),
        )

        return 'Success'

    def delete_text_shape(self, slide_id: int, shape_id: int) -> str:
        """
        Deletes a text shape from a specified slide.

        This function removes the text shape with the given shape_id from the slide
        with the given slide_id. It attempts to remove the shape from both the internal
        tracking system and the actual presentation object.

        Args:
            slide_id: The ID of the slide containing the text shape to delete.
            shape_id: The ID of the text shape to delete.

        Returns:
            str: A status message indicating the result of the operation:
                - "Success" if the shape was successfully deleted
                - "The unknown shape id" if the shape was not found or could not be deleted
        """
        logger.debug(f'delete_text_shape calls with parameters slide_id={slide_id}, shape_id={shape_id}')
        shape = self._get_text_frame_shape(slide_id, shape_id)
        self._delete_text_frame_shape(slide_id, shape_id)

        try:
            el = shape.shape_manager.element
            parent = el.getparent()
            parent.remove(el)
        except AttributeError:
            logger.warning(f'shape {shape_id} not found in shape_manager')
            return 'The unknown shape id'

        return 'Success'

    def save(self, path):
        self.pres.save(path)


if __name__ == '__main__':
    pr = PPTXManager('../test_data/test_dit.pptx')

    pr.add_slide_at_position(13, layout_index=0, title='Inserted Slide 1')

    pr.create_text_shape(
        -1,
        CreateTextFrameOpts(
            left=300,
            top=200,
            height=500,
            width=600,
            text='evaluate test adding',
            color=(0, 0, 0),
            size=25,
            bold=True,
            italic=False,
            underline=False,
        ),
    )

    pr.update_text_frame_shape(
        -1,
        1,
        TextFrameOpts(
            text='test after',
            bold=True,
            italic=True,
            size=70,
        ),
    )

    pr.save('test_create.pptx')
