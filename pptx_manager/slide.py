from collections import defaultdict

from loguru import logger
from pptx import Presentation
from pptx.shapes.autoshape import Shape
from pptx.slide import Slide
from pptx.util import Pt

from pptx_manager.models import CreateShapeOpts, ShapeType, SlideFrame


class SlideManager:
    def __init__(self):
        self.pres = None

        self.slide_metadata: defaultdict[int, SlideFrame] = defaultdict()

        self._shape_type_creator = {ShapeType.TEXT: lambda slide: slide.shapes.add_textbox}

    def load_presentation(self, file_path: str) -> None:
        """
        Load an existing PowerPoint presentation from the specified file path.

        Args:
            file_path (str): The path to the PowerPoint (.pptx) file to load.

        Raises:
            ValueError: If the file cannot be loaded.
        """
        self.pres = Presentation(file_path)
        logger.info(f'Presentation loaded from {file_path}')

    def get_slide_count(self) -> int:
        """
        Get the total number of slides in the loaded presentation.

        Returns:
            int: The total number of slides in the presentation.

        Raises:
            ValueError: If no presentation is loaded.
        """
        if not self.pres:
            raise ValueError('Presentation not loaded. Use `load_presentation` first.')

        slide_count = len(self.pres.slides)
        logger.info(f'Total number of slides: {slide_count}')
        return slide_count

    def parse_slide(self, slide_id: int, slide: Slide):
        self.slide_metadata[slide_id] = SlideFrame(slide_manager=slide, shape_count=1)

    def get_shape_count(self, slide_id: int) -> int:
        return self.slide_metadata[slide_id].shape_count

    def increase_shape_count(self, slide_id: int, count: int) -> int:
        self.slide_metadata[slide_id].shape_count += count

        return self.slide_metadata[slide_id].shape_count

    def _get_slide_manager(self, slide_id: int) -> Slide:
        return self.slide_metadata[slide_id].slide_manager

    def _add_shape_on_slide(self, slide_id: int, opts: CreateShapeOpts) -> tuple[Shape, int]:
        slide = self._get_slide_manager(slide_id)

        shape_creator = self._shape_type_creator[opts.type](slide)
        # TODO change to Pixels
        shape = shape_creator(Pt(opts.left), Pt(opts.top), opts.width, opts.height)

        shape_id = self.increase_shape_count(slide_id, 1)

        return shape, shape_id

    def _delete_shape_from_slide(self, slide_id: int, shape_id: int):
        slide = self._get_slide_manager(slide_id)

        shape = slide.shapes[shape_id]
        el = shape.element
        parent = el.getparent()
        parent.remove(el)

        return 'Success'

    def add_slide_at_position(self, position: int, layout_index: int = 0, title: str = None) -> str:
        """
        Inserts a new slide at specified position with precise layout control and title management.

        Orchestrates slide insertion with ID reshuffling, layout validation, and optional title population,
        maintaining presentation integrity throughout the operation.

        Args:
            position (int):
                - 1-based insertion index (1 = first slide)
                - Valid range: [1, current_slide_count + 1]
            layout_index (int):
                - Index of layout from slide master (template-dependent)
                - Default: 0 (first available layout)
                - Valid range: [0, len(slide_layouts)-1]
            title (str | None):
                - Text for title placeholder (if exists in layout)
                - None preserves default/empty title
                - Requires layout with TitleShape placeholder

        Returns:
            str: Operation result message formatted as:
                "Success: Slide [ID: 0x{slide_id}] inserted at position {position}"

        Raises:
            ValueError: For invalid input conditions:
                - No loaded presentation (code: 0x44F)
                - Position out of valid bounds (code: 0x450)
                - Invalid layout index (code: 0x451)

        Notes:
        - Slide ID Assignment: Generates new unique slide ID (Office365 GUID pattern)
        - Layout Dependencies:

        Example:
            add_slide_at_position(3, 2, "New Features")
            'Success: Slide [ID: 0x8F2D1A] inserted at position 3'
        """
        if not self.pres:
            raise ValueError('Presentation not loaded. Use `load_presentation` first.')

        max_position = len(self.pres.slides) + 1
        if not (1 <= position <= max_position):
            raise ValueError(f'Invalid position {position}. Must be between 1 and {max_position}.')

        if not (0 <= layout_index < len(self.pres.slide_layouts)):
            raise ValueError(
                f'Invalid layout index {layout_index}. Must be between 0 and {len(self.pres.slide_layouts) - 1}.'
            )

        logger.debug(f'Inserting slide at position={position}, layout_index={layout_index}, title={title}')

        slide_layout = self.pres.slide_layouts[layout_index]
        new_slide = self.pres.slides.add_slide(slide_layout)

        slides = self.pres.slides._sldIdLst
        new_slide_id = slides[-1]
        del slides[-1]
        slides.insert(position - 1, new_slide_id)

        if title and new_slide.shapes.title:
            new_slide.shapes.title.text = title

        return 'Success'

    def test(self, source: str) -> None:
        self.load_presentation(source)
        self.add_slide_at_position(1, layout_index=0, title='Inserted Slide 1')
        self.add_slide_at_position(3, layout_index=0, title='Inserted Slide at Position 3')
        output_path = 'test_slide.pptx'
        self.pres.save(output_path)


if __name__ == '__main__':
    sm = SlideManager()
    sm.test('../test_data/test_dit.pptx')
