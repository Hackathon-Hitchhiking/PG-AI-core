from loguru import logger
from pptx import Presentation


class SlideManager:
    def __init__(self):
        self.pres = None

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

    def add_slide_at_position(self, position: int, layout_index: int = 0, title: str = None) -> None:
        """
        Add a new slide at a specific position in the presentation.

        Args:
            position (int): The 1-based index where the slide should be inserted.
            layout_index (int): The index of the slide layout to use.
            title (str): Optional title text for the slide.

        Raises:
            ValueError: If no presentation is loaded, or if the position or layout index is invalid.
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

    def save_presentation(self, file_path: str) -> None:
        if not self.pres:
            raise ValueError('Presentation not loaded. Use `load_presentation` first.')

        self.pres.save(file_path)
        logger.info(f'Presentation saved to {file_path}')

    def test(self, source: str) -> None:
        self.load_presentation(source)
        self.add_slide_at_position(1, layout_index=0, title='Inserted Slide 1')
        self.add_slide_at_position(3, layout_index=0, title='Inserted Slide at Position 3')
        output_path = '../test_data/test_output.pptx'
        self.save_presentation(output_path)


if __name__ == '__main__':
    sm = SlideManager()
    sm.test('../test_data/test_dit.pptx')
