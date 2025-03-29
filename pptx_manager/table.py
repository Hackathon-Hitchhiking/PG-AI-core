from loguru import logger
from pptx import Presentation
from pptx.util import Inches


class TableManager:
    def __init__(self):
        self.pres = None

    def load_presentation(self, file_path: str) -> None:
        """
        Loads a PowerPoint presentation from the specified file.

        Args:
            file_path (str): Path to the PowerPoint (.pptx) file.

        Raises:
            ValueError: If the presentation cannot be loaded.
        """
        self.pres = Presentation(file_path)
        logger.info(f'Loaded presentation: {file_path}')

    def get_table_indexes(self, slide_num: int) -> list[int]:
        """
        Returns indexes of all tables on the specified slide.

        Args:
            slide_num (int): 1-based slide number.

        Returns:
            List[int]: List of shape indexes for tables on the slide.

        Raises:
            ValueError: If no presentation is loaded or if the slide number is invalid.
        """
        if not self.pres:
            raise ValueError('Presentation not loaded. Use `load_presentation` first.')

        if not (1 <= slide_num <= len(self.pres.slides)):
            raise ValueError(f'Invalid slide number: {slide_num}')

        slide = self.pres.slides[slide_num - 1]
        return [i for i, shape in enumerate(slide.shapes) if shape.has_table]

    def parse_tables(self) -> list[dict]:
        """
        Extracts data from all tables in the presentation.

        Returns:
            List[Dict]: A list of dictionaries containing table metadata and data.
                        Each dictionary has the format:
                        {
                            'slide_num': int,
                            'table_index': int,
                            'data': List[List[str]]
                        }

        Raises:
            ValueError: If no presentation is loaded.
        """
        if not self.pres:
            raise ValueError('Presentation not loaded. Use `load_presentation` first.')

        tables = []
        for slide_num, slide in enumerate(self.pres.slides, 1):
            table_indexes = self.get_table_indexes(slide_num)
            for table_idx in table_indexes:
                table = slide.shapes[table_idx].table
                tables.append(
                    {
                        'slide_num': slide_num,
                        'table_index': table_idx,
                        'data': [[cell.text.strip() for cell in row.cells] for row in table.rows],
                    }
                )
                logger.debug(f'Parsed table on slide {slide_num}, index {table_idx}')

        logger.info(f'Total tables parsed: {len(tables)}')
        return tables

    def create_table(
        self,
        slide_num: int,
        rows: int,
        cols: int,
        left: float = 1.0,
        top: float = 1.0,
        width: float = 6.0,
        height: float = 4.0,
    ) -> int:
        """
        Creates a new table on the specified slide.

        Args:
            slide_num (int): Slide number (1-based index).
            rows (int): Number of rows in the table.
            cols (int): Number of columns in the table.
            left (float): Left position of the table in inches.
            top (float): Top position of the table in inches.
            width (float): Width of the table in inches.
            height (float): Height of the table in inches.

        Returns:
            int: The index of the created table on the slide.

        Raises:
            ValueError: If no presentation is loaded or if the slide number is invalid.
        """
        if not self.pres:
            raise ValueError('Presentation not loaded. Use `load_presentation` first.')

        if not (1 <= slide_num <= len(self.pres.slides)):
            raise ValueError(f'Invalid slide number: {slide_num}')

        slide = self.pres.slides[slide_num - 1]
        table_shape = slide.shapes.add_table(rows, cols, Inches(left), Inches(top), Inches(width), Inches(height))

        logger.info(f'Created a {rows}x{cols} table on slide {slide_num}')

        return len(slide.shapes) - 1

    def update_cell(self, slide_num: int, table_index: int, row: int, col: int, text: str) -> None:
        """
        Updates content of a specific cell in a table.

        Args:
            slide_num (int): Slide number (1-based index).
            table_index (int): Index of the table on the slide.
            row (int): Row number (0-based index).
            col (int): Column number (0-based index).
            text (str): New content for the cell.

        Raises:
            ValueError: If no presentation is loaded or if parameters are invalid.
        """
        if not self.pres:
            raise ValueError('Presentation not loaded. Use `load_presentation` first.')

        try:
            slide = self.pres.slides[slide_num - 1]
            shape = slide.shapes[table_index]

            if not shape.has_table:
                raise ValueError('The specified shape is not a table.')

            table = shape.table

            if row >= len(table.rows) or col >= len(table.columns):
                raise IndexError('Row or column index out of bounds.')

            cell = table.cell(row, col)
            cell.text_frame.clear()
            cell.text_frame.paragraphs[0].text = text
            logger.debug(f'Updated cell [{row},{col}] in table {table_index} on slide {slide_num}')

        except IndexError as e:
            logger.error(f'Invalid cell position ({row}, {col}): {str(e)}')
            raise ValueError('Invalid cell position.') from e

    def save_presentation(self, file_path: str) -> None:
        """
        Saves the modified presentation to a specified file path.

        Args:
            file_path (str): Path to save the PowerPoint (.pptx) file.

        Raises:
            ValueError: If no presentation is loaded.
        """
        if not self.pres:
            raise ValueError('Presentation not loaded. Use `load_presentation` first.')

        self.pres.save(file_path)
        logger.info(f'Presentation saved to {file_path}')

    def test(self, source: str) -> None:
        """
        Test function to demonstrate working with tables.

        Args:
            source (str): Path to the source PowerPoint (.pptx) file.

        Raises:
            Exception: If any operation fails during testing.
        """
        try:
            self.load_presentation(source)

            existing_tables = self.parse_tables()

            logger.info(f'Found {len(existing_tables)} existing tables')

            new_table_index = self.create_table(1, 3, 4)

            if existing_tables:
                first_table = existing_tables[0]
                self.update_cell(
                    first_table['slide_num'],
                    first_table['table_index'],
                    0,
                    0,
                    'Обновились или я тебя удалю, это будет больно',
                )

            self.update_cell(1, new_table_index, 0, 0, 'New Table Content')

            self.save_presentation('../test_data/table_test.pptx')

            logger.success('Test completed successfully.')

        except Exception as e:
            logger.error(f'Test failed with error: {str(e)}')


if __name__ == '__main__':
    tm = TableManager()
    tm.test('../test_data/test_dit.pptx')
