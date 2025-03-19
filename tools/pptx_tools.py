"""
PowerPoint Automation Tool using python-pptx
"""
from __future__ import annotations

import contextlib
import tempfile
import uuid

from pathlib import Path
from typing import Any

from pptx import Presentation
from pptx.chart.data import ChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.slide import Slide as pptx_Slide
from pptx.slide import SlideLayout as pptx_SlideLayout
from pptx.table import _Cell
from pptx.util import Inches, Pt
from pptx.dml.fill import FillFormat
from pptx.shapes.base import BaseShape
from .schemas import (
    Background,
    ChartShape,
    Coordinates,
    FillType,
    FontStyle,
    GeometricShape,
    ImageShape,
    PresentationInfo,
    PresentationMeta,
    Slide,
    SlideDesign,
    SlideLayout,
    TableCell,
    TableShape,
    TextContent,
    TextShape,
)
from .schemas import ChartData as SchemaChartData
from pptx.enum.dml import MSO_FILL_TYPE, MSO_FILL, MSO_COLOR_TYPE

def safe_rgb_to_hex(rgb_color_obj) -> str | None:
    """
    Safely extract a #RRGGBB string from older or newer python-pptx RGBColor objects.
    Returns None if we can't parse a color.
    """
    if not rgb_color_obj:
        return None

    # Some older python-pptx versions store the actual integer in ._rgb
    raw = getattr(rgb_color_obj, "_rgb", None)  # might be int or None
    if isinstance(raw, int):
        # raw is a 24-bit integer like 0xRRGGBB
        r = (raw >> 16) & 0xFF
        g = (raw >> 8) & 0xFF
        b = raw & 0xFF
        return f"#{r:02X}{g:02X}{b:02X}"

    # Some older versions might let you do just int(rgb_color_obj)
    if isinstance(rgb_color_obj, int):
        r = (rgb_color_obj >> 16) & 0xFF
        g = (rgb_color_obj >> 8) & 0xFF
        b = rgb_color_obj & 0xFF
        return f"#{r:02X}{g:02X}{b:02X}"

    # If all else fails, convert to string and see if it's already #AABBCC
    # or do a fallback
    as_str = str(rgb_color_obj).strip()
    if as_str.startswith("#") and len(as_str) == 7:
        return as_str  # e.g. "#RRGGBB"
    # else return None
    return None

class PPTXManager:
    """
    Manages the creation and manipulation of a PowerPoint (.pptx) presentation.

    Attributes:
        prs (Presentation): The underlying python-pptx Presentation object.
        _temp_files (list[str]): Internal list of paths to any temporary files created.

    """

    def __init__(self, file_path: str | None = None) -> None:
        """
        Initialize the PPTXManager.

        If a file path is provided, opens the existing presentation; otherwise,
        creates a new blank presentation.

        Args:
            file_path (str | None): Path to an existing .pptx file or None to create a new presentation.
        """
        self.prs = Presentation(file_path) if file_path else Presentation()
        self._temp_files = []  # For image handling

    def save(self, path: str) -> None:
        """
        Save the current presentation to a file.

        This method writes the current in-memory presentation state to the specified path.

        Args:
            path (str): The file path where the .pptx should be saved.

        Returns:
            None
        """
        self.prs.save(path)
        self._cleanup_temp_files()

    def close(self) -> None:
        """
        Close the manager and clean up any resources, including temporary files.

        Returns:
            None
        """
        self._cleanup_temp_files()

    def _cleanup_temp_files(self) -> None:
        """
        Internal method to remove any temporary files that were created for images, etc.

        Returns:
            None
        """
        for f in self._temp_files:
            with contextlib.suppress(Exception):
                Path(f).unlink()
        self._temp_files = []

    def get_presentation_info(self) -> PresentationInfo:
        """
        Retrieve metadata and structural information about the presentation.

        This includes slide count, slide dimensions, layout names, and
        parsed details of each slide.

        Returns:
            PresentationInfo: A dataclass containing presentation metadata, layout info, and slide details.
        """
        return PresentationInfo(
            meta=self._get_metadata(),
            dimensions=self._get_slide_dimensions(),
            slides_count=len(self.prs.slides),
            slides=[self._parse_slide(i+1, slide) for i, slide in enumerate(self.prs.slides)],
            master_layouts=self._get_master_layouts()
        )


    def get_slide_details(self, slide_num: int) -> Slide:
        """
        Retrieve detailed information about a specific slide, identified by its slide number (1-based).

        Args:
            slide_num (int): The 1-based index of the slide to inspect.

        Returns:
            Slide: A detailed dataclass representation of the slide, including layout, background, and shapes.

        Raises:
            ValueError: If slide_num is out of range.
        """
        slide = self._get_slide(slide_num)
        return self._parse_slide(slide_num, slide)

    def analyze_slide_design(self, slide_num: int) -> SlideDesign:
        """
        Analyze the design of a specific slide, extracting theme colors, fonts, and default margins.

        Args:
            slide_num (int): The 1-based slide index.

        Returns:
            SlideDesign: A dataclass with theme colors, fonts, and default margin info.

        Raises:
            ValueError: If slide_num is out of range.
        """
        slide = self._get_slide(slide_num)
        return SlideDesign(
            theme_colors=self._get_theme_colors(slide),
            theme_fonts=self._get_theme_fonts(slide),
            default_margin=0.5
        )

    def create_new_slide(self, layout_name: str, position: int | None = None) -> int:
        """
        Create a new slide using an existing layout.

        Args:
            layout_name (str): The name of the layout to use (must be present in the presentation).
            position (int | None): If provided, insert the new slide at this 1-based position; otherwise, it is appended.

        Returns:
            int: The total number of slides after adding the new slide.

        Raises:
            ValueError: If the layout_name is not found.
        """
        layouts = self._get_master_layouts()
        layout = next((layout_item for layout_item in layouts if layout_item.name == layout_name), None)

        if not layout:
            msg = f"Layout '{layout_name}' not found"
            raise ValueError(msg)

        slide_layout = self.prs.slide_layouts[layout.idx]
        new_slide = self.prs.slides.add_slide(slide_layout)

        if position is not None:
            sldIdLst = self.prs.slides._sldIdLst
            slides = list(sldIdLst)
            new_id = slides[-1]  # newly inserted
            slides.remove(new_id)
            slides.insert(position - 1, new_id)

        return len(self.prs.slides)

    def duplicate_slide(self, slide_num: int) -> int:
        """
        Duplicate an existing slide by cloning its layout and shapes.

        Args:
            slide_num (int): The 1-based index of the slide to duplicate.

        Returns:
            int: The total number of slides after duplication.

        Raises:
            ValueError: If slide_num is out of range.
        """
        source = self._get_slide(slide_num)
        self.prs.slides.add_slide(source.slide_layout)
        return len(self.prs.slides)

    def delete_slide(self, slide_num: int) -> None:
        """
        Remove the specified slide from the presentation.

        Args:
            slide_num (int): The 1-based index of the slide to remove.

        Returns:
            None

        Raises:
            ValueError: If slide_num is out of range.
        """
        if slide_num < 1 or slide_num > len(self.prs.slides):
            msg = f"Invalid slide number: {slide_num}"
            raise ValueError(msg)

        sldIdLst = self.prs.slides._sldIdLst
        slides = list(sldIdLst)
        sldIdLst.remove(slides[slide_num - 1])

    def set_slide_background(self, slide_num: int, background: Background | dict) -> None:
        """
        Set the background for a specific slide.

        Args:
            slide_num (int): The 1-based index of the slide to modify.
            background (Background): A dataclass describing the background type, value, and transparency. {"type": "color", "value": "#FFC0CB", "transparency": 1}

        Returns:
            None

        Raises:
            ValueError: If slide_num is out of range.
        """
        if isinstance(background, dict):
            background = Background(**background)

        slide = self._get_slide(slide_num)
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor.from_string(background.value.lstrip('#') if background.value else "FFFFFF")
        slide.background.fill.transparency = background.transparency

    def apply_slide_template(self, slide_num: int, template_path: str) -> None:
        """
        Apply a slide template from an external PowerPoint file to the specified slide.

        Args:
            slide_num (int): The 1-based index of the slide to update.
            template_path (str): The path to the template .pptx file.

        Returns:
            None

        Raises:
            ValueError: If slide_num is out of range or template cannot be applied.
        """
        template_prs = Presentation(template_path)
        template_layout = template_prs.slide_layouts[0]

        self.delete_slide(slide_num)
        new_slide = self.prs.slides.add_slide(template_layout)
        sldIdLst = self.prs.slides._sldIdLst
        slides = list(sldIdLst)
        new_id = slides[-1]  # newly inserted
        slides.remove(new_id)
        slides.insert(slide_num - 1, new_id)
        sldIdLst.clear()
        for sId in slides:
            sldIdLst.append(sId)

    def _set_text_content(self, text_frame, content: str, style: FontStyle) -> None:
        """
        Internal helper to set the text content and apply a font style to a shape's text_frame.

        Args:
            text_frame: The text frame object from a shape.
            content (str): The text content to insert.
            style (FontStyle): A dataclass specifying font attributes (size, color, bold, etc.).

        Returns:
            None
        """
        text_frame.text = content
        self._apply_font_style(text_frame, style)

    def add_text_block(
        self,
        slide_num: int,
        content: str,
        pos: Coordinates | dict,
        style: FontStyle | dict
    ) -> str:
        """
        Add a text block (textbox) to a slide at specified coordinates.

        Args:
            slide_num (int): The 1-based index of the slide to which the text block is added.
            content (str): The text to display in the text block.
            pos (Coordinates): The position and size of the textbox (x, y, width, height in inches).
            style (FontStyle): A dataclass defining font style properties.

        Returns:
            str: A unique shape identifier (UUID) assigned to the new shape.

        Raises:
            ValueError: If slide_num is out of range.
        """
        if isinstance(pos, dict):
            pos = Coordinates(**pos)

        if isinstance(style, dict):
            style = FontStyle(**style)

        slide = self._get_slide(slide_num)
        textbox = slide.shapes.add_textbox(
            Inches(pos.x), Inches(pos.y),
            Inches(pos.width), Inches(pos.height)
        )
        self._set_text_content(textbox.text_frame, content, style)
        return self._set_shape_id(textbox)

    def edit_text_content(self, slide_num: int, shape_id: str, new_text: str) -> None:
        """
        Modify the text content of an existing text box shape.

        Args:
            slide_num (int): The 1-based index of the slide containing the shape.
            shape_id (str): The unique identifier of the text box shape.
            new_text (str): The new text to set.

        Returns:
            None

        Raises:
            ValueError: If slide_num is out of range or shape is not found or is not a text box.
        """
        shape = self._find_shape(slide_num, shape_id)
        if shape.shape_type != MSO_SHAPE_TYPE.TEXT_BOX:
            msg = "Shape is not a text box"
            raise ValueError(msg)
        shape.text_frame.text = new_text

    def format_text_style(self, slide_num: int, shape_id: str, style: FontStyle | dict) -> None:
        """
        Apply or update the font style of a text shape.

        Args:
            slide_num (int): The 1-based index of the slide containing the shape.
            shape_id (str): The unique identifier of the text shape.
            style (FontStyle): The new font style (size, color, bold, etc.) to apply.

        Returns:
            None

        Raises:
            ValueError: If slide_num or shape_id is invalid.
        """
        if isinstance(style, dict):
            style = FontStyle(**style)
        shape = self._find_shape(slide_num, shape_id)
        self._apply_font_style(shape.text_frame, style)

    def insert_image(self, slide_num: int, image_data: str | bytes, pos: Coordinates | dict) -> str:
        """
        Insert an image onto a slide at the specified coordinates.

        Args:
            slide_num (int): The 1-based index of the slide.
            image_data (str | bytes): File path to the image, or raw bytes of the image.
            pos (Coordinates): The position and size in inches.

        Returns:
            str: A unique shape identifier (UUID) assigned to the new image shape.

        Raises:
            ValueError: If slide_num is out of range.
        """
        if isinstance(pos, dict):
            pos = Coordinates(**pos)

        slide = self._get_slide(slide_num)

        img_path = self._bytes_to_tempfile(image_data) if isinstance(image_data, bytes) else image_data

        image = slide.shapes.add_picture(
            img_path,
            Inches(pos.x), Inches(pos.y),
            Inches(pos.width), Inches(pos.height)
        )
        return self._set_shape_id(image)

    def replace_image(self, slide_num: int, shape_id: str, new_image: str | bytes) -> None:
        """
        Replace an existing image shape with a new image, preserving position and size.

        Args:
            slide_num (int): The 1-based index of the slide.
            shape_id (str): The unique identifier of the existing image shape.
            new_image (str | bytes): The new image file path or raw bytes.

        Returns:
            None

        Raises:
            ValueError: If slide_num is out of range or the specified shape is not found.
        """
        old_shape = self._find_shape(slide_num, shape_id)
        pos = Coordinates(
            x=old_shape.left.inches,
            y=old_shape.top.inches,
            width=old_shape.width.inches,
            height=old_shape.height.inches
        )
        self.delete_shape(slide_num, shape_id)
        self.insert_image(slide_num, new_image, pos)

    def create_chart(
        self,
        slide_num: int,
        chart_type: str,
        data: SchemaChartData | dict,
        pos: Coordinates | dict
    ) -> str:
        """
        Create a new chart on a specified slide.

        Args:
            slide_num (int): The 1-based index of the slide.
            chart_type (str): The type of chart (e.g., 'bar', 'line', 'pie', 'area').
            data (SchemaChartData): Chart data containing categories and series.
            pos (Coordinates): Position and size in inches.

        Returns:
            str: A unique shape identifier (UUID) assigned to the chart.

        Raises:
            ValueError: If slide_num is out of range or chart_type is invalid.
        """
        if isinstance(pos, dict):
            pos = Coordinates(**pos)

        if isinstance(data, dict):
            data = SchemaChartData(**data)

        slide = self._get_slide(slide_num)
        chart_data = ChartData()
        chart_data.categories = data.categories

        for series in data.series:
            chart_data.add_series(series.name, series.values)

        chart = slide.shapes.add_chart(
            self._map_chart_type(chart_type),
            Inches(pos.x), Inches(pos.y),
            Inches(pos.width), Inches(pos.height),
            chart_data
        )
        return self._set_shape_id(chart)

    def modify_chart_data(self, slide_num: int, shape_id: str, new_data: SchemaChartData | dict) -> None:
        """
        Modify data in an existing chart.

        Args:
            slide_num (int): The 1-based index of the slide containing the chart.
            shape_id (str): The unique identifier of the chart shape.
            new_data (SchemaChartData): The new chart data to replace current data.

        Returns:
            None

        Raises:
            ValueError: If slide_num is out of range or shape_id is invalid/not a chart.
        """
        if isinstance(new_data, dict):
            new_data = SchemaChartData(**new_data)
        chart = self._find_shape(slide_num, shape_id)
        chart_data = ChartData()
        chart_data.categories = new_data.categories
        for series in new_data.series:
            chart_data.add_series(series.name, series.values)
        chart.chart.replace_data(chart_data)

    def create_table(self, slide_num: int, rows: int, cols: int, pos: Coordinates | dict) -> str:
        """
        Create a table on a specified slide.

        Args:
            slide_num (int): The 1-based index of the slide.
            rows (int): Number of rows in the new table.
            cols (int): Number of columns in the new table.
            pos (Coordinates): Position and size in inches.

        Returns:
            str: A unique shape identifier (UUID) for the new table shape.

        Raises:
            ValueError: If slide_num is out of range.
        """
        if isinstance(pos, dict):
            pos = Coordinates(**pos)
        slide = self._get_slide(slide_num)
        table_shape = slide.shapes.add_table(
            rows, cols,
            Inches(pos.x), Inches(pos.y),
            Inches(pos.width), Inches(pos.height)
        )
        return self._set_shape_id(table_shape)

    def edit_table_cell(self, slide_num: int, shape_id: str, cell: TableCell | dict) -> None:
        """
        Modify the content or merging properties of a single cell in a table.

        Args:
            slide_num (int): The 1-based index of the slide containing the table.
            shape_id (str): The unique identifier of the table shape.
            cell (TableCell): A dataclass describing row, col, content, and span info.

        Returns:
            None

        Raises:
            ValueError: If slide_num or shape_id is invalid, or shape is not a table.
        """
        if isinstance(cell, dict):
            cell = TableCell(**cell)
        graphic_frame = self._find_shape(slide_num, shape_id)
        if not hasattr(graphic_frame, 'table'):
            msg = "Shape is not a table"
            raise ValueError(msg)

        table = graphic_frame.table
        cell_obj = table.cell(cell.row, cell.col)
        cell_obj.text = cell.content
        if cell.span_rows > 1 or cell.span_cols > 1:
            cell_obj.merge(table.cell(
                cell.row + cell.span_rows - 1,
                cell.col + cell.span_cols - 1
            ))

    def _get_slide(self, slide_num: int) -> pptx_Slide:
        """
        Internal method to retrieve a slide object by 1-based index, validating range.

        Args:
            slide_num (int): The 1-based slide index.

        Returns:
            pptx.slide.Slide: The requested slide object.

        Raises:
            ValueError: If the index is out of range.
        """
        if slide_num < 1 or slide_num > len(self.prs.slides):
            msg = f"Invalid slide number: {slide_num}"
            raise ValueError(msg)
        return self.prs.slides[slide_num-1]

    def _find_shape(self, slide_num: int, shape_id: str) -> BaseShape:
        """
        Internal method to find a shape by UUID-like name on a given slide.

        Args:
            slide_num (int): The 1-based slide index.
            shape_id (str): The string name (UUID) assigned to the shape.

        Returns:
            BaseShape: The shape object if found.

        Raises:
            ValueError: If the shape cannot be found on the specified slide.
        """
        slide = self._get_slide(slide_num)
        for shape in slide.shapes:
            if shape.name == shape_id:
                return shape
        msg = f"Shape {shape_id} not found on slide {slide_num}"
        raise ValueError(msg)

    def _set_shape_id(self, shape) -> str:
        """
        Internal method to assign a UUID name to a shape.

        Args:
            shape: The shape object.

        Returns:
            str: The generated UUID name.
        """
        shape_id = str(uuid.uuid4())
        shape.name = shape_id
        return shape_id

    def _apply_font_style(self, text_frame, style: FontStyle) -> None:
        """
        Internal helper to apply font style to all runs in a text frame.

        Args:
            text_frame: The text frame (containing paragraphs and runs).
            style (FontStyle): The font style to apply.

        Returns:
            None
        """
        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                font = run.font
                font.name = style.name
                font.size = Pt(style.size)
                font.color.rgb = RGBColor.from_string(style.color.lstrip('#'))
                font.bold = style.bold
                font.italic = style.italic
                font.underline = style.underline

    def _bytes_to_tempfile(self, data: bytes) -> str:
        """
        Internal helper to persist raw image bytes to a temporary file.

        Args:
            data (bytes): The image data in bytes.

        Returns:
            str: The path to the temporary file.
        """
        with tempfile.NamedTemporaryFile(delete=False) as f:
            f.write(data)
            self._temp_files.append(f.name)
            return f.name

    def _map_chart_type(self, chart_type: str) -> XL_CHART_TYPE:
        """
        Internal helper to map a string chart type to the XL_CHART_TYPE enum.

        Args:
            chart_type (str): The string specifying the chart type (e.g., 'bar', 'line').

        Returns:
            XL_CHART_TYPE: The corresponding python-pptx chart type enum value.
        """
        types = {
            'bar': XL_CHART_TYPE.COLUMN_CLUSTERED,
            'line': XL_CHART_TYPE.LINE,
            'pie': XL_CHART_TYPE.PIE,
            'area': XL_CHART_TYPE.AREA
        }
        return types.get(chart_type.lower(), XL_CHART_TYPE.COLUMN_CLUSTERED)

    def _parse_slide(self, slide_num: int, slide: pptx_Slide) -> Slide:
        """
        Internal method that converts a pptx Slide object into a Slide dataclass.

        Args:
            slide_num (int): 1-based index of the slide.
            slide (pptx.slide.Slide): The pptx Slide object.

        Returns:
            Slide: A dataclass describing the slide.
        """
        parsed_shapes = []
        for shp in slide.shapes:
            model = self._parse_shape(shp)
            if model is not None:
                parsed_shapes.append(model)

        return Slide(
            number=slide_num,
            layout=self._parse_layout(slide.slide_layout),
            background=self._parse_background(slide.background),
            shapes=parsed_shapes,  # Only real shapes
            design=self.analyze_slide_design(slide_num),
        )

    def _get_shape_parser(self, shape_type) -> callable | None:
        """
        Retrieve the appropriate parser function based on a shape's type.

        Args:
            shape_type (MSO_SHAPE_TYPE): The type enum of the shape.

        Returns:
            callable | None: A function that parses the shape into a corresponding dataclass, or None.
        """
        parsers = {
            MSO_SHAPE_TYPE.TEXT_BOX: self._parse_text_shape,
            MSO_SHAPE_TYPE.PICTURE: self._parse_image_shape,
            MSO_SHAPE_TYPE.CHART: self._parse_chart_shape,
            MSO_SHAPE_TYPE.TABLE: self._parse_table_shape,
            MSO_SHAPE_TYPE.AUTO_SHAPE: self._parse_geometric_shape
        }
        return parsers.get(shape_type)

    def _parse_shape(self, shape: BaseShape) -> TextShape | ImageShape | ChartShape | TableShape | GeometricShape | None:
        """
        Parse a shape into one of the known shape dataclasses (TextShape, ImageShape, etc.).

        Args:
            shape: The pptx shape object.

        Returns:
            TextShape | ImageShape | ChartShape | TableShape | GeometricShape | None:
            The parsed shape if recognized, otherwise None.
        """
        common = {
            'id': shape.name,
            'coordinates': Coordinates(
                x=shape.left.inches,
                y=shape.top.inches,
                width=shape.width.inches,
                height=shape.height.inches
            )
        }
        parser = self._get_shape_parser(shape.shape_type)
        if parser:
            try:
                return parser(shape, **common)
            except Exception as e:
                print(f"Error parsing shape: {e}")
                return None
        else:
            # shape type not recognized
            return None

    def _parse_text_shape(self, shape, **common: dict[str, Any]) -> TextShape:
        """
        Parse a text box shape into a TextShape dataclass.

        Args:
            shape: The text box shape.
            **common: Common shape attributes (id, coordinates).

        Returns:
            TextShape: The parsed text shape including content and style.
        """
        return TextShape(
            **common,
            content=TextContent(text=shape.text),
            style=self._parse_font_style(shape.text_frame)
        )

    def _parse_image_shape(self, shape, **common: dict[str, Any]) -> ImageShape | None:
        """
        Parse an image shape into an ImageShape dataclass.

        Args:
            shape: The image shape.
            **common: Common shape attributes (id, coordinates).

        Returns:
            ImageShape | None: The parsed image shape, or None if there's an error.
        """
        try:
            return ImageShape(
                **common,
                path=shape.image.filename,
                crop=self._parse_image_crop(shape),
                brightness=getattr(shape, 'brightness', 1.0),
                contrast=getattr(shape, 'contrast', 1.0)
            )
        except Exception as e:
            print(f"Error parsing image: {str(e)}")
            return None

    def _parse_chart_shape(self, shape: BaseShape, **common: dict[str, Any]) -> ChartShape:
        """
        Parse a chart shape into a ChartShape dataclass.

        Args:
            shape: The chart shape.
            **common: Common shape attributes (id, coordinates).

        Returns:
            ChartShape: The parsed chart shape with chart type, data, and title.
        """
        chart = shape.chart
        # --- 1) Chart Type ---
        raw_chart_type = str(chart.chart_type)
        chart_type_str = raw_chart_type.split('.')[-1]
        chart_type_str = chart_type_str.split(' ')[0]

        type_map = {
            "BAR": "bar",
            "BAR_STACKED": "bar",
            "BAR_CLUSTERED": "bar",
            "COLUMN": "bar",
            "COLUMN_CLUSTERED": "bar",
            "COLUMN_STACKED": "bar",
            "LINE": "line",
            "LINE_MARKERS": "line",
            "LINE_STACKED": "line",
            "PIE": "pie",
            "PIE_EXPLODED": "pie",
            "AREA": "area",
            "AREA_STACKED": "area"
        }
        chart_type = type_map.get(chart_type_str, "bar")

        # --- 2) Categories ---
        categories = []
        try:
            first_plot = chart.plots[0]
            if hasattr(first_plot, "categories"):
                categories = list(first_plot.categories)
        except (AttributeError, IndexError) as e:
            print(f"Error parsing chart categories: {e}")
            categories = []

        # --- 3) Series Data ---
        series_data = []
        try:
            for s in chart.series:
                safe_values = []
                for val in s.values:
                    if val is None:
                        safe_values.append(0.0)
                    else:
                        safe_values.append(float(val))
                series_data.append({"name": s.name or "Series", "values": safe_values})
        except AttributeError as e:
            print(f"Error parsing chart series: {e}")
            return None

        # --- 4) Build ChartShape ---
        title = chart.chart_title.text if chart.has_title else None
        return ChartShape(
            **common,
            chart_type=chart_type,
            data=SchemaChartData(
                categories=categories,
                series=series_data
            ),
            title=title
        )

    def _parse_table_shape(self, shape, **common: dict[str, Any]) -> TableShape:
        """
        Parse a table shape into a TableShape dataclass.

        Args:
            shape: The table shape.
            **common: Common shape attributes (id, coordinates).

        Returns:
            TableShape: The parsed table shape, including row/column count and cell data.
        """
        table = shape.table
        rows_count = len(table.rows)
        cols_count = len(table.columns)
        cells_data = []

        for row_i in range(rows_count):
            for col_i in range(cols_count):
                cell_obj = table.cell(row_i, col_i)
                content = cell_obj.text
                cells_data.append(TableCell(
                    row=row_i,
                    col=col_i,
                    content=content,
                    span_rows=1,
                    span_cols=1,
                ))

        return TableShape(
            **common,
            rows=rows_count,
            cols=cols_count,
            cells=cells_data
        )

    def _parse_geometric_shape(self, shape: BaseShape, **common: dict[str, Any]) -> GeometricShape:
        """
        Parse an auto shape (geometric) into a GeometricShape dataclass.

        Args:
            shape: The geometric shape.
            **common: Common shape attributes (id, coordinates).

        Returns:
            GeometricShape: The parsed geometric shape with fill and outline properties.
        """

        shape_type_name = getattr(shape.auto_shape_type, "name", "rectangle").lower()

        # We won't bail out even if shape_type_name is "rounded_rectangle" – 
        # because our updated GeometricShape accepts any string.

        # 1) FILL COLOR
        fill_rgb = None
        try:
            # If fill is NO_FILL, we skip. 
            # 'NO_FILL' might be `_NoneFill` or `_NoFill` in older versions.
            if shape.fill and shape.fill.type in (MSO_FILL_TYPE.SOLID, MSO_FILL_TYPE.PATTERNED, MSO_FILL_TYPE.GRADIENT):
                raw_color_obj = shape.fill.fore_color.rgb  # older python-pptx might store an object
                fill_rgb = safe_rgb_to_hex(raw_color_obj)
        except Exception as e:
            print(f"Error parsing fill color: {e}")

        # 2) OUTLINE COLOR
        outline_rgb = None
        try:
            # If there's a line color object with .rgb
            if shape.line and shape.line.color and shape.line.color.type not in [MSO_COLOR_TYPE.SCHEME, None]:
                outline_rgb = safe_rgb_to_hex(shape.line.color.rgb)
        except Exception as e:
            print(f"Error parsing outline color: {e}")

        # 3) Outline width
        outline_width_in = 0.0
        try:
            if shape.line and shape.line.width:
                outline_width_in = shape.line.width.inches
        except Exception as e:
            print(f"Error parsing outline width: {e}")

        return GeometricShape(
            **common,
            shape_type=shape_type_name,
            fill=fill_rgb,
            outline=outline_rgb,
            outline_width=outline_width_in
        )

    def _parse_image_crop(self, shape) -> Coordinates | None:
        """
        Parse cropping parameters of an image shape.

        Args:
            shape: The image shape with potential crop attributes.

        Returns:
            Coordinates | None: A Coordinates instance describing the crop area (0.0 - 1.0),
                                or None if no cropping is applied.
        """
        if shape.crop_left or shape.crop_right or shape.crop_top or shape.crop_bottom:
            return Coordinates(
                x=shape.crop_left,
                y=shape.crop_top,
                width=1 - (shape.crop_left + shape.crop_right),
                height=1 - (shape.crop_top + shape.crop_bottom)
            )
        return None

    def _parse_table_cell(self, cell: _Cell) -> TableCell:
        """
        Parse a pptx _Cell object into a TableCell dataclass.

        Args:
            cell (_Cell): The table cell object.

        Returns:
            TableCell: A dataclass with row, col, content, and span information.
        """
        return TableCell(
            row=cell.row_idx,
            col=cell.col_idx,
            content=cell.text,
            span_rows=cell.span_height,
            span_cols=cell.span_width
        )

    def _parse_font_style(self, text_frame) -> FontStyle:
        """
        Parse font styling from the first run in a text frame.

        Args:
            text_frame: The text frame containing paragraphs and runs.

        Returns:
            FontStyle: Parsed font style (font name, size, color, bold, italic, etc.).
                       Defaults are applied if attributes are missing or can't be accessed.
        """
        try:
            first_run = text_frame.paragraphs[0].runs[0]
            font = first_run.font

            # Default color if we cannot parse
            color_str = "#000000"

            # Safely convert font.color.rgb to hex, if present
            if font.color and font.color.rgb:  
                r, g, b = font.color.rgb.red, font.color.rgb.green, font.color.rgb.blue
                color_str = f"#{r:02X}{g:02X}{b:02X}"
            elif font.color and hasattr(font.color, "theme_color") and font.color.theme_color:
                # If theme color is used, store it as "theme:XYZ"
                color_str = f"theme:{font.color.theme_color}"

            return FontStyle(
                name=font.name or "Calibri",
                size=font.size.pt if font.size else 12.0,
                color=color_str,
                bold=bool(font.bold),
                italic=bool(font.italic),
                underline=bool(font.underline)
            )
        except Exception:
            # If anything goes wrong, return a default style
            return FontStyle()

    def _get_metadata(self) -> PresentationMeta:
        """
        Internal method to extract core properties (metadata) of the presentation.

        Returns:
            PresentationMeta: A dataclass with title, author, created date, etc.
        """
        cp = self.prs.core_properties
        return PresentationMeta(
            title=cp.title,
            author=cp.author,
            created=cp.created,
            modified=cp.modified,
            template=self.prs.slide_layouts[0].name if self.prs.slide_layouts else None
        )

    def _get_slide_dimensions(self) -> Coordinates:
        """
        Internal method to get the width and height of the presentation slides in inches.

        Returns:
            Coordinates: The width and height of the slides.
        """
        return Coordinates(
            width=self.prs.slide_width.inches,
            height=self.prs.slide_height.inches
        )

    def _get_master_layouts(self) -> list[SlideLayout]:
        """
        Internal method to list all the master slide layouts in the presentation.

        Returns:
            list[SlideLayout]: A list of SlideLayout objects (name, display_name, idx).
        """
        return [
            SlideLayout(
                name=layout.name,
                display_name=layout.name,
                idx=idx
            ) for idx, layout in enumerate(self.prs.slide_layouts)
        ]

    def _get_theme_fonts(self, slide: pptx_Slide) -> dict[str, str]:
        """
        Extract theme fonts from the slide's layout.

        Args:
            slide (pptx.slide.Slide): The slide to analyze.

        Returns:
            dict[str, str]: A dictionary of theme font mappings (latin, complex, east_asian).
                            May be empty or partial if not defined.
        """
        if hasattr(slide.slide_layout, "fonts"):
            fonts = slide.slide_layout.fonts
            return {
                "latin": getattr(fonts.latin, "typeface", ""),
                "complex": getattr(fonts.complex, "typeface", ""),
                "east_asian": getattr(fonts.east_asian, "typeface", ""),
            }
        return {}

    def _find_layout(self, layout_name: str) -> SlideLayout:
        """
        Find a slide layout by its name.

        Args:
            layout_name (str): The layout name or display name to match.

        Returns:
            SlideLayout: The matching layout object.

        Raises:
            ValueError: If no matching layout is found.
        """
        for layout in self._get_master_layouts():
            if layout_name in (layout.name, layout.display_name):
                return layout
        msg = f"Layout '{layout_name}' not found"
        raise ValueError(msg)

    def _reorder_slide(self, slide: pptx_Slide, position: int) -> None:
        """
        Internal method to reorder a specific slide to a new position within the presentation.

        Args:
            slide (pptx.slide.Slide): The slide object to reorder.
            position (int): The new 1-based position for the slide.

        Returns:
            None
        """
        temp_prs = Presentation()

        for i in range(position - 1):
            if i < len(self.prs.slides):
                source = self.prs.slides[i]
                dest = temp_prs.slides.add_slide(source.slide_layout)
                dest.shapes.clone(source.shapes)

        dest = temp_prs.slides.add_slide(slide.slide_layout)
        dest.shapes.clone(slide.shapes)

        for i in range(position - 1, len(self.prs.slides) - 1):
            source = self.prs.slides[i]
            dest = temp_prs.slides.add_slide(source.slide_layout)
            dest.shapes.clone(source.shapes)
        self.prs = temp_prs

    def _parse_layout(self, layout: pptx_SlideLayout) -> SlideLayout:
        """
        Internal method to parse a pptx SlideLayout into a SlideLayout dataclass.

        Args:
            layout (pptx.slide.SlideLayout): The layout object.

        Returns:
            SlideLayout: Parsed layout info (name, display_name, idx).
        """
        try:
            idx = self.prs.slide_layouts.index(layout)
        except ValueError:
            idx = -1
        return SlideLayout(
            name=layout.name,
            display_name=layout.name,
            idx=idx
        )

    def _parse_background(self, background: Background) -> Background:
        """
        Internal method to parse the background fill of a slide.

        Args:
            background (Background): The background object.

        Returns:
            Background: A dataclass with type, value, and transparency.
        """
        fill = background.fill
        bg_type = "none"
        value = None
        transparency = 1.0

        try:
            if fill.type == MSO_FILL.SOLID and fill.fore_color.rgb:
                bg_type = "color"
                r, g, b = fill.fore_color.rgb.red, fill.fore_color.rgb.green, fill.fore_color.rgb.blue
                value = f"#{r:02X}{g:02X}{b:02X}"
                transparency = fill.transparency or 0.0

            elif fill.type == MSO_FILL.PICTURE and hasattr(fill, "picture"):
                bg_type = "image"
                value = getattr(fill.picture, "url", "embedded")
                transparency = fill.transparency or 0.0

            elif fill.type == MSO_FILL.PATTERNED:
                bg_type = "pattern"
                value = fill.pattern
                transparency = fill.transparency or 0.0

            else:
                bg_type = "none"
                value = None
                transparency = 0.0
        except Exception:
            # If anything unexpected happens, default to no fill
            bg_type = "none"
            value = None
            transparency = 0.0

        return Background(
            type=bg_type,
            value=value,
            transparency=transparency
        )

    def delete_shape(self, slide_num: int, shape_id: str) -> None:
        """
        Delete a shape from a slide by its unique identifier.

        Args:
            slide_num (int): The 1-based index of the slide.
            shape_id (str): The UUID identifying the shape.

        Returns:
            None

        Raises:
            ValueError: If slide_num is out of range or shape not found.
        """
        shape = self._find_shape(slide_num, shape_id)
        sp = shape.element
        sp.getparent().remove(sp)

    def _get_theme_colors(self, slide) -> list[str]:
        """
        Extract theme colors from the slide layout, if available.

        Args:
            slide (pptx.slide.Slide): The slide to analyze.

        Returns:
            list[str]: A list of theme color names or hex values, if found, otherwise empty list.
        """
        try:
            return [str(c) for c in slide.slide_layout.color_scheme.colors]
        except AttributeError:
            return []