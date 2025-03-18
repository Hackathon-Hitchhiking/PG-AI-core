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


class PPTXManager:
    """Main PowerPoint manipulation class"""

    def __init__(self, file_path: str | None = None) -> None:
        self.prs = Presentation(file_path) if file_path else Presentation()
        self._temp_files = []  # For image handling

    #region Core Methods
    def save(self, path: str) -> None:
        """Save presentation to file"""
        self.prs.save(path)
        self._cleanup_temp_files()

    def close(self) -> None:
        """Clean up resources"""
        self._cleanup_temp_files()

    def _cleanup_temp_files(self) -> None:
        """Remove temporary files"""
        for f in self._temp_files:
            with contextlib.suppress(Exception):
                Path(f).unlink()
        self._temp_files = []
    #endregion

    #region Analysis Methods
    def get_presentation_info(self) -> PresentationInfo:
        """Get complete presentation metadata"""
        return PresentationInfo(
            meta=self._get_metadata(),
            dimensions=self._get_slide_dimensions(),
            slides_count=len(self.prs.slides),
            slides=[self._parse_slide(i, slide) for i, slide in enumerate(self.prs.slides)],
            master_layouts=self._get_master_layouts()
        )

    def get_slide_details(self, slide_num: int) -> Slide:
        """Get detailed information about a specific slide"""
        slide = self._get_slide(slide_num)
        return self._parse_slide(slide_num, slide)

    def analyze_slide_design(self, slide_num: int) -> SlideDesign:
        """Анализ дизайна с проверкой атрибутов"""
        slide = self._get_slide(slide_num)
        return SlideDesign(
            theme_colors=self._get_theme_colors(slide),
            theme_fonts=self._get_theme_fonts(slide),
            default_margin=0.5
        )
    #endregion

    #region Slide Management
    def create_new_slide(self, layout_name: str, position: int | None = None) -> int:
        layouts = self._get_master_layouts()
        layout = next((layout_item for layout_item in layouts if layout_item.name == layout_name), None)

        if not layout:
            msg = f"Layout '{layout_name}' not found"
            raise ValueError(msg)

        slide_layout = self.prs.slide_layouts[layout.idx]
        new_slide = self.prs.slides.add_slide(slide_layout)

        if position is not None:
            self._reorder_slide(new_slide, position)

        return len(self.prs.slides)

    def duplicate_slide(self, slide_num: int) -> int:
        """Безопасное клонирование слайдов"""
        source = self._get_slide(slide_num)
        dest = self.prs.slides.add_slide(source.slide_layout)

        try:
            dest.shapes.clone(source.shapes)
        except Exception as e:
            print(f"Error cloning shapes: {str(e)}")

        return len(self.prs.slides)


    def delete_slide(self, slide_num: int) -> None:
        """Remove specified slide"""
        temp_prs = Presentation()
        for idx, slide in enumerate(self.prs.slides, start=1):
            if idx != slide_num:
                new_slide = temp_prs.slides.add_slide(slide.slide_layout)
                new_slide.shapes.clone(slide.shapes)
        self.prs = temp_prs

    def reorder_slides(self, new_order: list[int]) -> None:
        """Reorder slides according to provided indices"""
        # Create a new presentation and copy slides in desired order
        temp_prs = Presentation()
        for idx in new_order:
            source = self.prs.slides[idx-1]
            dest = temp_prs.slides.add_slide(source.slide_layout)
            dest.shapes.clone(source.shapes)
        # Replace original slides with reordered ones
        self.prs = temp_prs
    #endregion

    #region Background & Design
    def set_slide_background(self, slide_num: int, background: Background) -> None:
        """Set slide background"""
        slide = self._get_slide(slide_num)
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor.from_string(background.value.lstrip('#'))
        fill.transparency = background.transparency

    def apply_slide_template(self, slide_num: int, template_path: str) -> None:
        """Apply template from another presentation"""
        template_prs = Presentation(template_path)
        template_slide = template_prs.slides[0]
        target_slide = self._get_slide(slide_num)
        target_slide.slide_layout = template_slide.slide_layout
    #endregion

    #region Text Manipulation
    def _set_text_content(self, text_frame, content: str, style: FontStyle) -> None:
        text_frame.text = content
        self._apply_font_style(text_frame, style)

    def add_text_block(self, slide_num: int, content: str, pos: Coordinates,
                    style: FontStyle) -> str:
        slide = self._get_slide(slide_num)
        textbox = slide.shapes.add_textbox(
            Inches(pos.x), Inches(pos.y),
            Inches(pos.width), Inches(pos.height)
        )
        self._set_text_content(textbox.text_frame, content, style)
        return self._set_shape_id(textbox)

    def edit_text_content(self, slide_num: int, shape_id: str, new_text: str) -> None:
        """Modify existing text content"""
        shape = self._find_shape(slide_num, shape_id)
        if shape.shape_type != MSO_SHAPE_TYPE.TEXT_BOX:
            msg = "Shape is not a text box"
            raise ValueError(msg)
        shape.text_frame.text = new_text

    def format_text_style(self, slide_num: int, shape_id: str, style: FontStyle) -> None:
        """Apply text formatting"""
        shape = self._find_shape(slide_num, shape_id)
        self._apply_font_style(shape.text_frame, style)
    #endregion

    #region Image Handling
    def insert_image(self, slide_num: int, image_data: str | bytes,
                    pos: Coordinates) -> str:
        """Add image to slide from file or bytes"""
        slide = self._get_slide(slide_num)

        img_path = self._bytes_to_tempfile(image_data) if isinstance(image_data, bytes) else image_data

        image = slide.shapes.add_picture(
            img_path,
            Inches(pos.x), Inches(pos.y),
            Inches(pos.width), Inches(pos.height)
        )
        return self._set_shape_id(image)

    def replace_image(self, slide_num: int, shape_id: str, new_image: str | bytes) -> None:
        """Replace existing image"""
        old_shape = self._find_shape(slide_num, shape_id)
        pos = Coordinates(
            x=old_shape.left.inches,
            y=old_shape.top.inches,
            width=old_shape.width.inches,
            height=old_shape.height.inches
        )
        self.delete_shape(slide_num, shape_id)
        return self.insert_image(slide_num, new_image, pos)
    #endregion

    #region Charts & Graphs
    def create_chart(self, slide_num: int, chart_type: str, data: SchemaChartData,
                    pos: Coordinates) -> str:
        """Create new chart on slide"""
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

    def modify_chart_data(self, slide_num: int, shape_id: str,
                         new_data: SchemaChartData) -> None:
        """Update chart data"""
        chart = self._find_shape(slide_num, shape_id)
        chart_data = ChartData()
        chart_data.categories = new_data.categories
        for series in new_data.series:
            chart_data.add_series(series.name, series.values)
        chart.chart.replace_data(chart_data)
    #endregion

    #region Tables
    def create_table(self, slide_num: int, rows: int, cols: int,
                    pos: Coordinates) -> str:
        slide = self._get_slide(slide_num)
        table_shape = slide.shapes.add_table(
            rows, cols,
            Inches(pos.x), Inches(pos.y),
            Inches(pos.width), Inches(pos.height)
        )
        return self._set_shape_id(table_shape)

    def edit_table_cell(self, slide_num: int, shape_id: str, cell: TableCell) -> None:
        """Modify table cell content"""
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
    #endregion

    #region Helper Methods
    def _get_slide(self, slide_num: int) -> pptx_Slide:
        """Validate and return slide object"""
        if slide_num < 1 or slide_num > len(self.prs.slides):
            msg = f"Invalid slide number: {slide_num}"
            raise ValueError(msg)
        return self.prs.slides[slide_num-1]

    def _find_shape(self, slide_num: int, shape_id: str) -> Any:
        """Find shape by ID on specified slide"""
        slide = self._get_slide(slide_num)
        for shape in slide.shapes:
            if shape.name == shape_id:
                return shape
        msg = f"Shape {shape_id} not found on slide {slide_num}"
        raise ValueError(msg)

    def _set_shape_id(self, shape) -> str:
        """Assign unique ID to shape"""
        shape_id = str(uuid.uuid4())
        shape.name = shape_id
        return shape_id

    def _apply_font_style(self, text_frame, style: FontStyle) -> None:
        """Apply font formatting to text frame"""
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
        """Save bytes to temporary file"""
        with tempfile.NamedTemporaryFile(delete=False) as f:
            f.write(data)
            self._temp_files.append(f.name)
            return f.name

    def _map_chart_type(self, chart_type: str) -> XL_CHART_TYPE:
        """Convert string to pptx chart type enum"""
        types = {
            'bar': XL_CHART_TYPE.COLUMN_CLUSTERED,
            'line': XL_CHART_TYPE.LINE,
            'pie': XL_CHART_TYPE.PIE,
            'area': XL_CHART_TYPE.AREA
        }
        return types.get(chart_type.lower(), XL_CHART_TYPE.COLUMN_CLUSTERED)
    #endregion

    #region Parsing Methods
    def _parse_slide(self, slide_num: int, slide) -> Slide:
        """Convert pptx slide to Pydantic model"""
        return Slide(
            number=slide_num,
            layout=self._parse_layout(slide.slide_layout),
            background=self._parse_background(slide.background),
            shapes=[self._parse_shape(shape) for shape in slide.shapes],
            design=self.analyze_slide_design(slide_num)
        )

    def _parse_shape(self, shape) -> TextShape | ImageShape | None:
        """Convert pptx shape to schema model"""
        shape_type = shape.shape_type
        common = {
            'id': shape.name,
            'coordinates': Coordinates(
                x=shape.left.inches,
                y=shape.top.inches,
                width=shape.width.inches,
                height=shape.height.inches
            )
        }

        if shape_type == MSO_SHAPE_TYPE.TEXT_BOX:
            return TextShape(
                **common,
                content=TextContent(text=shape.text),
                style=self._parse_font_style(shape.text_frame)
            )
        if shape_type == MSO_SHAPE_TYPE.PICTURE:
            return ImageShape(
                **common,
                path=shape.image.filename
            )
        # Handle other shape types
        return None

    def _parse_font_style(self, text_frame) -> FontStyle:
        """Безопасный парсинг стилей шрифта"""
        try:
            first_run = text_frame.paragraphs[0].runs[0]
            font = first_run.font
            color = "#000000"

            if font.color and hasattr(font.color, 'rgb') and font.color.rgb:
                color = f"#{int(font.color.rgb):06x}"
            elif font.color and hasattr(font.color, 'theme_color'):
                color = f"theme:{font.color.theme_color}"

            return FontStyle(
                name=font.name or "Calibri",
                size=font.size.pt if font.size else 12.0,
                color=color,
                bold=font.bold,
                italic=font.italic,
                underline=font.underline
            )
        except Exception:
            return FontStyle()
    #endregion

    def _get_metadata(self) -> PresentationMeta:
        """Extract presentation metadata"""
        cp = self.prs.core_properties
        return PresentationMeta(
            title=cp.title,
            author=cp.author,
            created=cp.created,
            modified=cp.modified,
            template=self.prs.slide_layouts[0].name if self.prs.slide_layouts else None
        )

    def _get_slide_dimensions(self) -> Coordinates:
        """Get slide dimensions in inches"""
        return Coordinates(
            width=self.prs.slide_width.inches,
            height=self.prs.slide_height.inches
        )

    def _get_master_layouts(self) -> list[SlideLayout]:
        return [
            SlideLayout(
                name=layout.name,
                display_name=layout.name,
                idx=idx
            ) for idx, layout in enumerate(self.prs.slide_layouts)
        ]
    def _get_theme_fonts(self, slide: pptx_Slide) -> dict[str, str]:
        """Extract theme fonts from slide layout"""
        fonts = slide.slide_layout.fonts
        return {
            'latin': fonts.latin.typeface,
            'complex': fonts.complex.typeface,
            'east_asian': fonts.east_asian.typeface
        }

    def _find_layout(self, layout_name: str) -> SlideLayout:
        """Find slide layout by name"""
        for layout in self._get_master_layouts():
            if layout_name in (layout.name, layout.display_name):
                return layout
        msg = f"Layout '{layout_name}' not found"
        raise ValueError(msg)

    def _reorder_slide(self, slide: pptx_Slide, position: int) -> None:
        """Internal method for slide reordering"""
        temp_prs = Presentation()
        # Copy slides up to position
        for i in range(position - 1):
            if i < len(self.prs.slides):
                source = self.prs.slides[i]
                dest = temp_prs.slides.add_slide(source.slide_layout)
                dest.shapes.clone(source.shapes)
        # Insert new slide
        dest = temp_prs.slides.add_slide(slide.slide_layout)
        dest.shapes.clone(slide.shapes)
        # Copy remaining slides
        for i in range(position - 1, len(self.prs.slides) - 1):
            source = self.prs.slides[i]
            dest = temp_prs.slides.add_slide(source.slide_layout)
            dest.shapes.clone(source.shapes)
        self.prs = temp_prs

    def _parse_layout(self, layout: pptx_SlideLayout) -> SlideLayout:
        try:
            idx = self.prs.slide_layouts.index(layout)
        except ValueError:
            idx = -1
        return SlideLayout(
            name=layout.name,
            display_name=layout.name,
            idx=idx
        )

    def _parse_background(self, background: Any) -> Background:
        """Enhanced background parsing with fill type handling"""
        fill = background.fill
        bg_type = "none"
        value = None
        transparency = 1.0

        with contextlib.suppress(Exception):
            if fill.type == FillType.SOLID:
                bg_type = "color"
                color = fill.fore_color
                if color.type == FillType.SOLID:  # RGB
                    value = f"#{color.rgb:06x}"
                elif color.type == FillType.THEME:  # Theme color
                    value = f"theme:{color.theme_color}"
                transparency = fill.transparency

            elif fill.type == FillType.PICTURE:
                bg_type = "image"
                value = fill.picture.url if hasattr(fill.picture, 'url') else "embedded"
                transparency = fill.transparency

            elif fill.type == FillType.PATTERN:
                bg_type = "pattern"
                value = fill.pattern
                transparency = fill.transparency

        return Background(
            type=bg_type,
            value=value,
            transparency=transparency
        )

    def delete_shape(self, slide_num: int, shape_id: str) -> None:
        """Delete shape from slide by ID"""
        shape = self._find_shape(slide_num, shape_id)
        sp = shape.element
        sp.getparent().remove(sp)

    #endregion

    #region Enhanced Shape Parsing
    def _get_shape_parser(self, shape_type) -> callable | None:
        """Return the appropriate parser method based on shape type"""
        parsers = {
            MSO_SHAPE_TYPE.TEXT_BOX: self._parse_text_shape,
            MSO_SHAPE_TYPE.PICTURE: self._parse_image_shape,
            MSO_SHAPE_TYPE.CHART: self._parse_chart_shape,
            MSO_SHAPE_TYPE.TABLE: self._parse_table_shape,
            MSO_SHAPE_TYPE.AUTO_SHAPE: self._parse_geometric_shape
        }
        return parsers.get(shape_type)

    def _parse_shape(self, shape) -> TextShape | ImageShape | ChartShape | TableShape | GeometricShape | None:
        """Convert pptx shape to schema model with full type support"""
        common = {
            'id': shape.name,
            'coordinates': Coordinates(
                x=shape.left.inches,
                y=shape.top.inches,
                width=shape.width.inches,
                height=shape.height.inches
            )
        }

        try:
            parser = self._get_shape_parser(shape.shape_type)
            if parser:
                return parser(shape, **common)
        except Exception as e:
            print(f"Error parsing shape: {e}")
        return None

    def _parse_text_shape(self, shape, **common: dict[str, Any]) -> TextShape:
        """Parse text box shape"""
        return TextShape(
            **common,
            content=TextContent(text=shape.text),
            style=self._parse_font_style(shape.text_frame)
        )

    def _parse_image_shape(self, shape, **common: dict[str, Any]) -> ImageShape:
        """Парсинг изображений с обработкой атрибутов"""
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

    def _parse_chart_shape(self, shape, **common: dict[str, Any]) -> ChartShape:
        """Parse chart shape"""
        chart = shape.chart
        return ChartShape(
            **common,
            chart_type=str(chart.chart_type).split("_")[-1].lower(),
            data=SchemaChartData(
                categories=[c.label for c in chart.categories],
                series=[{"name": s.name, "values": s.values} for s in chart.series]
            ),
            title=chart.chart_title.text if chart.has_title else None
        )

    def _parse_table_shape(self, shape, **common: dict[str, Any]) -> TableShape:
        """Parse table shape"""
        table = shape.table
        return TableShape(
            **common,
            rows=len(table.rows),
            cols=len(table.columns),
            cells=[self._parse_table_cell(cell) for cell in table.iter_cells()]
        )

    def _parse_geometric_shape(self, shape, **common: dict[str, Any]) -> GeometricShape:
        """Parse geometric shape"""
        return GeometricShape(
            **common,
            shape_type=shape.auto_shape_type.name.lower(),
            fill=f"#{shape.fill.fore_color.rgb:06x}",
            outline=f"#{shape.line.color.rgb:06x}",
            outline_width=shape.line.width.inches
        )

    def _parse_image_crop(self, shape) -> Coordinates | None:
        """Parse image cropping parameters"""
        if shape.crop_left or shape.crop_right or shape.crop_top or shape.crop_bottom:
            return Coordinates(
                x=shape.crop_left,
                y=shape.crop_top,
                width=1 - (shape.crop_left + shape.crop_right),
                height=1 - (shape.crop_top + shape.crop_bottom)
            )
        return None

    def _parse_table_cell(self, cell: _Cell) -> TableCell:
        """Parse table cell data"""
        return TableCell(
            row=cell.row_idx,
            col=cell.col_idx,
            content=cell.text,
            span_rows=cell.span_height,
            span_cols=cell.span_width
        )

    def _get_theme_colors(self, slide) -> list[str]:
        """Получение цветов темы с обработкой исключений"""
        try:
            return [str(c) for c in slide.slide_layout.color_scheme.colors]
        except AttributeError:
            return []
