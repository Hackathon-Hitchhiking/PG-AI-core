from collections import defaultdict
from typing import Dict, List, Optional, Tuple, Union, Iterator

import numpy as np
from loguru import logger
from pydantic import BaseModel, validator
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.shapes.autoshape import Shape
from pptx.slide import Slide
from pptx.util import Emu, Pt

class ShapeOpts(BaseModel):
    left: Optional[float] = None
    top: Optional[float] = None
    width: Optional[float] = None
    height: Optional[float] = None
    shape_type: Optional[int] = None
    color: Optional[Tuple[int, int, int]] = None
    line_color: Optional[Tuple[int, int, int]] = None
    line_width: Optional[float] = None
    transparency: Optional[float] = None
    rotation: Optional[float] = None
    text: Optional[str] = None
    text_color: Optional[Tuple[int, int, int]] = None
    text_size: Optional[int] = None
    text_bold: Optional[bool] = None
    text_italic: Optional[bool] = None
    adjustments: Optional[List[int]] = None
    rounding: Optional[float] = None

class GeometricShape(BaseModel):
    shape_id: int
    slide_id: int
    shape_type: int
    left: float
    top: float
    width: float
    height: float
    color: Tuple[int, int, int] = (255, 255, 255)
    line_color: Tuple[int, int, int] = (0, 0, 0)
    line_width: float = 1.0
    transparency: float = 0.0
    rotation: float = 0.0
    text: str = ""
    text_color: Tuple[int, int, int] = (0, 0, 0)
    text_size: int = 12
    text_bold: bool = False
    text_italic: bool = False
    adjustments: List[int] = []
    
    shape_manager: Optional[Shape] = None
    
    class Config:
        arbitrary_types_allowed = True

# Define conversion constants
EMUS_PER_PIXEL = 9525  # Commonly used value for 96 dpi

def px_to_emu(px: float) -> int:
    """
    Преобразует пиксели в EMU (английские метрические единицы).
    
    Аргументы:
        px (float): Значение в пикселях.
        
    Возвращает:
        int: Значение в EMU.
    """
    return int(px * EMUS_PER_PIXEL)

def emu_to_px(emu: int) -> float:
    """
    Преобразует EMU (английские метрические единицы) в пиксели.
    
    Аргументы:
        emu (int): Значение в EMU.
        
    Возвращает:
        float: Значение в пикселях.
    """
    return emu / EMUS_PER_PIXEL

def hex_to_rgb(hex_color):
    """
    Преобразует HEX-код цвета в кортеж RGB.
    
    Аргументы:
        hex_color (str): HEX-код цвета (например, '#FF0000').
        
    Возвращает:
        Tuple[int, int, int]: Кортеж RGB.
    """
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

def rgb_to_tuple(rgb_color):
    """
    Преобразует объект цвета в кортеж (r, g, b).
    Обрабатывает различные типы цветов и возвращает значение по умолчанию при ошибке.
    """
    try:
        if rgb_color is None:
            return (0, 0, 0)
        
        # Если это уже объект RGBColor с атрибутом rgb
        if hasattr(rgb_color, 'rgb') and rgb_color.rgb is not None:
            # Проверяем, является ли rgb_color.rgb числом или другим объектом
            if isinstance(rgb_color.rgb, int):
                rgb_hex = rgb_color.rgb
                r = (rgb_hex >> 16) & 255
                g = (rgb_hex >> 8) & 255
                b = rgb_hex & 255
                return (r, g, b)
            # Если это объект RGBColor с атрибутами r, g, b
            elif hasattr(rgb_color, 'r') and hasattr(rgb_color, 'g') and hasattr(rgb_color, 'b'):
                return (rgb_color.r, rgb_color.g, rgb_color.b)
        
        # Обработка схемных цветов
        if hasattr(rgb_color, 'theme_color'):
            return (128, 128, 128)  # Серый цвет по умолчанию для схемных цветов
    except (AttributeError, TypeError) as e:
        logger.debug(f"Ошибка при преобразовании цвета: {e}")
    
    return (0, 0, 0)  # Черный цвет по умолчанию

class FigureManager:
    """
    Менеджер геометрических фигур для работы с презентациями PowerPoint.
    
    Класс предоставляет функциональность для добавления, редактирования и удаления
    геометрических фигур в презентациях PowerPoint, а также управления их свойствами,
    такими как положение, размер, цвет, прозрачность и скругление.
    
    Все размеры и координаты указываются в пикселях.
    """
    
    def __init__(self):
        """
        Инициализирует новый экземпляр FigureManager.
        
        Создает пустой менеджер фигур без загруженной презентации и с пустым
        словарем фигур.
        """
        self.pres = None
        self.shapes: Dict[int, List[GeometricShape]] = defaultdict(list)
        self.shape_id_counter = defaultdict(int)
        
    def load_presentation(self, source_path: str) -> None:
        """
        Загружает презентацию PowerPoint из файла.
        
        Загружает существующую презентацию по указанному пути и анализирует
        все имеющиеся в ней фигуры, добавляя их в коллекцию фигур менеджера.
        
        Аргументы:
            source_path (str): Путь к файлу презентации PowerPoint (.pptx).
        """
        self.pres = Presentation(source_path)
        self._parse_existing_shapes()
        
    def create_presentation(self) -> None:
        """
        Создает новую пустую презентацию PowerPoint.
        
        Инициализирует новый объект презентации без слайдов и фигур.
        """
        self.pres = Presentation()
        
    def save_presentation(self, target_path: str) -> None:
        """
        Сохраняет презентацию в файл.
        
        Сохраняет текущую презентацию по указанному пути. Если презентация
        не была загружена или создана, вызывает исключение.
        
        Аргументы:
            target_path (str): Путь для сохранения файла презентации.
            
        Вызывает:
            ValueError: Если презентация не загружена или не создана.
        """
        if self.pres:
            self.pres.save(target_path)
        else:
            raise ValueError("Презентация не загружена или не создана")
    
    def add_slide(self) -> int:
        """
        Добавляет новый слайд в презентацию и возвращает его ID.
        
        Создает новый пустой слайд в текущей презентации и возвращает
        его порядковый номер (начиная с 1).
        
        Возвращает:
            int: ID добавленного слайда (порядковый номер, начиная с 1).
            
        Вызывает:
            ValueError: Если презентация не загружена или не создана.
        """
        if not self.pres:
            raise ValueError("Презентация не загружена или не создана")
        
        slide_layout = self.pres.slide_layouts[6]
        slide = self.pres.slides.add_slide(slide_layout)
        return len(self.pres.slides)
    
    def add_shape(self, slide_id: int, shape_type: int, left: float, top: float, 
                 width: float, height: float, **kwargs) -> GeometricShape:
        """
        Добавляет новую фигуру на слайд и возвращает объект фигуры.
        
        Создает новую геометрическую фигуру указанного типа на заданном слайде
        с указанными параметрами положения и размера. Дополнительные параметры
        могут быть переданы через kwargs.
        
        Аргументы:
            slide_id (int): ID слайда, на который добавляется фигура.
            shape_type (int): Тип фигуры (константа из MSO_SHAPE).
            left (float): Позиция левого края фигуры в пикселях.
            top (float): Позиция верхнего края фигуры в пикселях.
            width (float): Ширина фигуры в пикселях.
            height (float): Высота фигуры в пикселях.
            **kwargs: Дополнительные параметры фигуры.
        
        Возвращает:
            GeometricShape: Объект, представляющий добавленную фигуру.
            
        Вызывает:
            ValueError: Если презентация не загружена или не создана, или если указан неверный ID слайда.
        """
        if not self.pres:
            raise ValueError("Презентация не загружена или не создана")
        
        if slide_id <= 0 or slide_id > len(self.pres.slides):
            raise ValueError(f"Неверный ID слайда: {slide_id}")
        
        slide = self.pres.slides[slide_id - 1]
        
        left_emu = Emu(px_to_emu(left))
        top_emu = Emu(px_to_emu(top))
        width_emu = Emu(px_to_emu(width))
        height_emu = Emu(px_to_emu(height))
        
        shape = slide.shapes.add_shape(
            shape_type, left_emu, top_emu, width_emu, height_emu
        )
        
        self.shape_id_counter[slide_id] += 1
        shape_id = self.shape_id_counter[slide_id]
        
        geometric_shape = GeometricShape(
            shape_id=shape_id,
            slide_id=slide_id,
            shape_type=shape_type,
            left=left,
            top=top,
            width=width,
            height=height,
            shape_manager=shape,
            **kwargs
        )
        
        self._apply_shape_properties(geometric_shape, **kwargs)
        
        self.shapes[slide_id].append(geometric_shape)
        
        return geometric_shape
    
    def update_shape(self, slide_id: int, shape_id: int, opts: Union[ShapeOpts, Dict]) -> None:
        """
        Обновляет свойства существующей фигуры.
        
        Изменяет указанные свойства фигуры на заданном слайде. Свойства, 
        которые не указаны в opts, остаются без изменений.
        
        Аргументы:
            slide_id (int): ID слайда, содержащего фигуру.
            shape_id (int): ID фигуры для обновления.
            opts (Union[ShapeOpts, Dict]): Объект или словарь с новыми свойствами фигуры.
                Может содержать следующие поля:
                - left (float): Новая позиция левого края в пикселях.
                - top (float): Новая позиция верхнего края в пикселях.
                - width (float): Новая ширина в пикселях.
                - height (float): Новая высота в пикселях.
                - color (Tuple[int, int, int]): Новый цвет заливки (RGB).
                - line_color (Tuple[int, int, int]): Новый цвет контура (RGB).
                - line_width (float): Новая ширина линии контура.
                - transparency (float): Новое значение прозрачности (0.0-1.0).
                - rotation (float): Новый угол поворота в градусах.
                - adjustments (List[int]): Новые значения настроек формы фигуры.
                - rounding (float): Новое значение скругления углов (0.0-1.0).
                
        Вызывает:
            ValueError: Если фигура с указанным ID не найдена на слайде.
        """
        if isinstance(opts, dict):
            opts = ShapeOpts(**opts)
        
        shape = self._get_shape_by_id(slide_id, shape_id)
        if not shape:
            raise ValueError(f"Фигура с ID {shape_id} не найдена на слайде {slide_id}")
        
        if opts.left is not None or opts.top is not None or opts.width is not None or opts.height is not None:
            self._update_shape_position(shape, opts)
        
        if opts.color is not None:
            self._update_shape_fill_color(shape, opts.color)
        
        if opts.line_color is not None:
            self._update_shape_line_color(shape, opts.line_color)
        
        if opts.line_width is not None:
            self._update_shape_line_width(shape, opts.line_width)
        
        if opts.transparency is not None:
            self._update_shape_transparency(shape, opts.transparency)
        
        if opts.rotation is not None:
            self._update_shape_rotation(shape, opts.rotation)
        
        if opts.adjustments is not None:
            self._update_shape_adjustments(shape, opts.adjustments)
            
        if opts.rounding is not None:
            self.set_shape_rounding(slide_id, shape_id, opts.rounding)
    
    def update_shape_position(self, slide_id: int, shape_id: int, left: float = None, top: float = None, 
                             width: float = None, height: float = None) -> None:
        """
        Обновляет положение и размер фигуры.
        
        Изменяет положение и/или размер фигуры на указанные значения.
        Параметры, которые не указаны (None), остаются без изменений.
        
        Аргументы:
            slide_id (int): ID слайда, содержащего фигуру.
            shape_id (int): ID фигуры для обновления.
            left (float, optional): Новая позиция левого края в пикселях.
            top (float, optional): Новая позиция верхнего края в пикселях.
            width (float, optional): Новая ширина в пикселях.
            height (float, optional): Новая высота в пикселях.
            
        Вызывает:
            ValueError: Если фигура с указанным ID не найдена на слайде.
        """
        shape = self._get_shape_by_id(slide_id, shape_id)
        if not shape:
            raise ValueError(f"Фигура с ID {shape_id} не найдена на слайде {slide_id}")
        
        ppt_shape = shape.shape_manager
        
        if left is not None:
            ppt_shape.left = Emu(px_to_emu(left))
            shape.left = left
        
        if top is not None:
            ppt_shape.top = Emu(px_to_emu(top))
            shape.top = top
        
        if width is not None:
            ppt_shape.width = Emu(px_to_emu(width))
            shape.width = width
        
        if height is not None:
            ppt_shape.height = Emu(px_to_emu(height))
            shape.height = height
    
    def update_shape_color(self, slide_id: int, shape_id: int, color: Tuple[int, int, int]) -> None:
        """
        Обновляет цвет заливки фигуры.
        
        Изменяет цвет заливки фигуры на указанный цвет RGB.
        
        Аргументы:
            slide_id (int): ID слайда, содержащего фигуру.
            shape_id (int): ID фигуры для обновления.
            color (Tuple[int, int, int]): Новый цвет заливки (RGB).
            
        Вызывает:
            ValueError: Если фигура с указанным ID не найдена на слайде.
        """
        shape = self._get_shape_by_id(slide_id, shape_id)
        if not shape:
            raise ValueError(f"Фигура с ID {shape_id} не найдена на слайде {slide_id}")
        
        self._update_shape_fill_color(shape, color)
    
    def update_shape_line(self, slide_id: int, shape_id: int, color: Tuple[int, int, int] = None, 
                         width: float = None) -> None:
        """
        Обновляет свойства линии контура фигуры.
        
        Изменяет цвет и/или ширину линии контура фигуры.
        Параметры, которые не указаны (None), остаются без изменений.
        
        Аргументы:
            slide_id (int): ID слайда, содержащего фигуру.
            shape_id (int): ID фигуры для обновления.
            color (Tuple[int, int, int], optional): Новый цвет линии (RGB).
            width (float, optional): Новая ширина линии.
            
        Вызывает:
            ValueError: Если фигура с указанным ID не найдена на слайде.
        """
        shape = self._get_shape_by_id(slide_id, shape_id)
        if not shape:
            raise ValueError(f"Фигура с ID {shape_id} не найдена на слайде {slide_id}")
        
        if color is not None:
            self._update_shape_line_color(shape, color)
        
        if width is not None:
            self._update_shape_line_width(shape, width)
    
    def update_shape_transparency(self, slide_id: int, shape_id: int, transparency: float) -> None:
        """
        Обновляет прозрачность фигуры.
        
        Изменяет прозрачность заливки фигуры на указанное значение.
        
        Аргументы:
            slide_id (int): ID слайда, содержащего фигуру.
            shape_id (int): ID фигуры для обновления.
            transparency (float): Новое значение прозрачности (0.0-1.0).
            
        Вызывает:
            ValueError: Если фигура с указанным ID не найдена на слайде или значение прозрачности вне диапазона.
        """
        if not 0.0 <= transparency <= 1.0:
            raise ValueError("Значение прозрачности должно быть в диапазоне от 0.0 до 1.0")
        
        shape = self._get_shape_by_id(slide_id, shape_id)
        if not shape:
            raise ValueError(f"Фигура с ID {shape_id} не найдена на слайде {slide_id}")
        
        self._update_shape_transparency(shape, transparency)
    
    def update_shape_rotation(self, slide_id: int, shape_id: int, rotation: float) -> None:
        """
        Обновляет поворот фигуры.
        
        Изменяет угол поворота фигуры на указанное значение.
        
        Аргументы:
            slide_id (int): ID слайда, содержащего фигуру.
            shape_id (int): ID фигуры для обновления.
            rotation (float): Новый угол поворота в градусах.
            
        Вызывает:
            ValueError: Если фигура с указанным ID не найдена на слайде.
        """
        shape = self._get_shape_by_id(slide_id, shape_id)
        if not shape:
            raise ValueError(f"Фигура с ID {shape_id} не найдена на слайде {slide_id}")
        
        self._update_shape_rotation(shape, rotation)
    
    def set_shape_rounding(self, slide_id: int, shape_id: int, rounding_value: float) -> None:
        """
        Sets the rounding value for a shape.
        
        Sets the corner rounding radius for shapes that support this feature.
        Works with ROUNDED_RECTANGLE shapes and other shapes that have
        adjustment values. The rounding value is typically in the range
        from 0 to 1, where 0 means no rounding and 1 means maximum rounding.
        
        If the shape doesn't support rounding, the operation will be skipped.
        
        Args:
            slide_id (int): ID of the slide containing the shape.
            shape_id (int): ID of the shape to modify.
            rounding_value (float): Rounding value (from 0.0 to 1.0).
                
        Raises:
            ValueError: If the shape with the specified ID is not found on the slide or
                    the rounding value is outside the valid range.
        """
        if not 0.0 <= rounding_value <= 1.0:
            raise ValueError("Rounding value must be between 0.0 and 1.0")
        
        shape = self._get_shape_by_id(slide_id, shape_id)
        if not shape:
            raise ValueError(f"Shape with ID {shape_id} not found on slide {slide_id}")
        
        ppt_shape = shape.shape_manager
        
        # Calculate the adjustment value for PowerPoint (scale from 0-1 to 0-100000)
        adj_value = int(rounding_value * 100000)
        
        try:
            # Check if the shape has adjustment values and can be rounded
            if hasattr(ppt_shape, 'adjustment_values') and len(ppt_shape.adjustment_values) > 0:
                ppt_shape.adjustment_values[0] = adj_value
                
                # Update our internal representation
                if not shape.adjustments:
                    shape.adjustments = [adj_value]
                else:
                    shape.adjustments[0] = adj_value
                
                logger.info(f"Successfully set rounding {rounding_value} for shape with ID {shape_id}")
                return
            
            # Alternative way to access adjustments
            elif hasattr(ppt_shape, 'adjustments') and ppt_shape.adjustments:
                ppt_shape.adjustments[0] = adj_value
                
                # Update our internal representation
                if not shape.adjustments:
                    shape.adjustments = [adj_value]
                else:
                    shape.adjustments[0] = adj_value
                
                logger.info(f"Successfully set rounding {rounding_value} for shape with ID {shape_id}")
                return
            
            # If we get here, the shape doesn't support rounding
            shape_type_name = "unknown"
            if hasattr(ppt_shape, 'auto_shape_type'):
                shape_type_name = f"{ppt_shape.auto_shape_type} ({ppt_shape.auto_shape_type.value})"
            
            logger.info(f"Shape type {shape_type_name} does not support rounding. Operation skipped.")
        
        except IndexError:
            logger.info(f"Shape does not have adjustment settings for rounding (no element with index 0). Operation skipped.")
        except Exception as e:
            logger.warning(f"Failed to set rounding: {str(e)}")

    
    def delete_shape(self, slide_id: int, shape_id: int) -> None:
        """
        Удаляет фигуру со слайда.
        
        Удаляет указанную фигуру с заданного слайда как из презентации PowerPoint,
        так и из внутренней коллекции фигур менеджера.
        
        Аргументы:
            slide_id (int): ID слайда, содержащего фигуру.
            shape_id (int): ID фигуры для удаления.
            
        Вызывает:
            ValueError: Если фигура с указанным ID не найдена на слайде.
        """
        shape = self._get_shape_by_id(slide_id, shape_id)
        if not shape:
            raise ValueError(f"Фигура с ID {shape_id} не найдена на слайде {slide_id}")
        
        sp = shape.shape_manager._element
        sp.getparent().remove(sp)
        
        shapes_list = self.shapes[slide_id]
        for i, s in enumerate(shapes_list):
            if s.shape_id == shape_id:
                del shapes_list[i]
                break
    
    def get_shape_json(self, slide_id: int, shape_id: int = None) -> Union[Dict, List[Dict]]:
        """
        Получает JSON-представление фигуры или всех фигур на слайде.
        
        Возвращает словарь с данными указанной фигуры или список словарей
        со всеми фигурами на слайде, если shape_id не указан.
        
        Аргументы:
            slide_id (int): ID слайда.
            shape_id (int, optional): ID фигуры. По умолчанию None.
        
        Возвращает:
            Union[Dict, List[Dict]]: Словарь с данными фигуры или список словарей с данными всех фигур.
        """
        if shape_id is not None:
            shape = self._get_shape_by_id(slide_id, shape_id)
            if not shape:
                return {}
            return shape.dict(exclude={'shape_manager'})
        else:
            return [s.dict(exclude={'shape_manager'}) for s in self.shapes[slide_id]]
    
    def get_all_shapes_json(self) -> Dict[int, List[Dict]]:
        """
        Получает JSON-представление всех фигур в презентации.
        
        Возвращает словарь, где ключи - ID слайдов, а значения - списки словарей
        с данными всех фигур на соответствующих слайдах.
        
        Возвращает:
            Dict[int, List[Dict]]: Словарь с данными всех фигур в презентации.
        """
        result = {}
        for slide_id, shapes_list in self.shapes.items():
            result[slide_id] = [s.dict(exclude={'shape_manager'}) for s in shapes_list]
        return result
    
    def get_shape_ids(self, slide_id: int) -> List[int]:
        """
        Получает список всех ID фигур на указанном слайде.
        
        Аргументы:
            slide_id (int): ID слайда.
            
        Возвращает:
            List[int]: Список ID фигур на слайде.
        """
        return [shape.shape_id for shape in self.shapes[slide_id]]
    
    def find_shapes(self, slide_id: int, **criteria) -> List[GeometricShape]:
        """
        Находит фигуры на слайде, соответствующие указанным критериям.
        
        Аргументы:
            slide_id (int): ID слайда.
            **criteria: Критерии поиска в виде пар ключ-значение.
                Например: text="Заголовок", color=(255, 0, 0)
            
        Возвращает:
            List[GeometricShape]: Список найденных фигур.
        """
        result = []
        
        for shape in self.shapes[slide_id]:
            match = True
            for key, value in criteria.items():
                if not hasattr(shape, key) or getattr(shape, key) != value:
                    match = False
                    break
            
            if match:
                result.append(shape)
        
        return result
    
    
    def _capture_shape_properties(self, ppt_shape) -> dict:
        """
        Сохраняет все свойства фигуры в словарь.
        
        Аргументы:
            ppt_shape: Объект фигуры python-pptx.
            
        Возвращает:
            dict: Словарь со всеми свойствами фигуры.
        """
        properties = {
            'left': ppt_shape.left,
            'top': ppt_shape.top,
            'width': ppt_shape.width,
            'height': ppt_shape.height,
            'rotation': ppt_shape.rotation if hasattr(ppt_shape, 'rotation') else 0,
        }
        
        # Сохраняем свойства заливки
        if hasattr(ppt_shape, 'fill'):
            fill_props = {}
            if hasattr(ppt_shape.fill, 'fore_color') and ppt_shape.fill.fore_color:
                fill_props['fore_color'] = ppt_shape.fill.fore_color.rgb
            if hasattr(ppt_shape.fill, 'back_color') and ppt_shape.fill.back_color:
                fill_props['back_color'] = ppt_shape.fill.back_color.rgb
            if hasattr(ppt_shape.fill, 'transparency'):
                fill_props['transparency'] = ppt_shape.fill.transparency
            if hasattr(ppt_shape.fill, 'type'):
                fill_props['type'] = ppt_shape.fill.type
            properties['fill'] = fill_props
        
        # Сохраняем свойства линии
        if hasattr(ppt_shape, 'line'):
            line_props = {}
            if hasattr(ppt_shape.line, 'color') and ppt_shape.line.color:
                line_props['color'] = ppt_shape.line.color.rgb
            if hasattr(ppt_shape.line, 'width'):
                line_props['width'] = ppt_shape.line.width
            if hasattr(ppt_shape.line, 'dash_style'):
                line_props['dash_style'] = ppt_shape.line.dash_style
            properties['line'] = line_props
        
        # Сохраняем текстовые свойства
        if hasattr(ppt_shape, 'text') and ppt_shape.text:
            properties['text'] = ppt_shape.text
            if hasattr(ppt_shape, 'text_frame') and ppt_shape.text_frame.paragraphs:
                text_props = []
                for p in ppt_shape.text_frame.paragraphs:
                    p_props = {
                        'text': p.text,
                        'alignment': p.alignment if hasattr(p, 'alignment') else None,
                        'runs': []
                    }
                    for run in p.runs:
                        run_props = {
                            'text': run.text
                        }
                        if hasattr(run, 'font'):
                            font_props = {}
                            if hasattr(run.font, 'color') and run.font.color:
                                font_props['color'] = run.font.color.rgb
                            if hasattr(run.font, 'size'):
                                font_props['size'] = run.font.size
                            if hasattr(run.font, 'bold'):
                                font_props['bold'] = run.font.bold
                            if hasattr(run.font, 'italic'):
                                font_props['italic'] = run.font.italic
                            if hasattr(run.font, 'underline'):
                                font_props['underline'] = run.font.underline
                            run_props['font'] = font_props
                        p_props['runs'].append(run_props)
                    text_props.append(p_props)
                properties['text_frame'] = text_props
        
        return properties
    
    def _apply_shape_properties_from_dict(self, ppt_shape, properties: dict) -> None:
        """
        Применяет сохраненные свойства к фигуре.
        
        Аргументы:
            ppt_shape: Объект фигуры python-pptx.
            properties (dict): Словарь со свойствами фигуры.
        """
        # Применяем базовые свойства
        if 'rotation' in properties and hasattr(ppt_shape, 'rotation'):
            ppt_shape.rotation = properties['rotation']
        
        # Применяем свойства заливки
        if 'fill' in properties and hasattr(ppt_shape, 'fill'):
            fill_props = properties['fill']
            if 'type' in fill_props and fill_props['type'] != 0:
                ppt_shape.fill.solid()
                if 'fore_color' in fill_props and hasattr(ppt_shape.fill, 'fore_color'):
                    ppt_shape.fill.fore_color.rgb = fill_props['fore_color']
                if 'transparency' in fill_props and hasattr(ppt_shape.fill, 'transparency'):
                    ppt_shape.fill.transparency = fill_props['transparency']
        
        if 'line' in properties and hasattr(ppt_shape, 'line'):
            line_props = properties['line']
            if 'color' in line_props and hasattr(ppt_shape.line, 'color'):
                ppt_shape.line.color.rgb = line_props['color']
            if 'width' in line_props and hasattr(ppt_shape.line, 'width'):
                ppt_shape.line.width = line_props['width']
            if 'dash_style' in line_props and hasattr(ppt_shape.line, 'dash_style'):
                ppt_shape.line.dash_style = line_props['dash_style']
        
        if 'text' in properties and hasattr(ppt_shape, 'text_frame'):
            if 'text_frame' in properties:
                ppt_shape.text_frame.clear()
                for p_idx, p_props in enumerate(properties['text_frame']):
                    if p_idx == 0:
                        p = ppt_shape.text_frame.paragraphs[0]
                    else:
                        p = ppt_shape.text_frame.add_paragraph()
                    
                    if 'alignment' in p_props and p_props['alignment'] is not None:
                        p.alignment = p_props['alignment']
                    
                    for run_props in p_props['runs']:
                        run = p.add_run()
                        run.text = run_props['text']
                        
                        if 'font' in run_props:
                            font_props = run_props['font']
                            if 'color' in font_props and font_props['color'] is not None:
                                run.font.color.rgb = font_props['color']
                            if 'size' in font_props and font_props['size'] is not None:
                                run.font.size = font_props['size']
                            if 'bold' in font_props and font_props['bold'] is not None:
                                run.font.bold = font_props['bold']
                            if 'italic' in font_props and font_props['italic'] is not None:
                                run.font.italic = font_props['italic']
                            if 'underline' in font_props and font_props['underline'] is not None:
                                run.font.underline = font_props['underline']
            else:
                ppt_shape.text = properties['text']

    def _get_slide_by_id(self, slide_id: int) -> Optional[Slide]:
        """
        Получает объект слайда по его ID.
        
        Аргументы:
            slide_id (int): ID слайда.
            
        Возвращает:
            Optional[Slide]: Объект слайда или None, если слайд не найден.
        """
        if not self.pres:
            return None
        
        if slide_id <= 0 or slide_id > len(self.pres.slides):
            return None
        
        return self.pres.slides[slide_id - 1]
    
    def _parse_existing_shapes(self) -> None:
        """
        Анализирует существующие фигуры из загруженной презентации.
        
        Проходит по всем слайдам и фигурам в загруженной презентации,
        создает для них объекты GeometricShape и добавляет их в коллекцию фигур менеджера.
        """
        if not self.pres:
            return
        
        for slide_idx, slide in enumerate(self.pres.slides, 1):
            shape_idx = 1
            for shape in slide.shapes:
                if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
                    self._parse_shape(slide_idx, shape_idx, shape)
                    shape_idx += 1
                    self.shape_id_counter[slide_idx] = max(self.shape_id_counter[slide_idx], shape_idx)
    
    def _parse_shape(self, slide_id: int, shape_id: int, shape: Shape) -> Optional[GeometricShape]:
        """
        Анализирует фигуру из презентации и добавляет ее в коллекцию.
        
        Извлекает все свойства фигуры, создает объект GeometricShape
        и добавляет его в коллекцию фигур менеджера.
        
        Аргументы:
            slide_id (int): ID слайда, содержащего фигуру.
            shape_id (int): ID фигуры.
            shape (Shape): Объект фигуры из python-pptx.
            
        Возвращает:
            Optional[GeometricShape]: Созданный объект фигуры или None в случае ошибки.
        """
        try:
            left = emu_to_px(shape.left)
            top = emu_to_px(shape.top)
            width = emu_to_px(shape.width)
            height = emu_to_px(shape.height)
            
            # Получение цвета заливки с обработкой ошибок
            fill_color = (255, 255, 255)  # Белый цвет по умолчанию
            try:
                if hasattr(shape.fill, 'fore_color') and hasattr(shape.fill, 'type') and shape.fill.type != 0:
                    fill_color = rgb_to_tuple(shape.fill.fore_color)
            except Exception as e:
                pass
                #logger.debug(f"Не удалось получить цвет заливки: {e}")
            
            # Получение цвета линии с обработкой ошибок
            line_color = (0, 0, 0)  # Черный цвет по умолчанию
            try:
                if hasattr(shape.line, 'color') and shape.line.color is not None:
                    line_color = rgb_to_tuple(shape.line.color)
            except Exception as e:
                logger.debug(f"Не удалось получить цвет линии: {e}")
            
            # Получение ширины линии
            line_width = 1.0
            try:
                if hasattr(shape.line, 'width') and shape.line.width is not None:
                    line_width = emu_to_px(shape.line.width)
            except Exception as e:
                logger.debug(f"Не удалось получить ширину линии: {e}")
            
            # Инициализация текстовых свойств
            text = ""
            text_color = (0, 0, 0)
            text_size = 12
            text_bold = False
            text_italic = False
            
            # Получение текста и его форматирования
            try:
                if hasattr(shape, 'text') and shape.text:
                    text = shape.text
                    if hasattr(shape, 'text_frame') and shape.text_frame.paragraphs:
                        p = shape.text_frame.paragraphs[0]
                        if p.runs:
                            run = p.runs[0]
                            try:
                                if hasattr(run.font, 'color') and run.font.color is not None:
                                    text_color = rgb_to_tuple(run.font.color)
                            except Exception as e:
                                logger.debug(f"Не удалось получить цвет текста: {e}")
                            
                            try:
                                if hasattr(run.font, 'size') and run.font.size:
                                    text_size = run.font.size.pt
                            except Exception as e:
                                logger.debug(f"Не удалось получить размер текста: {e}")

                            text_bold = False if run.font.bold is None else run.font.bold
                            text_italic = False if run.font.italic is None else run.font.italic
            except Exception as e:
                logger.debug(f"Не удалось получить свойства текста: {e}")
            
            adjustments = []
            try:
                if hasattr(shape, 'adjustment_values') and shape.adjustment_values:
                    adjustments = list(shape.adjustment_values)
            except Exception as e:
                logger.debug(f"Не удалось получить настройки формы: {e}")

            geometric_shape = GeometricShape(
                shape_id=shape_id,
                slide_id=slide_id,
                shape_type=shape.auto_shape_type,
                left=left,
                top=top,
                width=width,
                height=height,
                color=fill_color,
                line_color=line_color,
                line_width=line_width,
                text=text,
                text_color=text_color,
                text_size=text_size,
                text_bold=text_bold,
                text_italic=text_italic,
                adjustments=adjustments,
                shape_manager=shape
            )
            
            self.shapes[slide_id].append(geometric_shape)
            return geometric_shape
            
        except Exception as e:
            logger.warning(f"Ошибка при анализе фигуры: {e}")
            return None
    
    def _get_shape_by_id(self, slide_id: int, shape_id: int) -> Optional[GeometricShape]:
        """
        Получает фигуру по её ID с определенного слайда.
        
        Ищет фигуру с указанным ID на заданном слайде в коллекции фигур.
        В отличие от метода _get_shape, этот метод использует правильный
        подход к поиску фигуры по её shape_id, а не по индексу в коллекции.
        
        Аргументы:
            slide_id (int): ID слайда.
            shape_id (int): ID фигуры.
            
        Возвращает:
            Optional[GeometricShape]: Найденная фигура или None, если фигура не найдена.
        """
        for shape in self.shapes[slide_id]:
            if shape.shape_id == shape_id:
                return shape
        return None
    
    def _get_shape(self, slide_id: int, shape_id: int) -> Optional[GeometricShape]:
        """
        Получает фигуру по ее ID с определенного слайда.
        
        Ищет фигуру с указанным ID на заданном слайде в коллекции фигур менеджера.
        
        Аргументы:
            slide_id (int): ID слайда.
            shape_id (int): ID фигуры.
            
        Возвращает:
            Optional[GeometricShape]: Найденная фигура или None, если фигура не найдена.
        """
        return self._get_shape_by_id(slide_id, shape_id)
    
    def _get_shapes(self, slide_id: int, shape_id: Optional[int] = None) -> Iterator[GeometricShape]:
        """
        Получает фигуры с определенного слайда, опционально фильтруя по ID фигуры.
        
        Возвращает итератор по фигурам на заданном слайде, опционально фильтруя
        только фигуру с указанным ID.
        
        Аргументы:
            slide_id (int): ID слайда.
            shape_id (Optional[int], optional): ID фигуры для фильтрации. По умолчанию None.
            
        Возвращает:
            Iterator[GeometricShape]: Итератор по фигурам.
        """
        if slide_id < 0:
            slide_id = sorted(list(self.shapes.keys()))[slide_id]
        
        for shape in self.shapes[slide_id]:
            if shape_id is not None and shape.shape_id != shape_id:
                continue
            yield shape
    
    def _apply_shape_properties(self, shape: GeometricShape, **kwargs) -> None:
        """
        Применяет свойства к новой созданной фигуре.
        
        Устанавливает различные свойства фигуры, такие как цвет заливки, цвет линии,
        ширина линии, прозрачность, поворот, текст и его форматирование, а также
        настройки формы.
        
        Аргументы:
            shape (GeometricShape): Объект фигуры.
            **kwargs: Словарь свойств для применения.
        """
        ppt_shape = shape.shape_manager
        
        if 'color' in kwargs:
            color = kwargs['color']
            try:
                ppt_shape.fill.solid()
                ppt_shape.fill.fore_color.rgb = RGBColor(color[0], color[1], color[2])
                shape.color = color
            except Exception as e:
                logger.warning(f"Не удалось установить цвет заливки: {e}")
        
        if 'line_color' in kwargs:
            line_color = kwargs['line_color']
            try:
                ppt_shape.line.color.rgb = RGBColor(line_color[0], line_color[1], line_color[2])
                shape.line_color = line_color
            except Exception as e:
                logger.warning(f"Не удалось установить цвет линии: {e}")
        
        if 'line_width' in kwargs:
            line_width = kwargs['line_width']
            try:
                ppt_shape.line.width = Emu(px_to_emu(line_width))
                shape.line_width = line_width
            except Exception as e:
                logger.warning(f"Не удалось установить ширину линии: {e}")
        
        if 'transparency' in kwargs:
            transparency = kwargs['transparency']
            try:
                if hasattr(ppt_shape.fill, 'transparency'):
                    ppt_shape.fill.transparency = transparency
                shape.transparency = transparency
            except Exception as e:
                logger.warning(f"Не удалось установить прозрачность: {e}")
        
        if 'rotation' in kwargs:
            rotation = kwargs['rotation']
            try:
                if hasattr(ppt_shape, 'rotation'):
                    ppt_shape.rotation = rotation
                shape.rotation = rotation
            except Exception as e:
                logger.warning(f"Не удалось установить поворот: {e}")
        
        if 'text' in kwargs:
            text = kwargs['text']
            try:
                if hasattr(ppt_shape, 'text_frame'):
                    ppt_shape.text_frame.clear()
                    p = ppt_shape.text_frame.paragraphs[0]
                    run = p.add_run()
                    run.text = text
                    
                    if 'text_color' in kwargs:
                        text_color = kwargs['text_color']
                        run.font.color.rgb = RGBColor(text_color[0], text_color[1], text_color[2])
                        shape.text_color = text_color
                    
                    if 'text_size' in kwargs:
                        text_size = kwargs['text_size']
                        run.font.size = Pt(text_size)
                        shape.text_size = text_size
                    
                    if 'text_bold' in kwargs:
                        text_bold = kwargs['text_bold']
                        run.font.bold = text_bold
                        shape.text_bold = text_bold
                    
                    if 'text_italic' in kwargs:
                        text_italic = kwargs['text_italic']
                        run.font.italic = text_italic
                        shape.text_italic = text_italic
                
                shape.text = text
            except Exception as e:
                logger.warning(f"Не удалось установить текст: {e}")
        
        if 'adjustments' in kwargs:
            adjustments = kwargs['adjustments']
            try:
                if hasattr(ppt_shape, 'adjustment_values'):
                    for i, value in enumerate(adjustments):
                        if i < len(ppt_shape.adjustment_values):
                            ppt_shape.adjustment_values[i] = value
                shape.adjustments = adjustments
            except Exception as e:
                logger.warning(f"Не удалось установить настройки формы: {e}")
                
        if 'rounding' in kwargs:
            rounding = kwargs['rounding']
            try:
                if hasattr(ppt_shape, 'adjustment_values') and len(ppt_shape.adjustment_values) > 0:
                    adj_value = int(rounding * 100000)
                    ppt_shape.adjustment_values[0] = adj_value
                    
                    if not shape.adjustments:
                        shape.adjustments = [adj_value]
                    else:
                        shape.adjustments[0] = adj_value
            except Exception as e:
                logger.warning(f"Не удалось установить скругление: {e}")
    
    def _update_shape_position(self, shape: GeometricShape, opts: ShapeOpts) -> None:
        """
        Обновляет положение и размер фигуры.
        
        Изменяет положение и/или размер фигуры в соответствии с указанными параметрами.
        
        Аргументы:
            shape (GeometricShape): Объект фигуры.
            opts (ShapeOpts): Объект с новыми параметрами положения и размера.
        """
        ppt_shape = shape.shape_manager
        
        if opts.left is not None:
            ppt_shape.left = Emu(px_to_emu(opts.left))
            shape.left = opts.left
        
        if opts.top is not None:
            ppt_shape.top = Emu(px_to_emu(opts.top))
            shape.top = opts.top
        
        if opts.width is not None:
            ppt_shape.width = Emu(px_to_emu(opts.width))
            shape.width = opts.width
        
        if opts.height is not None:
            ppt_shape.height = Emu(px_to_emu(opts.height))
            shape.height = opts.height
    
    def _update_shape_fill_color(self, shape: GeometricShape, color: Tuple[int, int, int]) -> None:
        """
        Обновляет цвет заливки фигуры.
        
        Изменяет цвет заливки фигуры на указанный цвет RGB.
        
        Аргументы:
            shape (GeometricShape): Объект фигуры.
            color (Tuple[int, int, int]): Новый цвет заливки (RGB).
        """
        ppt_shape = shape.shape_manager
        try:
            ppt_shape.fill.solid()
            ppt_shape.fill.fore_color.rgb = RGBColor(color[0], color[1], color[2])
            shape.color = color
        except Exception as e:
            logger.warning(f"Не удалось обновить цвет заливки: {e}")
    
    def _update_shape_line_color(self, shape: GeometricShape, color: Tuple[int, int, int]) -> None:
        """
        Обновляет цвет линии фигуры.
        
        Изменяет цвет контура фигуры на указанный цвет RGB.
        
        Аргументы:
            shape (GeometricShape): Объект фигуры.
            color (Tuple[int, int, int]): Новый цвет линии (RGB).
        """
        ppt_shape = shape.shape_manager
        try:
            ppt_shape.line.color.rgb = RGBColor(color[0], color[1], color[2])
            shape.line_color = color
        except Exception as e:
            logger.warning(f"Не удалось обновить цвет линии: {e}")
    
    def _update_shape_line_width(self, shape: GeometricShape, width: float) -> None:
        """
        Обновляет ширину линии фигуры.
        
        Изменяет ширину контура фигуры на указанное значение.
        
        Аргументы:
            shape (GeometricShape): Объект фигуры.
            width (float): Новая ширина линии в пикселях.
        """
        ppt_shape = shape.shape_manager
        try:
            ppt_shape.line.width = Emu(px_to_emu(width))
            shape.line_width = width
        except Exception as e:
            logger.warning(f"Не удалось обновить ширину линии: {e}")
    
    def _update_shape_transparency(self, shape: GeometricShape, transparency: float) -> None:
        """
        Обновляет прозрачность фигуры.
        
        Изменяет прозрачность заливки фигуры на указанное значение.
        
        Аргументы:
            shape (GeometricShape): Объект фигуры.
            transparency (float): Новое значение прозрачности (0.0-1.0).
        """
        ppt_shape = shape.shape_manager
        try:
            if hasattr(ppt_shape.fill, 'transparency'):
                ppt_shape.fill.transparency = transparency
            shape.transparency = transparency
        except Exception as e:
            logger.warning(f"Не удалось обновить прозрачность: {e}")
    
    def _update_shape_rotation(self, shape: GeometricShape, rotation: float) -> None:
        """
        Обновляет поворот фигуры.
        
        Изменяет угол поворота фигуры на указанное значение.
        
        Аргументы:
            shape (GeometricShape): Объект фигуры.
            rotation (float): Новый угол поворота в градусах.
        """
        ppt_shape = shape.shape_manager
        try:
            if hasattr(ppt_shape, 'rotation'):
                ppt_shape.rotation = rotation
            shape.rotation = rotation
        except Exception as e:
            logger.warning(f"Не удалось обновить поворот: {e}")
    
    def _update_shape_text(self, shape: GeometricShape, text: str) -> None:
        """
        Обновляет текст фигуры.
        
        Изменяет текстовое содержимое фигуры, сохраняя существующее форматирование текста.
        
        Аргументы:
            shape (GeometricShape): Объект фигуры.
            text (str): Новый текст.
        """
        ppt_shape = shape.shape_manager
        try:
            if hasattr(ppt_shape, 'text_frame'):
                ppt_shape.text_frame.clear()
                p = ppt_shape.text_frame.paragraphs[0]
                run = p.add_run()
                run.text = text
                
                run.font.color.rgb = RGBColor(shape.text_color[0], shape.text_color[1], shape.text_color[2])
                run.font.size = Pt(shape.text_size)
                run.font.bold = shape.text_bold
                run.font.italic = shape.text_italic
            
            shape.text = text
        except Exception as e:
            logger.warning(f"Не удалось обновить текст: {e}")
    
    def _update_shape_text_color(self, shape: GeometricShape, color: Tuple[int, int, int]) -> None:
        """
        Обновляет цвет текста фигуры.
        
        Изменяет цвет текста фигуры на указанный цвет RGB.
        
        Аргументы:
            shape (GeometricShape): Объект фигуры.
            color (Tuple[int, int, int]): Новый цвет текста (RGB).
        """
        ppt_shape = shape.shape_manager
        try:
            if hasattr(ppt_shape, 'text_frame') and ppt_shape.text_frame.paragraphs:
                for p in ppt_shape.text_frame.paragraphs:
                    for run in p.runs:
                        run.font.color.rgb = RGBColor(color[0], color[1], color[2])
            
            shape.text_color = color
        except Exception as e:
            logger.warning(f"Не удалось обновить цвет текста: {e}")
    
    def _update_shape_text_size(self, shape: GeometricShape, size: int) -> None:
        """
        Обновляет размер текста фигуры.
        
        Изменяет размер шрифта текста фигуры на указанное значение.
        
        Аргументы:
            shape (GeometricShape): Объект фигуры.
            size (int): Новый размер шрифта.
        """
        ppt_shape = shape.shape_manager
        try:
            if hasattr(ppt_shape, 'text_frame') and ppt_shape.text_frame.paragraphs:
                for p in ppt_shape.text_frame.paragraphs:
                    for run in p.runs:
                        run.font.size = Pt(size)
            
            shape.text_size = size
        except Exception as e:
            logger.warning(f"Не удалось обновить размер текста: {e}")
    
    def _update_shape_text_bold(self, shape: GeometricShape, bold: bool) -> None:
        """
        Обновляет жирность текста фигуры.
        
        Устанавливает или снимает жирное начертание текста фигуры.
        
        Аргументы:
            shape (GeometricShape): Объект фигуры.
            bold (bool): Новое значение жирности текста.
        """
        ppt_shape = shape.shape_manager
        try:
            if hasattr(ppt_shape, 'text_frame') and ppt_shape.text_frame.paragraphs:
                for p in ppt_shape.text_frame.paragraphs:
                    for run in p.runs:
                        run.font.bold = bold
            
            shape.text_bold = bold
        except Exception as e:
            logger.warning(f"Не удалось обновить жирность текста: {e}")
    
    def _update_shape_text_italic(self, shape: GeometricShape, italic: bool) -> None:
        """
        Обновляет курсив текста фигуры.
        
        Устанавливает или снимает курсивное начертание текста фигуры.
        
        Аргументы:
            shape (GeometricShape): Объект фигуры.
            italic (bool): Новое значение курсива текста.
        """
        ppt_shape = shape.shape_manager
        try:
            if hasattr(ppt_shape, 'text_frame') and ppt_shape.text_frame.paragraphs:
                for p in ppt_shape.text_frame.paragraphs:
                    for run in p.runs:
                        run.font.italic = italic
            
            shape.text_italic = italic
        except Exception as e:
            logger.warning(f"Не удалось обновить курсив текста: {e}")
    
    def _update_shape_adjustments(self, shape: GeometricShape, adjustments: List[int]) -> None:
        """
        Обновляет значения настроек формы фигуры.
        
        Изменяет значения настроек формы фигуры, которые влияют на ее геометрию.
        
        Аргументы:
            shape (GeometricShape): Объект фигуры.
            adjustments (List[int]): Список новых значений настроек.
        """
        ppt_shape = shape.shape_manager
        try:
            if hasattr(ppt_shape, 'adjustment_values'):
                for i, value in enumerate(adjustments):
                    if i < len(ppt_shape.adjustment_values):
                        ppt_shape.adjustment_values[i] = value
            
            shape.adjustments = adjustments
        except Exception as e:
            logger.warning(f"Не удалось обновить настройки формы: {e}")

    def test(self, source_path: str, target_path: str = "test_output.pptx"):
        """
        Тестовая функция для демонстрации возможностей FigureManager.
        
        Загружает презентацию, добавляет новый слайд с различными фигурами,
        демонстрирует возможности по созданию, редактированию и форматированию фигур,
        а затем сохраняет результат в новый файл.
        
        Аргументы:
            source_path (str): Путь к исходному файлу презентации.
            target_path (str, optional): Путь для сохранения результата. По умолчанию "test_output.pptx".
        """
        self.load_presentation(source_path)
        
        
        rect1 = self.add_shape(
            3,
            MSO_SHAPE.RECTANGLE, 
            200, 200, 500, 500,
            color=(255, 0, 0),
            line_color=(0, 0, 0),
            line_width=2.0,
            rounding=0.1
        )
        ## ДЛЯ КИРИЛЛА [update_shape_color, update_shape_position, update_shape_transparency, set_shape_rounding]
        
        #self.update_shape_position(3, 1, rounding = 0.5, color=(255, 0, 255))
        #self.update_shape_color(3, 1, color=(255, 255, 0))
        #self.update_shape_line(new_slide_id, rect3.shape_id, color=(255, 0, 0), width=3.0)
        #self.update_shape_transparency(3, 1, transparency=1)
        #self.update_shape_transparency(3, rect1.shape_id, transparency=0.2)
        #self.update_shape_rotation(2, 1, rotation=45.0)
        #self.update_shape_rotation(2, 2, rotation=45.0)
        #self.update_shape_rotation(2, 3, rotation=45.0)
        #self.update_shape_rotation(2, 4, rotation=45.0)

        #self.set_shape_rounding(3, 2, 0.5)
        #self.set_shape_rounding(3, 1, 1)
        #self.set_shape_rounding(3, 3, 0.1)

        self.save_presentation(target_path)


if __name__ == "__main__":
    manager = FigureManager()
    manager.test("test_data/test_dit.pptx")
