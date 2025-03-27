from pptx.shapes.base import BaseShape
from pptx.slide import Slide


def get_all_methods(cls):
    return [method for method in dir(cls) if not method.startswith('__')]


def hex_to_rgb(hex: str) -> tuple[int, ...]:
    hex_code = hex.lstrip('#')

    return tuple(int(hex_code[i : i + 2], 16) for i in (0, 2, 4))


def get_slide_from_shape(shape: BaseShape) -> Slide:
    for i in range(100):
        shape = shape._parent
        if isinstance(shape, Slide):
            return shape
    raise ValueError('shape is not in a slide')
